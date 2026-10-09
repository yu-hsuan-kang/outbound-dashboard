# -*- coding: utf-8 -*-
"""Sigenergy Daily Outbound Report Generator V3.3

Robust daily version:
- Reads the original shared XLSX with openpyxl first; falls back to direct XML on style errors.
- Date priority: YYYYMMDD in input filename > explicit --date > computer's current local date.
- Never guesses the report date from the maximum historical order date.
- Validates that the input filename/date and data are not obviously inconsistent.
- Generates a customer-facing email in the structure of the supplied manual daily report.
- Generates HTML email, TXT copy, Excel report, FTL exception CSV, and run metadata/log.
"""
import argparse, hashlib, json, logging, os, re, sys, traceback, zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
import pandas as pd

try:
    import openpyxl
except Exception:
    openpyxl = None

APP_VERSION = "V3.4 KPI Dashboard + Customer Email"
WAREHOUSE_SHEETS = {
    "AWL Moerdijk 出库": "AWL Moerdijk Outbound",
    "AWL Energy  出库": "AWL Energy Outbound",
    "Botlek出库": "Botlek Outbound",
}
NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
EXCEL_EPOCH=pd.Timestamp("1899-12-30")

def norm(v):
    if v is None or pd.isna(v): return ""
    return re.sub(r"\s+","",str(v)).strip().lower()

def clean(v):
    return "" if v is None or pd.isna(v) else str(v).strip()

def pick(df,cands):
    exact={norm(c):c for c in df.columns}
    for c in cands:
        if norm(c) in exact:return exact[norm(c)]
    ns=[norm(c) for c in cands]
    for col in df.columns:
        n=norm(col)
        if any(c and (c in n or n in c) for c in ns):return col
    return None

def sha256(path):
    h=hashlib.sha256()
    with open(path,"rb") as f:
        for b in iter(lambda:f.read(1024*1024),b""):h.update(b)
    return h.hexdigest()

def colnum(s):
    n=0
    for c in s:n=n*26+ord(c)-64
    return n

def shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():return []
    root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter('{%s}t'%NS['m'])) for si in root.findall('m:si',NS)]

def sheet_paths(z):
    wb=ET.fromstring(z.read('xl/workbook.xml')); rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rels={x.attrib['Id']:x.attrib['Target'] for x in rel}; out={}
    for s in wb.find('m:sheets',NS):
        target=rels[s.attrib['{%s}id'%NS['r']]]
        out[s.attrib['name']]=target.lstrip('/') if target.startswith('/') else (target if target.startswith('xl/') else 'xl/'+target)
    return out

def read_xml_sheet(z,path,ss):
    root=ET.fromstring(z.read(path)); rows=[]
    for row in root.findall('.//m:sheetData/m:row',NS):
        vals={}
        for c in row.findall('m:c',NS):
            ref=c.attrib.get('r',''); m=re.match(r'([A-Z]+)',ref)
            if not m:continue
            v=c.find('m:v',NS); val='' if v is None else v.text
            typ=c.attrib.get('t')
            if typ=='s' and val!='':
                try:val=ss[int(val)]
                except:val=''
            elif typ=='inlineStr':val=''.join(t.text or '' for t in c.iter('{%s}t'%NS['m']))
            elif typ=='b':val='TRUE' if val=='1' else 'FALSE'
            vals[m.group(1)]='' if val is None else val
        if vals:rows.append(vals)
    if not rows:return pd.DataFrame()
    keys=sorted(rows[0],key=colnum); headers=[rows[0].get(k) or k for k in keys]
    seen={}; cols=[]
    for h in headers:
        h=str(h).strip() or 'Unnamed'; seen[h]=seen.get(h,0)+1; cols.append(h if seen[h]==1 else f'{h}__{seen[h]}')
    df=pd.DataFrame([[r.get(k,'') for k in keys] for r in rows[1:]],columns=cols)
    keep=[c for c in df.columns if df[c].astype(str).str.strip().ne('').any()]
    return df.loc[:,keep]

def load_xml(path):
    with zipfile.ZipFile(path) as z:
        if z.testzip():raise ValueError('XLSX 压缩包损坏')
        paths=sheet_paths(z); ss=shared_strings(z)
        missing=[x for x in WAREHOUSE_SHEETS if x not in paths]
        if missing:raise ValueError('缺少仓库 Sheet: '+','.join(missing))
        return {s:read_xml_sheet(z,paths[s],ss) for s in WAREHOUSE_SHEETS}

def load_openpyxl(path):
    if openpyxl is None:raise RuntimeError('openpyxl 未安装')
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    try:
        if any(s not in wb.sheetnames for s in WAREHOUSE_SHEETS):raise ValueError('缺少仓库 Sheet')
        out={}
        for s in WAREHOUSE_SHEETS:
            vals=list(wb[s].values); out[s]=pd.DataFrame(vals[1:],columns=[clean(x) for x in vals[0]]) if vals else pd.DataFrame()
        return out
    finally:wb.close()

def load_raw(path,logger):
    try:
        x=load_openpyxl(path); logger.info('Input reader: openpyxl'); return x,'openpyxl'
    except Exception as e:
        logger.warning('openpyxl 读取失败，自动切换 XML reader: %s',e)
        x=load_xml(path); logger.info('Input reader: direct XLSX XML fallback'); return x,'xml'

def parse_date(v):
    if v is None or pd.isna(v) or str(v).strip()=='':return pd.NaT
    if isinstance(v,(pd.Timestamp,datetime)):return pd.Timestamp(v).normalize()
    s=str(v).strip()
    try:
        n=float(s)
        if 20000<=n<=80000:return (EXCEL_EPOCH+pd.to_timedelta(n,unit='D')).normalize()
    except:pass
    d=pd.to_datetime(s,errors='coerce')
    return pd.NaT if pd.isna(d) else pd.Timestamp(d).normalize()

def dser(s):return s.map(parse_date)
def nser(s):return pd.to_numeric(s,errors='coerce').fillna(0)
def next_workday(d):
    d=pd.Timestamp(d).normalize()+pd.Timedelta(days=1)
    while d.weekday()>=5:d+=pd.Timedelta(days=1)
    return d

def normalize(df,sheet):
    order=pick(df,['DN',' DN','order']); status=pick(df,['Order status','Order Status']); release=pick(df,['Release Date'])
    pcd=pick(df,['PCD','Picking Complete Date(APD)','Picking Complete Date']); planned=pick(df,['Outbound Planned Date','OutboundPlannedDate']); actual=pick(df,['Outbound Date'])
    pallet=pick(df,['Pallet']); qty=pick(df,['QTY','Quantity']); adr=pick(df,['ADR']); eta=pick(df,['ETA']); lsp=pick(df,['LSP']); track=pick(df,['Tracking/AWB','Tracking/AWB ']); detail=pick(df,['Status']); delivery=pick(df,['Delivery Method'])
    if not order:raise ValueError(f'{sheet} 找不到 DN/order')
    for name,c in [('Release Date',release),('PCD',pcd),('Outbound Planned Date',planned),('Outbound Date',actual)]:
        if not c:raise ValueError(f'{sheet} 缺少关键字段: {name}')
    def st(v):
        n=norm(v)
        if n in ('cancelled','canceled'):return 'Cancelled'
        if 'sigerequiredonhold' in n:return 'Sige Required On Hold'
        if 'waitingoc' in n:return 'Waiting OC'
        return clean(v)
    z=pd.DataFrame({'Warehouse':WAREHOUSE_SHEETS[sheet],'Order':df[order].map(clean),'Order status':df[status].map(st) if status else '',
      'Release Date':dser(df[release]),'PCD':dser(df[pcd]),'Planned Outbound':dser(df[planned]),'Outbound Date':dser(df[actual]),
      'Pallet':nser(df[pallet]) if pallet else 0,'QTY':nser(df[qty]) if qty else 0,'ADR':df[adr].map(clean) if adr else '',
      'ETA':dser(df[eta]) if eta else pd.NaT,'LSP':df[lsp].map(clean) if lsp else '','Tracking/AWB':df[track].map(clean) if track else '',
      'Status detail':df[detail].map(clean) if detail else '','Delivery Method':df[delivery].map(clean) if delivery else ''})
    return z[z.Order.ne('')].copy()

def determine_date(input_path, explicit):
    # User-friendly default: the date embedded in the input filename is authoritative.
    # --date is only a fallback when the filename does not contain YYYYMMDD.
    m=re.search(r'(20\d{2})(0\d|1[0-2])([0-3]\d)',os.path.basename(input_path))
    if m:
        try:
            return pd.Timestamp(datetime(int(m.group(1)),int(m.group(2)),int(m.group(3)))).normalize(),'input filename YYYYMMDD'
        except Exception:
            pass
    if explicit:
        return pd.Timestamp(explicit).normalize(),'explicit --date fallback'
    return pd.Timestamp.now().normalize(),'computer local date fallback'

def fmt_date(v):
    if pd.isna(v):return ''
    return pd.Timestamp(v).strftime('%Y/%m/%d').replace('/0','/')

def summary(df):return {'orders':int(len(df)),'pallet':int(df.Pallet.sum()),'qty':int(df.QTY.sum())}

def setup_log(folder):
    os.makedirs(folder,exist_ok=True); log=logging.getLogger('sigenergy_v31'); log.handlers.clear(); log.setLevel(logging.INFO); log.propagate=False
    fh=logging.FileHandler(os.path.join(folder,'run.log'),encoding='utf-8'); fh.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')); log.addHandler(fh)
    sh=logging.StreamHandler(sys.stdout); sh.setFormatter(logging.Formatter('%(levelname)s | %(message)s')); log.addHandler(sh); return log

def make_report(input_path,output_root=None,explicit_date=None):
    input_path=os.path.abspath(input_path)
    root=os.path.abspath(output_root or os.path.join(os.path.dirname(input_path),'Daily_Report_Output'))
    prelog=setup_log(root); prelog.info('=== %s ===',APP_VERSION); prelog.info('Input: %s',input_path); prelog.info('SHA256: %s',sha256(input_path))
    raw,reader=load_raw(input_path,prelog)
    all_df=pd.concat([normalize(raw[s],s) for s in WAREHOUSE_SHEETS],ignore_index=True)
    if all_df.empty:raise ValueError('没有有效订单数据')
    report_date,date_source=determine_date(input_path,explicit_date); nxt=next_workday(report_date)
    # Sanity check: warn if file-name date and system date differ, but do not silently switch date.
    system_date=pd.Timestamp.now().normalize()
    if date_source.startswith('input filename') and report_date!=system_date:
        prelog.warning('输入文件日期 %s 与电脑当前日期 %s 不同；按文件名日期生成日报。',report_date.date(),system_date.date())
    outdir=os.path.join(root,report_date.strftime('%Y%m%d')); log=setup_log(outdir)
    log.info('Reader=%s | ReportDate=%s | DateSource=%s | NextWorkday=%s',reader,report_date.date(),date_source,nxt.date())
    eq=lambda c,d:all_df[c].eq(pd.Timestamp(d))
    today=all_df[eq('Outbound Date',report_date)].copy(); new=all_df[eq('Release Date',report_date)].copy(); nxtdf=all_df[eq('Planned Outbound',nxt)].copy()
    exc_base=all_df[eq('Planned Outbound',report_date)&all_df['Outbound Date'].isna()].copy()
    exc=exc_base[exc_base['Delivery Method'].str.upper().str.strip().eq('FTL')].copy()
    status=[]
    for wh in WAREHOUSE_SHEETS.values():
        r=all_df[(all_df.Warehouse==wh)&eq('Release Date',report_date)]; p=all_df[(all_df.Warehouse==wh)&eq('PCD',report_date)]
        status.append({'Warehouse':wh,'Total Orders':len(r),'Packing Completed':len(p),'Cancelled':int((r['Order status']=='Cancelled').sum()),'Sige Required On Hold':int((r['Order status']=='Sige Required On Hold').sum()),'Waiting OC':int((r['Order status']=='Waiting OC').sum()),'Shortage/Abnormal Stock':int(r['Status detail'].str.contains('shortage|abnormal',case=False,na=False).sum())})
    status_df=pd.DataFrame(status); waiting=int((new['Order status']=='Waiting OC').sum()); shortage=int(new['Status detail'].str.contains('shortage|abnormal',case=False,na=False).sum())
    today_disp=today[['Order','Pallet','QTY','ADR','ETA','Warehouse']].copy(); today_disp['ETA']=today_disp.ETA.map(fmt_date); today_disp['WH']=today_disp.pop('Warehouse').map({'AWL Moerdijk Outbound':'MDK','AWL Energy Outbound':'Energy','Botlek Outbound':'Botlek'})
    exc_disp=exc[['Order','LSP','Planned Outbound','Tracking/AWB']].copy(); exc_disp['Planned Outbound']=exc_disp['Planned Outbound'].map(fmt_date)
    next_disp=nxtdf[['Order','Pallet','QTY','ADR','ETA','Warehouse']].copy(); next_disp['ETA']=next_disp.ETA.map(fmt_date); next_disp['WH']=next_disp.pop('Warehouse').map({'AWL Moerdijk Outbound':'MDK','AWL Energy Outbound':'Energy','Botlek Outbound':'Botlek'})
    sm={'today_outbound':summary(today),'new_orders':summary(new),'next_workday':summary(nxtdf),'exceptions':summary(exc)}
    base=f'Sigenergy_Daily_Outbound_Report_{report_date:%Y%m%d}'
    xlsx=os.path.join(outdir,base+'.xlsx'); email_html=os.path.join(outdir,base+'_EMAIL.html'); dashboard_html=os.path.join(outdir,base+'_DASHBOARD.html'); txt=os.path.join(outdir,base+'_email.txt'); csv=os.path.join(outdir,base+'_FTL_exceptions.csv'); meta=os.path.join(outdir,base+'_run_summary.json')
    with pd.ExcelWriter(xlsx,engine='openpyxl') as w:
        status_df.to_excel(w,index=False,sheet_name='Warehouse Status'); today_disp.to_excel(w,index=False,sheet_name='日清日结'); exc_disp.to_excel(w,index=False,sheet_name='当日出库异常'); next_disp.to_excel(w,index=False,sheet_name='Next Workday'); new.to_excel(w,index=False,sheet_name='New Orders'); all_df.to_excel(w,index=False,sheet_name='Normalized Data')
        for ws in w.book.worksheets:
            ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=min(45,max(10,max(len(str(c.value or '')) for c in col)+2))
    def table(df):
        if df.empty:return '<table class="mail"><tr><td>None</td></tr></table>'
        return df.to_html(index=False,border=0,classes='mail',na_rep='')
    status_mail=status_df.rename(columns={'Warehouse':'Warehouse','Total Orders':'Release Orders','Packing Completed':'Packing Completed','Cancelled':'Cancelled','Sige Required On Hold':'Sige Required On Hold','Waiting OC':'Waiting OC','Shortage/Abnormal Stock':'Shortage/Abnormal Stock'})
    email_html_doc=f"""<!doctype html><html><head><meta charset="utf-8"><style>body{{font-family:Arial,"Microsoft YaHei",sans-serif;color:#222;font-size:13px}}p{{margin:8px 0}}h3{{margin:18px 0 8px}}table.mail{{border-collapse:collapse;margin:6px 0 14px}}table.mail th,table.mail td{{border:1px solid #999;padding:5px 8px;text-align:center;white-space:nowrap}}table.mail th{{font-weight:700;background:#f2f2f2}}table.mail td:first-child{{text-align:left}}</style></head><body>
<p>目前等待处理（Waiting OC）状态订单<b>{waiting}</b>单，<b>{shortage}</b>单缺库存</p>
<h3>Warehouse Status</h3>{table(status_mail)}
<p><b>今日出库日清日结 {sm['today_outbound']['orders']} 單，共计 {sm['today_outbound']['pallet']} plts, {sm['today_outbound']['qty']} pcs</b></p>{table(today_disp.rename(columns={'Order':'order','Pallet':'Pallet','QTY':'QTY','ADR':'ADR','ETA':'ETA','WH':'WH'}))}
<p><b>今日整车出庫異常訂單共 {sm['exceptions']['orders']} 單</b></p>{table(exc_disp.rename(columns={'Order':'DN','LSP':'LSP','Planned Outbound':'OutboundPlannedDate','Tracking/AWB':'Tracking/AWB'}))}
<p><b>下个工作日三仓出库預計 {sm['next_workday']['orders']} 单, {sm['next_workday']['pallet']} 托，{sm['next_workday']['qty']} pcs ，具体出库订单可在腾讯在线文档中查看</b></p>{table(next_disp.rename(columns={'Order':'order','Pallet':'Pallet','QTY':'QTY','ADR':'ADR','ETA':'ETA','WH':'WH'}))}
</body></html>"""
    open(email_html,'w',encoding='utf-8').write(email_html_doc)
    def dash_table(df):
        if df.empty:return '<p class="empty">None</p>'
        return df.to_html(index=False,border=0,classes='tbl',na_rep='')
    dashboard_html_doc=f"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Sigenergy Daily Outbound Dashboard — {report_date:%Y-%m-%d}</title><style>
body{{font-family:Arial,"Microsoft YaHei",sans-serif;background:#f5f7fa;color:#172033;margin:0;padding:28px}}.wrap{{max-width:1250px;margin:auto}}h1{{font-size:26px;margin:0 0 6px}}h2{{font-size:18px;margin:0 0 12px}}.note{{font-size:13px;color:#667085;margin-bottom:18px}}.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}.card{{background:#fff;border:1px solid #e2e6ed;border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(16,24,40,.04)}}.k{{font-size:12px;color:#667085}}.v{{font-size:30px;font-weight:700;margin:6px 0}}section{{background:#fff;border:1px solid #e2e6ed;border-radius:12px;padding:18px;margin-top:16px;overflow:auto;box-shadow:0 2px 8px rgba(16,24,40,.03)}}table{{border-collapse:collapse;width:100%;font-size:13px}}th,td{{padding:8px 10px;border-bottom:1px solid #edf0f4;text-align:left;white-space:nowrap}}th{{background:#f8fafc;font-weight:700}}tr:hover td{{background:#fafbfc}}.empty{{color:#667085}}@media(max-width:900px){{.cards{{grid-template-columns:repeat(2,1fr)}}}}@media(max-width:560px){{body{{padding:14px}}.cards{{grid-template-columns:1fr}}}}
</style></head><body><div class="wrap"><h1>Sigenergy Daily Outbound Report — {report_date:%Y-%m-%d}</h1><div class="note">Source: original Sigenergy shared outbound workbook · {APP_VERSION} · Reader: {reader} · Date source: {date_source}</div>
<div class="cards"><div class="card"><div class="k">New Orders</div><div class="v">{sm['new_orders']['orders']}</div><div class="k">{sm['new_orders']['pallet']} pallets / {sm['new_orders']['qty']} pcs</div></div><div class="card"><div class="k">Today Outbound</div><div class="v">{sm['today_outbound']['orders']}</div><div class="k">{sm['today_outbound']['pallet']} pallets / {sm['today_outbound']['qty']} pcs</div></div><div class="card"><div class="k">Next Workday · {nxt:%Y-%m-%d}</div><div class="v">{sm['next_workday']['orders']}</div><div class="k">{sm['next_workday']['pallet']} pallets / {sm['next_workday']['qty']} pcs</div></div><div class="card"><div class="k">FTL Exceptions</div><div class="v">{sm['exceptions']['orders']}</div><div class="k">{sm['exceptions']['pallet']} pallets / {sm['exceptions']['qty']} pcs</div></div></div>
<section><h2>Warehouse Status</h2>{dash_table(status_df)}</section><section><h2>Today's Outbound</h2>{dash_table(today_disp)}</section><section><h2>Today's FTL Outbound Exceptions</h2>{dash_table(exc_disp)}</section><section><h2>Next Workday Plan</h2>{dash_table(next_disp)}</section></div></body></html>"""
    open(dashboard_html,'w',encoding='utf-8').write(dashboard_html_doc)
    exc_disp.to_csv(csv,index=False,encoding='utf-8-sig')
    email='\n'.join([f'目前等待处理（Waiting OC）状态订单{waiting}单，{shortage}单缺库存','', 'Warehouse Status', status_mail.to_string(index=False), '',f"今日出库日清日结 {sm['today_outbound']['orders']} 單，共计 {sm['today_outbound']['pallet']} plts, {sm['today_outbound']['qty']} pcs",today_disp.to_string(index=False),'',f"今日整车出庫異常訂單共 {sm['exceptions']['orders']} 單",exc_disp.to_string(index=False),'',f"下个工作日三仓出库預計 {sm['next_workday']['orders']} 单, {sm['next_workday']['pallet']} 托，{sm['next_workday']['qty']} pcs ，具体出库订单可在腾讯在线文档中查看"])
    open(txt,'w',encoding='utf-8').write(email)
    info={'version':APP_VERSION,'input':input_path,'input_sha256':sha256(input_path),'reader':reader,'report_date':str(report_date.date()),'date_source':date_source,'computer_date':str(system_date.date()),'next_workday':str(nxt.date()),'summary':sm,'outputs':outdir}
    json.dump(info,open(meta,'w',encoding='utf-8'),ensure_ascii=False,indent=2,default=str)
    log.info('Completed: %s',outdir)
    return {'outdir':outdir,'xlsx':xlsx,'email_html':email_html,'dashboard_html':dashboard_html,'txt':txt,'csv':csv,'meta':meta,'log':os.path.join(outdir,'run.log'),'date':report_date,'date_source':date_source,'reader':reader,'status':status_df,'summary':sm}

def choose_input_file():
    """Open a Windows file picker for interactive use."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askopenfilename(
            title='请选择 Sigenergy Outbound 原始 Excel 文件',
            filetypes=[('Excel files', '*.xlsx;*.xlsm;*.xltx;*.xltm'), ('All files', '*.*')]
        )
        root.destroy()
        return path
    except Exception as e:
        raise RuntimeError(f'无法打开文件选择窗口：{e}')

def main():
    ap=argparse.ArgumentParser(description='Sigenergy Daily Outbound Report Generator V3.4')
    ap.add_argument('input',nargs='?',help='原始 Excel 文件；省略时弹出文件选择窗口')
    ap.add_argument('output',nargs='?',help='可选输出目录')
    ap.add_argument('--date',help='YYYY-MM-DD；仅当文件名没有 YYYYMMDD 时使用')
    a=ap.parse_args()
    try:
        input_path = a.input or choose_input_file()
        if not input_path:
            print('已取消：未选择输入文件。')
            return
        if not os.path.isfile(input_path):
            raise FileNotFoundError(f'找不到输入文件：{input_path}')
        r=make_report(input_path,a.output,a.date)
        print(json.dumps({k:v for k,v in r.items() if k not in ('status','date')},ensure_ascii=False,default=str,indent=2))
        print(r['status'].to_string(index=False))
        print(f"\n日报生成完成：{r['outdir']}")
    except Exception as e:
        print('ERROR | 日报生成失败:',e,file=sys.stderr); traceback.print_exc(); sys.exit(1)
if __name__=='__main__':main()
