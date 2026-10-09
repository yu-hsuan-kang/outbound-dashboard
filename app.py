# -*- coding: utf-8 -*-
"""
Sigenergy Daily Outbound Report - Streamlit Web Version
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import streamlit as st

try:
    import openpyxl
except Exception:
    openpyxl = None

# --- 页面配置 ---
st.set_page_config(
    page_title="Sigenergy Daily Outbound Dashboard",
    page_icon="📦",
    layout="wide"
)

APP_VERSION = "V3.4 Web Edition"
WAREHOUSE_SHEETS = {
    "AWL Moerdijk 出库": "AWL Moerdijk Outbound",
    "AWL Energy  出库": "AWL Energy Outbound",
    "Botlek出库": "Botlek Outbound",
}
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}
EXCEL_EPOCH = pd.Timestamp("1899-12-30")

# --- 核心数据处理函数 (保留原逻辑) ---
def norm(v):
    if v is None or pd.isna(v): return ""
    return re.sub(r"\s+", "", str(v)).strip().lower()

def clean(v):
    return "" if v is None or pd.isna(v) else str(v).strip()

def pick(df, cands):
    exact = {norm(c): c for c in df.columns}
    for c in cands:
        if norm(c) in exact: return exact[norm(c)]
    ns = [norm(c) for c in cands]
    for col in df.columns:
        n = norm(col)
        if any(c and (c in n or n in c) for c in ns): return col
    return None

def colnum(s):
    n = 0
    for c in s: n = n * 26 + ord(c) - 64
    return n

def shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iter('{%s}t' % NS['m'])) for si in root.findall('m:si', NS)]

def sheet_paths(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rel = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rels = {x.attrib['Id']: x.attrib['Target'] for x in rel}
    out = {}
    for s in wb.find('m:sheets', NS):
        target = rels[s.attrib['{%s}id' % NS['r']]]
        out[s.attrib['name']] = target.lstrip('/') if target.startswith('/') else (target if target.startswith('xl/') else 'xl/' + target)
    return out

def read_xml_sheet(z, path, ss):
    root = ET.fromstring(z.read(path))
    rows = []
    for row in root.findall('.//m:sheetData/m:row', NS):
        vals = {}
        for c in row.findall('m:c', NS):
            ref = c.attrib.get('r', '')
            m = re.match(r'([A-Z]+)', ref)
            if not m: continue
            v = c.find('m:v', NS)
            val = '' if v is None else v.text
            typ = c.attrib.get('t')
            if typ == 's' and val != '':
                try: val = ss[int(val)]
                except: val = ''
            elif typ == 'inlineStr': val = ''.join(t.text or '' for t in c.iter('{%s}t' % NS['m']))
            elif typ == 'b': val = 'TRUE' if val == '1' else 'FALSE'
            vals[m.group(1)] = '' if val is None else val
        if vals: rows.append(vals)
    if not rows: return pd.DataFrame()
    keys = sorted(rows[0], key=colnum)
    headers = [rows[0].get(k) or k for k in keys]
    seen = {}
    cols = []
    for h in headers:
        h = str(h).strip() or 'Unnamed'
        seen[h] = seen.get(h, 0) + 1
        cols.append(h if seen[h] == 1 else f'{h}__{seen[h]}')
    df = pd.DataFrame([[r.get(k, '') for k in keys] for r in rows[1:]], columns=cols)
    keep = [c for c in df.columns if df[c].astype(str).str.strip().ne('').any()]
    return df.loc[:, keep]

def load_xml_bytes(file_bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        if z.testzip(): raise ValueError('XLSX 压缩包损坏')
        paths = sheet_paths(z)
        ss = shared_strings(z)
        # 寻找匹配的 Sheet
        found_sheets = {}
        for target_sheet in WAREHOUSE_SHEETS:
            # 允许忽略多余空格匹配 Sheet 名称
            matched = None
            for p in paths:
                if norm(p) == norm(target_sheet):
                    matched = p
                    break
            if matched:
                found_sheets[target_sheet] = read_xml_sheet(z, paths[matched], ss)
            else:
                raise ValueError(f'缺少仓库 Sheet: {target_sheet}')
        return found_sheets

def load_openpyxl_bytes(file_bytes):
    if openpyxl is None: raise RuntimeError('openpyxl 未安装')
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    try:
        out = {}
        sheet_map = {norm(s): s for s in wb.sheetnames}
        for s in WAREHOUSE_SHEETS:
            if norm(s) not in sheet_map:
                raise ValueError(f'缺少仓库 Sheet: {s}')
            actual_sheet_name = sheet_map[norm(s)]
            vals = list(wb[actual_sheet_name].values)
            out[s] = pd.DataFrame(vals[1:], columns=[clean(x) for x in vals[0]]) if vals else pd.DataFrame()
        return out
    finally:
        wb.close()

def parse_date(v):
    if v is None or pd.isna(v) or str(v).strip() == '': return pd.NaT
    if isinstance(v, (pd.Timestamp, datetime)): return pd.Timestamp(v).normalize()
    s = str(v).strip()
    try:
        n = float(s)
        if 20000 <= n <= 80000: return (EXCEL_EPOCH + pd.to_timedelta(n, unit='D')).normalize()
    except: pass
    d = pd.to_datetime(s, errors='coerce')
    return pd.NaT if pd.isna(d) else pd.Timestamp(d).normalize()

def dser(s): return s.map(parse_date)
def nser(s): return pd.to_numeric(s, errors='coerce').fillna(0)
def next_workday(d):
    d = pd.Timestamp(d).normalize() + pd.Timedelta(days=1)
    while d.weekday() >= 5: d += pd.Timedelta(days=1)
    return d

def normalize(df, sheet):
    order = pick(df, ['DN', ' DN', 'order'])
    status = pick(df, ['Order status', 'Order Status'])
    release = pick(df, ['Release Date'])
    pcd = pick(df, ['PCD', 'Picking Complete Date(APD)', 'Picking Complete Date'])
    planned = pick(df, ['Outbound Planned Date', 'OutboundPlannedDate'])
    actual = pick(df, ['Outbound Date'])
    pallet = pick(df, ['Pallet'])
    qty = pick(df, ['QTY', 'Quantity'])
    adr = pick(df, ['ADR'])
    eta = pick(df, ['ETA'])
    lsp = pick(df, ['LSP'])
    track = pick(df, ['Tracking/AWB', 'Tracking/AWB '])
    detail = pick(df, ['Status'])
    delivery = pick(df, ['Delivery Method'])
    
    if not order: raise ValueError(f'{sheet} 找不到 DN/order')
    for name, c in [('Release Date', release), ('PCD', pcd), ('Outbound Planned Date', planned), ('Outbound Date', actual)]:
        if not c: raise ValueError(f'{sheet} 缺少关键字段: {name}')
    
    def st(v):
        n = norm(v)
        if n in ('cancelled', 'canceled'): return 'Cancelled'
        if 'sigerequiredonhold' in n: return 'Sige Required On Hold'
        if 'waitingoc' in n: return 'Waiting OC'
        return clean(v)

    z = pd.DataFrame({
        'Warehouse': WAREHOUSE_SHEETS[sheet],
        'Order': df[order].map(clean),
        'Order status': df[status].map(st) if status else '',
        'Release Date': dser(df[release]),
        'PCD': dser(df[pcd]),
        'Planned Outbound': dser(df[planned]),
        'Outbound Date': dser(df[actual]),
        'Pallet': nser(df[pallet]) if pallet else 0,
        'QTY': nser(df[qty]) if qty else 0,
        'ADR': df[adr].map(clean) if adr else '',
        'ETA': dser(df[eta]) if eta else pd.NaT,
        'LSP': df[lsp].map(clean) if lsp else '',
        'Tracking/AWB': df[track].map(clean) if track else '',
        'Status detail': df[detail].map(clean) if detail else '',
        'Delivery Method': df[delivery].map(clean) if delivery else ''
    })
    return z[z.Order.ne('')].copy()

def fmt_date(v):
    if pd.isna(v): return ''
    return pd.Timestamp(v).strftime('%Y/%m/%d').replace('/0', '/')

def summary(df): return {'orders': int(len(df)), 'pallet': int(df.Pallet.sum()), 'qty': int(df.QTY.sum())}

# --- 网页端 UI 主程序 ---
st.title("📦 Sigenergy 日报与看板自动化生成器")
st.caption("上传当天的 Outbound 共享表格，一键生成看板、邮件文本及统计报表。")

# 侧边栏：文件上传与参数设置
with st.sidebar:
    st.header("1. 上传数据文件")
    uploaded_file = st.file_uploader("选择 Sigenergy Outbound Excel 文件", type=["xlsx", "xlsm"])
    
    st.header("2. 日期设置")
    override_date = st.date_input("手动指定报告日期（可选）", value=None, help="如果文件名包含 YYYYMMDD，系统会自动提取；如果未能自动读取，可在此指定。")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    filename = uploaded_file.name

    # 确定报告日期
    m = re.search(r'(20\d{2})(0\d|1[0-2])([0-3]\d)', filename)
    if m:
        report_date = pd.Timestamp(datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))).normalize()
        date_source = f"文件名解析 ({report_date.strftime('%Y-%m-%d')})"
    elif override_date:
        report_date = pd.Timestamp(override_date).normalize()
        date_source = f"手动选择 ({report_date.strftime('%Y-%m-%d')})"
    else:
        report_date = pd.Timestamp.now().normalize()
        date_source = f"系统当前日期 ({report_date.strftime('%Y-%m-%d')})"

    st.success(f"已识别报告日期：**{report_date.strftime('%Y-%m-%d')}** （来源：{date_source}）")

    # 读取 Excel
    try:
        try:
            raw = load_openpyxl_bytes(file_bytes)
            reader_used = "openpyxl"
        except Exception as e:
            raw = load_xml_bytes(file_bytes)
            reader_used = "direct XML fallback"

        all_df = pd.concat([normalize(raw[s], s) for s in WAREHOUSE_SHEETS], ignore_index=True)
        
        if all_df.empty:
            st.error("表格中未找到有效订单数据！")
            st.stop()

        nxt = next_workday(report_date)
        eq = lambda c, d: all_df[c].eq(pd.Timestamp(d))
        
        today = all_df[eq('Outbound Date', report_date)].copy()
        new = all_df[eq('Release Date', report_date)].copy()
        nxtdf = all_df[eq('Planned Outbound', nxt)].copy()
        exc_base = all_df[eq('Planned Outbound', report_date) & all_df['Outbound Date'].isna()].copy()
        exc = exc_base[exc_base['Delivery Method'].str.upper().str.strip().eq('FTL')].copy()

        status = []
        for wh in WAREHOUSE_SHEETS.values():
            r = all_df[(all_df.Warehouse == wh) & eq('Release Date', report_date)]
            p = all_df[(all_df.Warehouse == wh) & eq('PCD', report_date)]
            status.append({
                'Warehouse': wh,
                'Total Orders': len(r),
                'Packing Completed': len(p),
                'Cancelled': int((r['Order status'] == 'Cancelled').sum()),
                'Sige Required On Hold': int((r['Order status'] == 'Sige Required On Hold').sum()),
                'Waiting OC': int((r['Order status'] == 'Waiting OC').sum()),
                'Shortage/Abnormal Stock': int(r['Status detail'].str.contains('shortage|abnormal', case=False, na=False).sum())
            })
        status_df = pd.DataFrame(status)
        waiting = int((new['Order status'] == 'Waiting OC').sum())
        shortage = int(new['Status detail'].str.contains('shortage|abnormal', case=False, na=False).sum())

        today_disp = today[['Order', 'Pallet', 'QTY', 'ADR', 'ETA', 'Warehouse']].copy()
        today_disp['ETA'] = today_disp.ETA.map(fmt_date)
        today_disp['WH'] = today_disp.pop('Warehouse').map({'AWL Moerdijk Outbound': 'MDK', 'AWL Energy Outbound': 'Energy', 'Botlek Outbound': 'Botlek'})

        exc_disp = exc[['Order', 'LSP', 'Planned Outbound', 'Tracking/AWB']].copy()
        exc_disp['Planned Outbound'] = exc_disp['Planned Outbound'].map(fmt_date)

        next_disp = nxtdf[['Order', 'Pallet', 'QTY', 'ADR', 'ETA', 'Warehouse']].copy()
        next_disp['ETA'] = next_disp.ETA.map(fmt_date)
        next_disp['WH'] = next_disp.pop('Warehouse').map({'AWL Moerdijk Outbound': 'MDK', 'AWL Energy Outbound': 'Energy', 'Botlek Outbound': 'Botlek'})

        sm = {'today_outbound': summary(today), 'new_orders': summary(new), 'next_workday': summary(nxtdf), 'exceptions': summary(exc)}

        # --- 看板展示区 ---
        st.divider()
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("今日新下单量 (New Orders)", f"{sm['new_orders']['orders']} 单", f"{sm['new_orders']['pallet']} 托 / {sm['new_orders']['qty']} pcs")
        col2.metric("今日日清日结 (Today Outbound)", f"{sm['today_outbound']['orders']} 单", f"{sm['today_outbound']['pallet']} 托 / {sm['today_outbound']['qty']} pcs")
        col3.metric(f"下一个工作日计划 ({nxt.strftime('%m/%d')})", f"{sm['next_workday']['orders']} 单", f"{sm['next_workday']['pallet']} 托 / {sm['next_workday']['qty']} pcs")
        col4.metric("整车异常单量 (FTL Exceptions)", f"{sm['exceptions']['orders']} 单", f"{sm['exceptions']['pallet']} 托 / {sm['exceptions']['qty']} pcs")

        st.info(f"💡 当前等待处理 (Waiting OC) 状态订单：**{waiting}** 单 | 缺库存订单：**{shortage}** 单")

        # Tabs 标签页展示
        tab1, tab2, tab3, tab4 = st.tabs(["📊 三仓实时状态", "🚚 今日日清日结明细", "⚠️ 整车出库异常", "📅 下个工作日计划"])

        with tab1:
            st.dataframe(status_df, use_container_width=True)

        with tab2:
            st.dataframe(today_disp, use_container_width=True)

        with tab3:
            st.dataframe(exc_disp, use_container_width=True)

        with tab4:
            st.dataframe(next_disp, use_container_width=True)

        # --- 导出一键生成包 ---
        st.divider()
        st.subheader("📥 导出与下载")

        # 生成 Excel 文件流
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as w:
            status_df.to_excel(w, index=False, sheet_name='Warehouse Status')
            today_disp.to_excel(w, index=False, sheet_name='日清日结')
            exc_disp.to_excel(w, index=False, sheet_name='当日出库异常')
            next_disp.to_excel(w, index=False, sheet_name='Next Workday')
            new.to_excel(w, index=False, sheet_name='New Orders')
            all_df.to_excel(w, index=False, sheet_name='Normalized Data')
            for ws in w.book.worksheets:
                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = min(45, max(10, max(len(str(c.value or '')) for c in col) + 2))
        excel_data = excel_buffer.getvalue()

        # 生成 Email Text 文本
        status_mail = status_df.rename(columns={'Warehouse': 'Warehouse', 'Total Orders': 'Release Orders', 'Packing Completed': 'Packing Completed', 'Cancelled': 'Cancelled', 'Sige Required On Hold': 'Sige Required On Hold', 'Waiting OC': 'Waiting OC', 'Shortage/Abnormal Stock': 'Shortage/Abnormal Stock'})
        email_txt = '\n'.join([
            f'目前等待处理（Waiting OC）状态订单{waiting}单，{shortage}单缺库存', '',
            'Warehouse Status', status_mail.to_string(index=False), '',
            f"今日出库日清日结 {sm['today_outbound']['orders']} 單，共计 {sm['today_outbound']['pallet']} plts, {sm['today_outbound']['qty']} pcs",
            today_disp.to_string(index=False), '',
            f"今日整车出庫異常訂單共 {sm['exceptions']['orders']} 單",
            exc_disp.to_string(index=False), '',
            f"下个工作日三仓出库預計 {sm['next_workday']['orders']} 单, {sm['next_workday']['pallet']} 托，{sm['next_workday']['qty']} pcs ，具体出库订单可在腾讯在线文档中查看"
        ])

        # 生成 ZIP 打包文件
        zip_buffer = io.BytesIO()
        base_name = f"Sigenergy_Daily_Outbound_Report_{report_date:%Y%m%d}"
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr(f"{base_name}.xlsx", excel_data)
            z.writestr(f"{base_name}_email.txt", email_txt.encode('utf-8'))
            z.writestr(f"{base_name}_FTL_exceptions.csv", exc_disp.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig'))
        zip_data = zip_buffer.getvalue()

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="📦 一键打包下载全部结果 (.ZIP)",
                data=zip_data,
                file_name=f"{base_name}_ALL.zip",
                mime="application/zip"
            )
        with c2:
            st.download_button(
                label="📊 仅下载 Excel 汇总报表 (.xlsx)",
                data=excel_data,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with st.expander("📧 查看/复制邮件正文文本"):
            st.code(email_txt, language="text")

    except Exception as e:
        st.error(f"处理文件时出错: {e}")
        st.exception(e)

else:
    st.info("👈 请在左侧边栏上传当天的 Excel 文件以开始生成看板。")