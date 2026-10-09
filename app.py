# -*- coding: utf-8 -*-
"""
Sigenergy Daily Outbound Report - Streamlit Web Version
V3.5: HTML 邮件正文（带边框表格）+ Excel/CSV/TXT/ZIP 导出

运行：
    pip install streamlit pandas openpyxl
    streamlit run app.py
"""

import html
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


# ---------------- 页面配置 ----------------
st.set_page_config(
    page_title="Sigenergy Daily Outbound Dashboard",
    page_icon="📦",
    layout="wide",
)

APP_VERSION = "V3.5 Web Edition"
WAREHOUSE_SHEETS = {
    "AWL Moerdijk 出库": "AWL Moerdijk Outbound",
    "AWL Energy  出库": "AWL Energy Outbound",
    "Botlek出库": "Botlek Outbound",
}
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EXCEL_EPOCH = pd.Timestamp("1899-12-30")
WAREHOUSE_SHORT = {
    "AWL Moerdijk Outbound": "MDK",
    "AWL Energy Outbound": "Energy",
    "Botlek Outbound": "Botlek",
}


# ---------------- 通用数据处理函数 ----------------
def norm(v):
    if v is None or pd.isna(v):
        return ""
    return re.sub(r"\s+", "", str(v)).strip().lower()


def clean(v):
    return "" if v is None or pd.isna(v) else str(v).strip()


def pick(df, cands):
    exact = {norm(c): c for c in df.columns}
    for candidate in cands:
        if norm(candidate) in exact:
            return exact[norm(candidate)]
    normalized_candidates = [norm(c) for c in cands]
    for col in df.columns:
        n = norm(col)
        if any(c and (c in n or n in c) for c in normalized_candidates):
            return col
    return None


def colnum(s):
    n = 0
    for c in s:
        n = n * 26 + ord(c) - 64
    return n


def shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return [
        "".join(t.text or "" for t in si.iter("{%s}t" % NS["m"]))
        for si in root.findall("m:si", NS)
    ]


def sheet_paths(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rel = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rels = {x.attrib["Id"]: x.attrib["Target"] for x in rel}
    out = {}
    for sheet in wb.find("m:sheets", NS):
        target = rels[sheet.attrib["{%s}id" % NS["r"]]]
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = "xl/" + target
        out[sheet.attrib["name"]] = path
    return out


def read_xml_sheet(z, path, ss):
    root = ET.fromstring(z.read(path))
    rows = []
    for row in root.findall(".//m:sheetData/m:row", NS):
        vals = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            v = cell.find("m:v", NS)
            val = "" if v is None else v.text
            typ = cell.attrib.get("t")
            if typ == "s" and val != "":
                try:
                    val = ss[int(val)]
                except (ValueError, IndexError):
                    val = ""
            elif typ == "inlineStr":
                val = "".join(
                    t.text or "" for t in cell.iter("{%s}t" % NS["m"])
                )
            elif typ == "b":
                val = "TRUE" if val == "1" else "FALSE"
            vals[match.group(1)] = "" if val is None else val
        if vals:
            rows.append(vals)

    if not rows:
        return pd.DataFrame()

    keys = sorted(rows[0], key=colnum)
    headers = [rows[0].get(k) or k for k in keys]
    seen = {}
    columns = []
    for header in headers:
        header = str(header).strip() or "Unnamed"
        seen[header] = seen.get(header, 0) + 1
        columns.append(header if seen[header] == 1 else f"{header}__{seen[header]}")

    df = pd.DataFrame(
        [[row.get(k, "") for k in keys] for row in rows[1:]],
        columns=columns,
    )
    keep = [c for c in df.columns if df[c].astype(str).str.strip().ne("").any()]
    return df.loc[:, keep]


def load_xml_bytes(file_bytes):
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        if z.testzip():
            raise ValueError("XLSX 压缩包损坏")
        paths = sheet_paths(z)
        ss = shared_strings(z)
        found_sheets = {}
        normalized_paths = {norm(name): name for name in paths}
        for target_sheet in WAREHOUSE_SHEETS:
            matched = normalized_paths.get(norm(target_sheet))
            if matched:
                found_sheets[target_sheet] = read_xml_sheet(z, paths[matched], ss)
            else:
                raise ValueError(f"缺少仓库 Sheet: {target_sheet}")
        return found_sheets


def load_openpyxl_bytes(file_bytes):
    if openpyxl is None:
        raise RuntimeError("openpyxl 未安装")
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), read_only=True, data_only=True
    )
    try:
        out = {}
        sheet_map = {norm(sheet): sheet for sheet in wb.sheetnames}
        for sheet in WAREHOUSE_SHEETS:
            if norm(sheet) not in sheet_map:
                raise ValueError(f"缺少仓库 Sheet: {sheet}")
            actual_sheet_name = sheet_map[norm(sheet)]
            values = list(wb[actual_sheet_name].values)
            if values:
                headers = [clean(x) or f"Unnamed_{i + 1}" for i, x in enumerate(values[0])]
                # 防止 Excel 中存在重名列导致 pandas 报错或列覆盖
                seen = {}
                unique_headers = []
                for header in headers:
                    seen[header] = seen.get(header, 0) + 1
                    unique_headers.append(
                        header if seen[header] == 1 else f"{header}__{seen[header]}"
                    )
                out[sheet] = pd.DataFrame(values[1:], columns=unique_headers)
            else:
                out[sheet] = pd.DataFrame()
        return out
    finally:
        wb.close()


def parse_date(v):
    if v is None or pd.isna(v) or str(v).strip() == "":
        return pd.NaT
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).normalize()
    s = str(v).strip()
    try:
        n = float(s)
        if 20000 <= n <= 80000:
            return (EXCEL_EPOCH + pd.to_timedelta(n, unit="D")).normalize()
    except (ValueError, TypeError, OverflowError):
        pass
    d = pd.to_datetime(s, errors="coerce")
    return pd.NaT if pd.isna(d) else pd.Timestamp(d).normalize()


def dser(series):
    return series.map(parse_date)


def nser(series):
    """解析托盘数/QTY 等计数字段；例如 23TBD、23 TBD 按 23 计算。"""
    def parse_count(value):
        if value is None or pd.isna(value):
            return 0
        if isinstance(value, (int, float)):
            try:
                return int(round(float(value)))
            except (ValueError, TypeError, OverflowError):
                return 0

        text = str(value).strip().replace(",", "")
        if not text:
            return 0
        # 只读取开头的数字部分，因此“23TBD”或“23 TBD”会得到 23。
        match = re.match(r"^\s*([+-]?\d+(?:\.\d+)?)", text)
        if not match:
            return 0
        try:
            return int(round(float(match.group(1))))
        except (ValueError, TypeError, OverflowError):
            return 0

    return series.map(parse_count).astype("int64")


def next_workday(d):
    d = pd.Timestamp(d).normalize() + pd.Timedelta(days=1)
    while d.weekday() >= 5:
        d += pd.Timedelta(days=1)
    return d


def normalize(df, sheet):
    order = pick(df, ["DN", " DN", "order"])
    status_col = pick(df, ["Order status", "Order Status"])
    release = pick(df, ["Release Date"])
    pcd = pick(df, ["PCD", "Picking Complete Date(APD)", "Picking Complete Date"])
    planned = pick(df, ["Outbound Planned Date", "OutboundPlannedDate"])
    actual = pick(df, ["Outbound Date"])
    pallet = pick(df, ["Pallet"])
    qty = pick(df, ["QTY", "Quantity"])
    adr = pick(df, ["ADR"])
    eta = pick(df, ["ETA"])
    lsp = pick(df, ["LSP"])
    track = pick(df, ["Tracking/AWB", "Tracking/AWB "])
    detail = pick(df, ["Status"])
    delivery = pick(df, ["Delivery Method"])

    if not order:
        raise ValueError(f"{sheet} 找不到 DN/order")
    for name, col in [
        ("Release Date", release),
        ("PCD", pcd),
        ("Outbound Planned Date", planned),
        ("Outbound Date", actual),
    ]:
        if not col:
            raise ValueError(f"{sheet} 缺少关键字段: {name}")

    def normalize_status(value):
        n = norm(value)
        if n in ("cancelled", "canceled"):
            return "Cancelled"
        if "sigerequiredonhold" in n:
            return "Sige Required On Hold"
        if "waitingoc" in n:
            return "Waiting OC"
        return clean(value)

    z = pd.DataFrame({
        "Warehouse": WAREHOUSE_SHEETS[sheet],
        "Order": df[order].map(clean),
        "Order status": df[status_col].map(normalize_status) if status_col else "",
        "Release Date": dser(df[release]),
        "PCD": dser(df[pcd]),
        "Planned Outbound": dser(df[planned]),
        "Outbound Date": dser(df[actual]),
        "Pallet": nser(df[pallet]) if pallet else 0,
        "QTY": nser(df[qty]) if qty else 0,
        "ADR": df[adr].map(clean) if adr else "",
        "ETA": dser(df[eta]) if eta else pd.NaT,
        "LSP": df[lsp].map(clean) if lsp else "",
        "Tracking/AWB": df[track].map(clean) if track else "",
        "Status detail": df[detail].map(clean) if detail else "",
        "Delivery Method": df[delivery].map(clean) if delivery else "",
    })
    return z[z["Order"].ne("")].copy()


def fmt_date(v):
    if pd.isna(v):
        return ""
    return pd.Timestamp(v).strftime("%Y/%m/%d").replace("/0", "/")


def summary(df):
    return {
        "orders": int(len(df)),
        "pallet": int(df["Pallet"].sum()),
        "qty": int(df["QTY"].sum()),
    }


# ---------------- HTML 邮件生成 ----------------
def _display_value(value):
    """把 pandas/日期/空值转换成适合邮件表格显示的文本。"""
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return fmt_date(value)
    return str(value)


def html_table(df, columns=None, headers=None):
    """
    生成兼容常见邮件客户端的 HTML 表格。
    样式使用 inline CSS，避免邮件客户端剥离 <style> 后丢失边框。
    """
    view = df.copy()
    if columns is not None:
        view = view.loc[:, columns]
    if headers is None:
        headers = list(view.columns)

    parts = [
        '<table border="1" cellpadding="6" cellspacing="0" '
        'style="border-collapse:collapse;border:1px solid #999;'
        'font-family:Arial,sans-serif;font-size:13px;color:#222;'
        'margin:6px 0 16px 0;">',
        "<thead><tr>",
    ]
    for header in headers:
        parts.append(
            '<th style="border:1px solid #999;background-color:#eaf0f6;'
            'padding:6px 8px;text-align:left;font-weight:bold;">'
            f"{html.escape(str(header))}</th>"
        )
    parts.append("</tr></thead><tbody>")

    for row in view.itertuples(index=False, name=None):
        parts.append("<tr>")
        for value in row:
            text_value = _display_value(value)
            parts.append(
                '<td style="border:1px solid #999;padding:6px 8px;'
                'vertical-align:top;">'
                f"{html.escape(text_value).replace(chr(10), '<br>')}</td>"
            )
        parts.append("</tr>")

    if len(view) == 0:
        parts.append(
            f'<tr><td colspan="{max(len(headers), 1)}" '
            'style="border:1px solid #999;padding:8px;color:#666;'
            'font-style:italic;">暂无记录</td></tr>'
        )
    parts.append("</tbody></table>")
    return "".join(parts)


def build_rich_email_html(report_date, next_date, status_df, today_disp,
                          exc_disp, next_disp, sm, waiting, shortage):
    """生成中文富文本邮件正文，表格使用内联边框样式。"""
    date_text = pd.Timestamp(report_date).strftime("%Y-%m-%d")
    next_text = pd.Timestamp(next_date).strftime("%Y-%m-%d")

    status_mail = status_df.rename(columns={
        "Warehouse": "仓库",
        "Total Orders": "释放订单数",
        "Packing Completed": "完成打包数",
        "Cancelled": "已取消",
        "Sige Required On Hold": "Sige要求暂缓",
        "Waiting OC": "等待OC",
        "Shortage/Abnormal Stock": "缺库存/异常库存",
    }).copy()
    status_mail["仓库"] = status_mail["仓库"].replace({
        "AWL Moerdijk Outbound": "AWL Moerdijk",
        "AWL Energy Outbound": "AWL Energy",
        "Botlek Outbound": "Botlek",
    })

    today_mail = today_disp.rename(columns={
        "Order": "订单号（DN）", "Pallet": "托盘数", "QTY": "数量（pcs）",
        "ADR": "ADR", "ETA": "预计到达日期", "WH": "仓库",
    }).copy()
    exc_mail = exc_disp.rename(columns={
        "Order": "订单号（DN）", "LSP": "物流服务商（LSP）",
        "Planned Outbound": "计划出库日期", "Tracking/AWB": "运单号（Tracking/AWB）",
    }).copy()
    next_mail = next_disp.rename(columns={
        "Order": "订单号（DN）", "Pallet": "托盘数", "QTY": "数量（pcs）",
        "ADR": "ADR", "ETA": "预计到达日期", "WH": "仓库",
    }).copy()

    for frame in (today_mail, next_mail):
        frame["仓库"] = frame["仓库"].replace({"MDK": "Moerdijk", "Energy": "Energy", "Botlek": "Botlek"})

    sections = [
        '<!doctype html><html><head><meta charset="utf-8"></head>',
        '<body style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#222;">',
        f'<p>大家好，</p><p>以下是 <strong>{html.escape(date_text)}</strong> 的 Sigenergy 每日出库报告。</p>',
        '<p style="margin:12px 0 6px;"><strong>目前等待处理（Waiting OC）状态订单 '
        f'{waiting} 单，{shortage} 单缺库存/异常库存。</strong></p>',
        '<h3 style="font-family:Arial,sans-serif;margin:18px 0 6px;font-size:16px;">三仓状态汇总</h3>',
        html_table(status_mail),
        '<h3 style="font-family:Arial,sans-serif;margin:18px 0 6px;font-size:16px;">今日出库日清日结</h3>',
        f'<p>今日出库日清日结 <strong>{sm["today_outbound"]["orders"]} 单</strong>，共计 '
        f'<strong>{sm["today_outbound"]["pallet"]} plts，{sm["today_outbound"]["qty"]} pcs</strong>。</p>',
        html_table(today_mail),
        '<h3 style="font-family:Arial,sans-serif;margin:18px 0 6px;font-size:16px;">今日整车出库异常</h3>',
        f'<p>今日整车出库异常订单共 <strong>{sm["exceptions"]["orders"]} 单</strong>。</p>',
        html_table(exc_mail),
        '<h3 style="font-family:Arial,sans-serif;margin:18px 0 6px;font-size:16px;">下个工作日出库计划</h3>',
        f'<p>下个工作日（{html.escape(next_text)}）三仓出库预计 '
        f'<strong>{sm["next_workday"]["orders"]} 单，{sm["next_workday"]["pallet"]} 托，'
        f'{sm["next_workday"]["qty"]} pcs</strong>，具体出库订单可在腾讯在线文档中查看。</p>',
        html_table(next_mail),
        '<p>谢谢！</p>',
        '</body></html>',
    ]
    return "".join(sections)


# ---------------- 网页端 UI 主程序 ----------------
st.title("📦 Sigenergy 日报与看板自动化生成器")
st.caption(f"{APP_VERSION} · 上传当天的 Outbound 共享表格，一键生成看板、邮件正文及统计报表。")

with st.sidebar:
    st.header("1. 上传数据文件")
    uploaded_file = st.file_uploader(
        "选择 Sigenergy Outbound Excel 文件",
        type=["xlsx", "xlsm"],
    )
    st.header("2. 日期设置")
    override_date = st.date_input(
        "手动指定报告日期（可选）",
        value=None,
        help="如果文件名包含 YYYYMMDD，系统会自动提取；如果未能自动读取，可在此指定。",
    )

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name

    # 先从文件名解析日期；若解析失败再使用手动日期，最后使用系统日期。
    match = re.search(r"(20\d{2})(0\d|1[0-2])([0-3]\d)", filename)
    if match:
        try:
            report_date = pd.Timestamp(datetime(
                int(match.group(1)), int(match.group(2)), int(match.group(3))
            )).normalize()
            date_source = f"文件名解析 ({report_date:%Y-%m-%d})"
        except ValueError:
            report_date = None
    else:
        report_date = None

    if report_date is None and override_date:
        report_date = pd.Timestamp(override_date).normalize()
        date_source = f"手动选择 ({report_date:%Y-%m-%d})"
    elif report_date is None:
        report_date = pd.Timestamp.now().normalize()
        date_source = f"系统当前日期 ({report_date:%Y-%m-%d})"

    st.success(
        f"已识别报告日期：**{report_date:%Y-%m-%d}** （来源：{date_source}）"
    )

    try:
        try:
            raw = load_openpyxl_bytes(file_bytes)
            reader_used = "openpyxl"
        except Exception as openpyxl_error:
            try:
                raw = load_xml_bytes(file_bytes)
                reader_used = "direct XML fallback"
            except Exception as xml_error:
                raise ValueError(
                    "Excel 读取失败。openpyxl 错误："
                    f"{openpyxl_error}；XML fallback 错误：{xml_error}"
                ) from xml_error

        all_df = pd.concat(
            [normalize(raw[sheet], sheet) for sheet in WAREHOUSE_SHEETS],
            ignore_index=True,
        )

        if all_df.empty:
            st.error("表格中未找到有效订单数据！")
            st.stop()

        next_date = next_workday(report_date)
        date_eq = lambda column, day: all_df[column].eq(pd.Timestamp(day))

        # 当日已实际出库
        today = all_df[date_eq("Outbound Date", report_date)].copy()
        # 当日新释放订单
        new = all_df[date_eq("Release Date", report_date)].copy()
        # 下一个工作日计划出库
        next_df = all_df[date_eq("Planned Outbound", next_date)].copy()
        # 异常基数：计划出库日期为报告日，但尚无实际出库日期
        exception_base = all_df[
            date_eq("Planned Outbound", report_date)
            & all_df["Outbound Date"].isna()
        ].copy()
        # FTL 异常：在上述基数中进一步筛选 Delivery Method = FTL
        exceptions = exception_base[
            exception_base["Delivery Method"].str.upper().str.strip().eq("FTL")
        ].copy()

        warehouse_status = []
        for warehouse in WAREHOUSE_SHEETS.values():
            released = all_df[
                (all_df["Warehouse"] == warehouse)
                & date_eq("Release Date", report_date)
            ]
            packed = all_df[
                (all_df["Warehouse"] == warehouse)
                & date_eq("PCD", report_date)
            ]
            warehouse_status.append({
                "Warehouse": warehouse,
                "Total Orders": len(released),
                "Packing Completed": len(packed),
                "Cancelled": int((released["Order status"] == "Cancelled").sum()),
                "Sige Required On Hold": int(
                    (released["Order status"] == "Sige Required On Hold").sum()
                ),
                "Waiting OC": int((released["Order status"] == "Waiting OC").sum()),
                "Shortage/Abnormal Stock": int(
                    released["Status detail"].str.contains(
                        "shortage|abnormal", case=False, na=False
                    ).sum()
                ),
            })
        status_df = pd.DataFrame(warehouse_status)
        waiting = int((new["Order status"] == "Waiting OC").sum())
        shortage = int(
            new["Status detail"].str.contains(
                "shortage|abnormal", case=False, na=False
            ).sum()
        )

        today_disp = today[
            ["Order", "Pallet", "QTY", "ADR", "ETA", "Warehouse"]
        ].copy()
        today_disp["ETA"] = today_disp["ETA"].map(fmt_date)
        today_disp["WH"] = today_disp.pop("Warehouse").map(WAREHOUSE_SHORT)

        exc_disp = exceptions[
            ["Order", "LSP", "Planned Outbound", "Tracking/AWB"]
        ].copy()
        exc_disp["Planned Outbound"] = exc_disp["Planned Outbound"].map(fmt_date)

        next_disp = next_df[
            ["Order", "Pallet", "QTY", "ADR", "ETA", "Warehouse"]
        ].copy()
        next_disp["ETA"] = next_disp["ETA"].map(fmt_date)
        next_disp["WH"] = next_disp.pop("Warehouse").map(WAREHOUSE_SHORT)

        sm = {
            "today_outbound": summary(today),
            "new_orders": summary(new),
            "next_workday": summary(next_df),
            "exceptions": summary(exceptions),
        }

        st.caption(f"Excel reader: {reader_used}")
        st.divider()

        # KPI 指标卡
        col1, col2, col3, col4 = st.columns(4)
        col1.metric(
            "今日新下单量 (New Orders)",
            f"{sm['new_orders']['orders']} 单",
            f"{sm['new_orders']['pallet']} 托 / {sm['new_orders']['qty']} pcs",
        )
        col2.metric(
            "今日日清日结 (Today Outbound)",
            f"{sm['today_outbound']['orders']} 单",
            f"{sm['today_outbound']['pallet']} 托 / {sm['today_outbound']['qty']} pcs",
        )
        col3.metric(
            f"下一个工作日计划 ({next_date:%m/%d})",
            f"{sm['next_workday']['orders']} 单",
            f"{sm['next_workday']['pallet']} 托 / {sm['next_workday']['qty']} pcs",
        )
        col4.metric(
            "整车异常单量 (FTL Exceptions)",
            f"{sm['exceptions']['orders']} 单",
            f"{sm['exceptions']['pallet']} 托 / {sm['exceptions']['qty']} pcs",
        )

        st.info(
            f"💡 当前等待处理 (Waiting OC) 状态订单：**{waiting}** 单"
            f" | 缺库存/异常库存订单：**{shortage}** 单"
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 三仓实时状态",
            "🚚 今日日清日结明细",
            "⚠️ 整车出库异常",
            "📅 下个工作日计划",
        ])
        with tab1:
            st.dataframe(status_df, use_container_width=True)
        with tab2:
            st.dataframe(today_disp, use_container_width=True)
        with tab3:
            st.dataframe(exc_disp, use_container_width=True)
        with tab4:
            st.dataframe(next_disp, use_container_width=True)

        # ---------------- 导出 Excel ----------------
        st.divider()
        st.subheader("📥 导出与下载")

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            status_df.to_excel(writer, index=False, sheet_name="Warehouse Status")
            today_disp.to_excel(writer, index=False, sheet_name="日清日结")
            exc_disp.to_excel(writer, index=False, sheet_name="当日出库异常")
            next_disp.to_excel(writer, index=False, sheet_name="Next Workday")
            new.to_excel(writer, index=False, sheet_name="New Orders")
            all_df.to_excel(writer, index=False, sheet_name="Normalized Data")

            for worksheet in writer.book.worksheets:
                worksheet.freeze_panes = "A2"
                worksheet.auto_filter.ref = worksheet.dimensions
                for column_cells in worksheet.columns:
                    max_length = max(
                        (len(str(cell.value or "")) for cell in column_cells),
                        default=10,
                    )
                    worksheet.column_dimensions[
                        column_cells[0].column_letter
                    ].width = min(45, max(10, max_length + 2))
        excel_data = excel_buffer.getvalue()

        # ---------------- 邮件正文：纯文本 + 富 HTML ----------------
        status_txt = status_df.rename(columns={
            "Warehouse": "仓库",
            "Total Orders": "释放订单数",
            "Packing Completed": "完成打包数",
            "Cancelled": "已取消",
            "Sige Required On Hold": "Sige要求暂缓",
            "Waiting OC": "等待OC",
            "Shortage/Abnormal Stock": "缺库存/异常库存",
        }).copy()
        status_txt["仓库"] = status_txt["仓库"].replace({
            "AWL Moerdijk Outbound": "AWL Moerdijk",
            "AWL Energy Outbound": "AWL Energy",
            "Botlek Outbound": "Botlek",
        })
        today_txt = today_disp.rename(columns={"Order": "订单号（DN）", "Pallet": "托盘数", "QTY": "数量（pcs）", "ADR": "ADR", "ETA": "预计到达日期", "WH": "仓库"})
        exc_txt = exc_disp.rename(columns={"Order": "订单号（DN）", "LSP": "物流服务商（LSP）", "Planned Outbound": "计划出库日期", "Tracking/AWB": "运单号（Tracking/AWB）"})
        next_txt = next_disp.rename(columns={"Order": "订单号（DN）", "Pallet": "托盘数", "QTY": "数量（pcs）", "ADR": "ADR", "ETA": "预计到达日期", "WH": "仓库"})
        email_txt = "\n".join([
            f"目前等待处理（Waiting OC）状态订单{waiting}单，{shortage}单缺库存/异常库存。",
            "",
            "三仓状态汇总",
            status_txt.to_string(index=False),
            "",
            f"今日出库日清日结 {sm['today_outbound']['orders']} 单，共计 {sm['today_outbound']['pallet']} plts, {sm['today_outbound']['qty']} pcs。",
            today_txt.to_string(index=False),
            "",
            f"今日整车出库异常订单共 {sm['exceptions']['orders']} 单。",
            exc_txt.to_string(index=False),
            "",
            f"下个工作日三仓出库预计 {sm['next_workday']['orders']} 单, {sm['next_workday']['pallet']} 托，{sm['next_workday']['qty']} pcs，具体出库订单可在腾讯在线文档中查看。",
            next_txt.to_string(index=False),
        ])

        email_html = build_rich_email_html(
            report_date=report_date,
            next_date=next_date,
            status_df=status_df,
            today_disp=today_disp,
            exc_disp=exc_disp,
            next_disp=next_disp,
            sm=sm,
            waiting=waiting,
            shortage=shortage,
        )
        html_data = email_html.encode("utf-8")

        # ---------------- ZIP 打包 ----------------
        base_name = f"Sigenergy_Daily_Outbound_Report_{report_date:%Y%m%d}"
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(f"{base_name}.xlsx", excel_data)
            archive.writestr(
                f"{base_name}_email.txt", email_txt.encode("utf-8")
            )
            archive.writestr(
                f"{base_name}_email.html", html_data
            )
            archive.writestr(
                f"{base_name}_FTL_exceptions.csv",
                exc_disp.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
            )
        zip_data = zip_buffer.getvalue()

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button(
                label="📦 一键打包下载全部结果 (.ZIP)",
                data=zip_data,
                file_name=f"{base_name}_ALL.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                label="📊 下载 Excel 汇总报表 (.xlsx)",
                data=excel_data,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c3:
            st.download_button(
                label="✉️ 下载带边框表格的邮件正文 (.HTML)",
                data=html_data,
                file_name=f"{base_name}_email.html",
                mime="text/html",
                use_container_width=True,
            )

        with st.expander("📧 预览富文本邮件正文（表格带边框）", expanded=True):
            st.caption(
                "建议：下载 HTML 文件并用浏览器打开，按 Ctrl+A、Ctrl+C，"
                "然后粘贴到网易邮箱正文。这样通常比复制纯文本更容易保留表格边框。"
            )
            st.components.v1.html(email_html, height=900, scrolling=True)

        with st.expander("📝 查看/复制纯文本邮件正文"):
            st.code(email_txt, language="text")

        with st.expander("🔎 查看 FTL 异常筛选明细"):
            st.write(
                "筛选条件：Planned Outbound = 报告日期、Outbound Date 为空、"
                "Delivery Method = FTL。"
            )
            st.dataframe(exceptions, use_container_width=True)

    except Exception as e:
        st.error(f"处理文件时出错：{e}")
        st.exception(e)

else:
    st.info("👈 请在左侧边栏上传当天的 Excel 文件以开始生成看板。")
