Sigenergy Daily Report V3.4 — KPI Dashboard + Customer Email

运行：python daily_outbound_report_v3_4.py
不带 input 参数时自动弹出 Windows 文件选择窗口。
日期优先级：文件名 YYYYMMDD > --date > 电脑日期。

输出：
- *_EMAIL.html：客户邮件格式
- *_DASHBOARD.html：内部 KPI Dashboard
- *.xlsx：Excel 报告
- *_email.txt：纯文本邮件
- *_FTL_exceptions.csv：FTL 异常
- *_run_summary.json：运行信息
- run.log：日志

异常规则：
当日出库异常 = Planned Outbound Date = 报告日期 AND Outbound Date 为空。
FTL 异常 = 当日出库异常 AND Delivery Method = FTL。

兼容性：openpyxl 读取失败时自动使用 XLSX XML fallback。
