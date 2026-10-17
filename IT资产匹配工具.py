# -*- coding: utf-8 -*-
r"""
IT资产台账匹配工具 3.0  ——  通用 Excel 列映射匹配 / 回填工具

能力:
  * 自动识别上传台账的首行表头, 所有列都可在下拉框中自由选择
  * 台账「源列」的数据  ->  去资产清单「查找列」比对  ->  命中后取「返回列」的值
    ->  写回台账指定的「回填列」(可选已存在的列, 也可新建列)
  * 资产清单支持多工作表, 各 sheet 列位置不同也能各自按表头名定位
  * 匹配模式: 智能(优先提取IP, 无IP则整格) / 仅IP / 整格精确 / 按分隔符拆分
  * 未匹配行可自定义填充值(默认留空)

GUI:   双击 IT资产匹配工具.exe   (或 启动IT资产匹配工具.bat)
CLI:   IT资产匹配工具.exe --cli <台账.xlsx> <资产清单.xlsx> [输出.xlsx]

图标:  imag/logo.ico (exe 图标/标题栏/任务栏)  imag/logo_ui.png (界面 logo)
打包:  pyinstaller --noconfirm --onefile --noconsole --clean ^
         --name IT资产匹配工具 --icon imag\logo.ico ^
         --add-data "imag\logo.ico;imag" ^
         --add-data "imag\logo_128.png;imag" ^
         --add-data "imag\logo_ui.png;imag" ^
         IT资产匹配工具.py
"""

import os
import re
import sys
import datetime

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    import openpyxl
except ImportError:
    sys.stderr.write("缺少依赖 openpyxl, 请先执行: pip install openpyxl\n")
    sys.exit(1)


# ------------------------------------------------- 路径 / 内置资源 (源码与 exe 通用)
def app_dir():
    """程序所在目录: 打包成 exe 后 = exe 所在目录, 未打包 = 脚本所在目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(*parts):
    """资源定位: 先找 程序同级目录, 再找 exe 内置资源(_MEIPASS)"""
    local = os.path.join(app_dir(), *parts)
    if os.path.exists(local):
        return local
    base = getattr(sys, "_MEIPASS", None)
    if base:
        bundled = os.path.join(base, *parts)
        if os.path.exists(bundled):
            return bundled          # 兜底: 找不到也返回路径, 调用侧自行判断
    return local


ICON_ICO = resource_path("imag", "logo.ico")        # 多尺寸图标: 标题栏/任务栏/exe
ICON_PNG = resource_path("imag", "logo_128.png")    # 备用位图图标
LOGO_UI  = resource_path("imag", "logo_ui.png")     # 界面标题左侧 logo

APP_ID = "BraveKunlun.ITAssetMatcher.3.0"           # 任务栏独立分组标识


def _bind_app_id():
    """未打包运行时, 让任务栏用窗口自身图标而不是 python 图标"""
    if getattr(sys, "frozen", False):
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


VERSION = "3.0"
APP_TITLE = "IT资产台账匹配工具%s" % VERSION

# 资产清单默认值: exe/脚本同级优先, 保证打包后拷贝到别的目录仍能自动带出
_INV_CANDIDATES = [
    os.path.join(app_dir(), "资产清单.xlsx"),
    os.path.join(app_dir(), "imag", "资产清单.xlsx"),
    r"C:\Users\admin\Desktop\资产匹配\资产清单.xlsx",
]
DEFAULT_INVENTORY = next((p for p in _INV_CANDIDATES if os.path.exists(p)),
                         _INV_CANDIDATES[-1])
NEW_COL = "（新建列…）"
ALL_SHEETS = "（全部工作表）"

# ------------------------------------------------- Windows 11 Fluent 配色
C_BG       = "#f3f3f3"   # 窗口底色
C_PANEL    = "#ffffff"   # 卡片
C_LINE     = "#e0e0e0"   # 卡片描边
C_CTRL_BD  = "#d1d1d1"   # 控件描边
C_TEXT     = "#1b1b1b"
C_MUTED    = "#616161"
C_ACCENT   = "#0078d4"   # Win11 强调色
C_ACCENT_H = "#005a9e"
C_ACCENT_P = "#004c87"
C_BTN      = "#ffffff"
C_BTN_H    = "#f6f6f6"
C_BTN_P    = "#ebebeb"
C_SEL      = "#cfe4fa"   # 列表选中
C_HEAD     = "#fafafa"   # 表头
C_OK_ROW   = "#e8f4ea"
C_NO_ROW   = "#fbeaea"
C_OK_DOT   = "#0f7b3f"
C_NO_DOT   = "#c42b1c"

FONT_UI    = ("Microsoft YaHei UI", 9)
FONT_UI_S  = ("Microsoft YaHei UI", 9)
FONT_SM    = ("Microsoft YaHei UI", 8)
FONT_BOLD  = ("Microsoft YaHei UI", 9, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 14, "bold")
FONT_SUB   = ("Microsoft YaHei UI", 9)


# ================================================================= 文本 / 取值
_FULL2HALF = {}
for _i in range(0xFF01, 0xFF5F):
    _FULL2HALF[chr(_i)] = chr(_i - 0xFEE0)
_FULL2HALF[0x3000] = " "


def normalize_text(value):
    """全角转半角 + 去首尾空白; 数字型去掉多余小数"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).translate(_FULL2HALF).strip()


IP_RE = re.compile(r"\d{1,3}(?:\.\d{1,3}){3}")
RANGE_RE = re.compile(
    r"(\d{1,3}(?:\.\d{1,3}){3})\s*[-~—－至到]\s*(?:(\d{1,3}(?:\.\d{1,3}){3})|(\d{1,3}))"
)
DEFAULT_DELIMS = "\r\n\t,，;；、/| "


def _valid_ip(ip):
    parts = ip.split(".")
    return len(parts) == 4 and all(p.isdigit() and 0 <= int(p) <= 255 for p in parts)


def split_ips(value):
    """提取全部 IP(含区间展开), 去重保序"""
    text = normalize_text(value)
    if not text:
        return []
    found = []

    def add(ip):
        ip = ".".join(str(int(x)) for x in ip.split("."))
        if _valid_ip(ip) and ip not in found:
            found.append(ip)

    for m in RANGE_RE.finditer(text):
        start, end_ip, end_num = m.group(1), m.group(2), m.group(3)
        if not _valid_ip(start):
            continue
        sp = [int(x) for x in start.split(".")]
        if end_ip and _valid_ip(end_ip):
            ep = [int(x) for x in end_ip.split(".")]
            if sp[:3] == ep[:3] and ep[3] >= sp[3]:
                for i in range(sp[3], ep[3] + 1):
                    add(".".join(str(x) for x in sp[:3] + [i]))
        elif end_num:
            e = int(end_num)
            if sp[3] <= e <= 255:
                for i in range(sp[3], e + 1):
                    add(".".join(str(x) for x in sp[:3] + [i]))

    for m in IP_RE.finditer(RANGE_RE.sub(" ", text)):
        add(m.group(0))
    return found


def extract_keys(value, mode="auto", delims=DEFAULT_DELIMS):
    """按匹配模式从单元格取出用于比对的键列表"""
    text = normalize_text(value)
    if not text:
        return []
    if mode == "ip":
        return split_ips(text)
    if mode == "cell":
        return [text]
    if mode == "split":
        pattern = "[" + re.escape(delims or DEFAULT_DELIMS) + "]+"
        out = []
        for part in re.split(pattern, text):
            part = part.strip()
            if part and part not in out:
                out.append(part)
        return out
    # auto: 优先提取 IP, 提取不到则整格
    ips = split_ips(text)
    return ips if ips else [text]


# ================================================================= Excel 解析
HEADER_SCAN_ROWS = 12
HEADER_HINTS = ("IP", "序号", "序列", "主机名", "操作系统", "涉及业务系统",
                "设备类型", "软件名称", "网络类型", "开放端口", "厂商", "域名")


def read_header(ws, fixed_row=None):
    """返回 (表头行号, {表头名: 列号}); fixed_row 指定时直接用该行"""
    if fixed_row:
        names = {}
        for c in range(1, ws.max_column + 1):
            text = normalize_text(ws.cell(fixed_row, c).value)
            if text:
                names.setdefault(text, c)
        return fixed_row, names

    best = (-1, None, {})
    for r in range(1, min(ws.max_row, HEADER_SCAN_ROWS) + 1):
        names, filled = {}, 0
        for c in range(1, ws.max_column + 1):
            text = normalize_text(ws.cell(r, c).value)
            if not text:
                continue
            filled += 1
            names.setdefault(text, c)
        if filled < 2:
            continue
        score = sum(1 for k in HEADER_HINTS if any(k in n for n in names)) + filled / 100.0
        if score > best[0]:
            best = (score, r, names)
    return best[1], best[2]


def locate_column(names, candidates, fuzzy=True):
    for cand in candidates:
        for name, col in names.items():
            if name == cand:
                return col
    if fuzzy:
        for cand in candidates:
            for name, col in names.items():
                if cand in name:
                    return col
    return None


def col_letter(idx):
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def scan_workbook(path, hdr_row_fixed=None):
    """扫描整个工作簿: 返回 [{name, hdr_row, headers, max_col, max_row}]"""
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        hdr_row, names = read_header(ws, hdr_row_fixed)
        sheets.append({
            "name": ws.title,
            "hdr_row": hdr_row or 1,
            "headers": names,
            "max_col": ws.max_column,
            "max_row": ws.max_row,
            "has_header": bool(names),
        })
    wb.close()
    return sheets


def pick_default(headers, keywords, fallback_first=True):
    """按关键词在表头名列表中挑一个默认值

    fallback_first=False 时, 关键词全部未命中则返回 None,
    避免把「序号/序列」这类无关列误当成回填列(会造成整列数据填错)。
    """
    ordered = sorted(headers.items(), key=lambda kv: kv[1])
    for kw in keywords:
        for name, col in ordered:
            if kw == name:
                return name
    for kw in keywords:
        for name, col in ordered:
            if kw in name:
                return name
    return (ordered[0][0] if ordered else None) if fallback_first else None


# 回填列候选: 从最具体到最宽松
TARGET_KEYWORDS = [
    "涉及业务系统/设备", "涉及系统/设备", "涉及业务系统（设备）", "涉及业务系统(设备)",
    "涉及业务系统", "涉及系统", "所属业务系统", "所属系统", "归属系统",
    "业务系统", "系统名称", "涉及业务",
]


def pick_target(headers):
    """挑一个最像「业务系统」的回填列; 挑不到返回 None(调用方改用新建列)"""
    hit = pick_default(headers, TARGET_KEYWORDS, fallback_first=False)
    if hit:
        return hit
    for name, col in sorted(headers.items(), key=lambda kv: kv[1]):
        if "涉及" in name and "系统" in name:
            return name
    return None


# ================================================================= 匹配核心
def build_lookup(inv_path, sheet_names, key_header, val_header,
                 mode="auto", delims=DEFAULT_DELIMS, log=None):
    """构建 {键: [返回值, ...]}, 各 sheet 按表头名各自定位列"""
    def _log(msg):
        if log:
            log(msg)

    wb = openpyxl.load_workbook(inv_path, data_only=True)
    index, detail = {}, []
    for ws in wb.worksheets:
        if sheet_names and ws.title not in sheet_names:
            continue
        hdr_row, names = read_header(ws)
        if not hdr_row:
            _log("  [跳过] %-22s 未识别到表头" % ws.title)
            continue
        kc = locate_column(names, [key_header])
        vc = locate_column(names, [val_header])
        if not kc or not vc:
            _log("  [跳过] %-22s 缺少「%s」或「%s」列"
                 % (ws.title, key_header, val_header))
            continue

        rows, hits = 0, 0
        for r in range(hdr_row + 1, ws.max_row + 1):
            keys = extract_keys(ws.cell(r, kc).value, mode, delims)
            if not keys:
                continue
            val = normalize_text(ws.cell(r, vc).value)
            if not val:
                continue
            rows += 1
            for k in keys:
                bucket = index.setdefault(k, [])
                if val not in bucket:
                    bucket.append(val)
                    hits += 1
        detail.append((ws.title, hdr_row, kc, vc, rows))
        _log("  %-22s 表头行%-2d 查找列%-2d(%s) 返回列%-2d(%s) 有效行%-3d 新增键%-3d"
             % (ws.title, hdr_row, kc, col_letter(kc), vc, col_letter(vc), rows, hits))
    wb.close()
    return index, detail


def run_match_v2(cfg, log=None):
    """cfg 为配置字典, 返回结果字典"""
    def _log(msg):
        if log:
            log(msg)

    up_path = cfg["upload_path"]
    inv_path = cfg["inventory_path"]
    mode = cfg.get("mode", "auto")
    delims = cfg.get("delims", DEFAULT_DELIMS)
    fill = cfg.get("unmatched_fill", "")

    wb = openpyxl.load_workbook(up_path, data_only=True)
    ws = wb[cfg["upload_sheet"]]
    hdr_row = int(cfg.get("upload_hdr_row") or 1)
    _, names = read_header(ws, hdr_row)

    src_col = locate_column(names, [cfg["src_header"]]) or 1
    tgt_name = cfg["target_header"]
    tgt_col = locate_column(names, [tgt_name])
    is_new_col = tgt_col is None
    if is_new_col:
        tgt_col = max(ws.max_column, max(names.values())) + 1

    _log("台账: %s  [工作表]%s  [表头行]%d  [源列]%s(%s)  [回填列]%s(%s%s)"
         % (os.path.basename(up_path), ws.title, hdr_row, cfg["src_header"],
            col_letter(src_col), tgt_name, col_letter(tgt_col),
            "，新建" if is_new_col else ""))

    inv_sheets = cfg.get("inventory_sheets") or []
    index, detail = build_lookup(inv_path, inv_sheets, cfg["key_header"],
                                 cfg["value_header"], mode, delims, _log)
    _log("资产清单: %s  命中 %d 个工作表, 索引 %d 个键"
         % (os.path.basename(inv_path), len(detail), len(index)))

    # 展示列
    columns = [(c, n) for n, c in sorted(names.items(), key=lambda kv: kv[1])]
    if is_new_col:
        columns.append((tgt_col, tgt_name))
    columns = sorted(columns, key=lambda t: t[0])

    rows_data, unmatched = [], []
    for r in range(hdr_row + 1, ws.max_row + 1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, ws.max_column + 1)):
            continue
        src_raw = ws.cell(r, src_col).value
        keys = extract_keys(src_raw, mode, delims)
        vals = []
        for k in keys:
            for v in index.get(k, []):
                if v not in vals:
                    vals.append(v)
        matched = bool(vals)
        if not matched:
            unmatched.append(normalize_text(src_raw) or "(空)")
        values = []
        for col, _n in columns:
            if col == tgt_col:
                values.append("、".join(vals) if matched else fill)
            else:
                v = ws.cell(r, col).value
                values.append("" if v is None else normalize_text(v))
        rows_data.append({
            "row": r, "keys": keys, "src_raw": normalize_text(src_raw),
            "values": values, "matched": matched, "result": vals,
        })

    return {
        "upload_path": up_path, "sheet_name": ws.title, "hdr_row": hdr_row,
        "columns": columns, "src_col": src_col, "target_col": tgt_col,
        "target_name": tgt_name, "is_new_col": is_new_col,
        "rows": rows_data, "unmatched": unmatched, "index": index,
        "detail": detail, "cfg": cfg,
    }


def export_result_v2(result, out_path):
    """保留原格式导出: 只写回填列(新建列时同时写表头)"""
    out_path = os.path.abspath(out_path)
    wb = openpyxl.load_workbook(result["upload_path"])
    ws = wb[result["sheet_name"]]
    tc = result["target_col"]
    if result["is_new_col"]:
        ws.cell(result["hdr_row"], tc).value = result["target_name"]
    fill = result["cfg"].get("unmatched_fill", "")
    for item in result["rows"]:
        ws.cell(item["row"], tc).value = "、".join(item["result"]) if item["matched"] else fill
    wb.save(out_path)
    return out_path


# =================================================================== 命令行
MODE_LABELS = {
    "智能（优先提取IP，无IP则整格）": "auto",
    "仅提取 IP（一格多IP自动拆分）": "ip",
    "整格精确匹配": "cell",
    "按分隔符拆分匹配": "split",
}


def cli_main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    upload, inv = argv[0], argv[1]
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(upload))[0]
    out = argv[2] if len(argv) > 2 else os.path.join(
        os.path.dirname(os.path.abspath(upload)), "%s_已匹配_%s.xlsx" % (base, stamp))

    up_sheets = scan_workbook(upload)
    inv_sheets = scan_workbook(inv)
    src = pick_default(up_sheets[0]["headers"], ["IP", "IP地址"])
    key = "IP地址"
    val = pick_default(_union_headers(inv_sheets), TARGET_KEYWORDS)
    tgt = pick_target(up_sheets[0]["headers"])

    cfg = {
        "upload_path": upload, "inventory_path": inv,
        "upload_sheet": up_sheets[0]["name"], "upload_hdr_row": up_sheets[0]["hdr_row"],
        "src_header": src, "key_header": key, "value_header": val,
        "target_header": tgt if tgt else "涉及业务系统/设备",
        "inventory_sheets": [], "mode": "auto", "delims": DEFAULT_DELIMS,
        "unmatched_fill": "",
    }
    result = run_match_v2(cfg, log=print)
    export_result_v2(result, out)

    total = len(result["rows"])
    ok = sum(1 for r in result["rows"] if r["matched"])
    print("\n" + "=" * 88)
    print("匹配结果  共 %d 行 / 已匹配 %d 行 / 未匹配 %d 行" % (total, ok, total - ok))
    print("列映射: 台账[%s] -> 清单[%s] 取值[%s] -> 回填[%s]"
          % (src, key, val, cfg["target_header"]))
    print("=" * 88)
    for r in result["rows"]:
        print("  行%-3d %-20s -> %s" % (
            r["row"], r["src_raw"][:20],
            "、".join(r["result"]) if r["matched"] else "【未匹配】"))
    if result["unmatched"]:
        print("-" * 88)
        print("未匹配: " + "、".join(result["unmatched"]))
    print("=" * 88)
    print("已导出: %s" % out)
    return 0


def _union_headers(sheets):
    out = {}
    for s in sheets:
        for n, c in s["headers"].items():
            out.setdefault(n, c)
    return out


# =================================================================== 自绘控件
def _round_rect(cv, x1, y1, x2, y2, r, **kw):
    """Canvas 圆角矩形(多边形平滑近似)"""
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


class RoundBtn(tk.Canvas):
    """圆角按钮: primary 实心蓝 / normal 白底描边; 支持 state=disabled"""

    def __init__(self, master, text, command=None, kind="normal",
                 width=None, font=FONT_UI, padx=18):
        self._font = tkfont.Font(font=font)
        tw = self._font.measure(text)
        w = width if width else tw + padx * 2
        h = 32
        super().__init__(master, width=w, height=h, bg=master.cget("bg")
                         if hasattr(master, "cget") else C_BG,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._text, self._command, self._kind = text, command, kind
        self._state = "normal"
        self._bw, self._bh = w, h   # 布局前 winfo_width()==1, 必须用构造尺寸绘制
        # 注意: 不能用 self._w/_h, 会覆盖 tkinter 内置的窗口路径属性
        self._pal = {
            "primary": (C_ACCENT, "#ffffff", C_ACCENT_H, C_ACCENT_P, C_ACCENT),
            "normal":  (C_BTN, C_TEXT, C_BTN_H, C_BTN_P, C_CTRL_BD),
        }[kind]
        self._draw()
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Button-1>", self._click)
        self.bind("<Enter>", lambda e: self._state == "normal" and self._draw(hover=True))
        self.bind("<Leave>", lambda e: self._state == "normal" and self._draw())
        self.bind("<space>", self._click)

    def _draw(self, hover=False):
        self.delete("all")
        bg, fg, hov, press, bd = self._pal
        if self._state == "disabled":
            bg, fg, bd = "#f5f5f5", "#a6a6a6", "#e3e3e3"
        elif hover:
            bg, fg = hov, fg
        _round_rect(self, 1, 1, self._bw - 2, self._bh - 2,
                    8, fill=bg, outline=bd)
        self.create_text(self._bw / 2, self._bh / 2,
                         text=self._text, fill=fg, font=self._font)

    def _click(self, _e=None):
        if self._state == "normal" and self._command:
            self._command()

    def configure(self, **kw):
        if "state" in kw:
            self._state = kw.pop("state")
            self.configure(cursor="hand2" if self._state == "normal" else "arrow")
            self._draw()
        if "text" in kw:
            self._text = kw.pop("text")
            self._draw()
        if kw:
            super().configure(**kw)

    def cget(self, key):
        if key == "state":
            return self._state
        if key == "text":
            return self._text
        return super().cget(key)

    state = property(lambda self: self._state)


class GridView(tk.Frame):
    """Canvas 自绘表格: 单元格边框线 + 行底色 + 垂直/水平自由滚动"""

    HDR_H = 30
    ROW_H = 26

    def __init__(self, parent):
        super().__init__(parent, bg=C_PANEL, bd=0, highlightthickness=0)
        self.font = tkfont.Font(family="Microsoft YaHei UI", size=9)
        self.hfont = tkfont.Font(family="Microsoft YaHei UI", size=9, weight="bold")
        self.headers, self.rows, self.bgrows, self.center = [], [], [], []
        self.widths, self.sel_row = [], -1
        self._busy = False
        self._sr = None          # 上次设置的 scrollregion, 用于阻断重绘反馈循环

        self.canvas = tk.Canvas(self, bg=C_PANEL, highlightthickness=0, bd=0)
        self.vs = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hs = ttk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.canvas.configure(yscrollcommand=self._yscroll,
                              xscrollcommand=self._xscroll)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vs.grid(row=0, column=1, sticky="ns")
        self.hs.grid(row=1, column=0, sticky="ew")
        # 右下角补块: 让底边滚动轴视觉上通栏
        self.corner = tk.Frame(self, bg="#eef0f4", bd=0, highlightthickness=0)
        self.corner.grid(row=1, column=1, sticky="nsew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.canvas.bind("<Configure>", lambda e: self._draw())
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

    # ---- 数据
    def set_data(self, headers, rows, row_tags=None, center_cols=(), max_w=420):
        self.headers = [str(h) for h in headers]
        self.rows = [[("" if v is None else str(v)) for v in r] for r in rows]
        self.bgrows = [(row_tags or [None] * len(rows))[i]
                       for i in range(len(rows))]
        self.center = set(center_cols)
        self.sel_row = -1
        self._measure(max_w)
        self._draw()

    def clear(self):
        self.headers, self.rows, self.bgrows, self.center = [], [], [], set()
        self.sel_row = -1
        self._draw()

    def _measure(self, max_w):
        self.widths = []
        for c, h in enumerate(self.headers):
            col = [r[c] for r in self.rows if c < len(r)] + [h]
            f = self.hfont if c == 0 or True else self.font
            w = max((self.font.measure(t) for t in col), default=0)
            w = max(w, self.hfont.measure(h)) + 26
            self.widths.append(int(max(76, min(max_w, w))))

    # ---- 绘制 (只画可视区, 表头冻结)
    def _draw(self):
        if self._busy:
            return
        self._busy = True
        try:
            self._render()
        finally:
            self._busy = False

    def _set_sr(self, box):
        """只在 scrollregion 真正变化时才写回。

        否则会形成死循环: configure(scrollregion) -> 触发 scrollcommand
        -> _draw -> 再次 configure(scrollregion) -> ... 界面永久卡死。
        """
        if self._sr != box:
            self._sr = box
            self.canvas.configure(scrollregion=box)

    def _render(self):
        cv = self.canvas
        cv.delete("all")
        if not self.headers:
            self._set_sr((0, 0, 1, 1))
            return
        widths, n = self.widths, len(self.rows)
        W = sum(widths) + 1
        H = self.ROW_H * n + 1
        self._set_sr((0, 0, W, H))

        vw = max(cv.winfo_width(), 1)
        vh = max(cv.winfo_height(), 1)
        top = cv.canvasy(0)              # 可视区顶部对应的内容坐标
        bot = top + vh
        left = cv.canvasx(0)
        right = left + vw
        r0 = max(0, int(top // self.ROW_H))
        r1 = min(n, int(bot // self.ROW_H) + 2)

        # 行底色 / 选中行
        for i in range(r0, r1):
            y0 = i * self.ROW_H
            tag = self.bgrows[i] if i < len(self.bgrows) else None
            if tag:
                cv.create_rectangle(1, y0, W - 1, y0 + self.ROW_H, width=0,
                                    fill=C_OK_ROW if tag == "ok" else C_NO_ROW)
            if i == self.sel_row:
                cv.create_rectangle(1, y0, W - 1, y0 + self.ROW_H,
                                    fill=C_SEL, width=0)

        # 单元格文字 (仅可视行 × 可视列)
        lh = self.font.metrics("linespace")
        x = 1
        for c, w in enumerate(widths):
            if x > right or x + w < left:
                x += w
                continue
            for i in range(r0, r1):
                if c >= len(self.rows[i]):
                    continue
                txt = self.rows[i][c]
                if self.font.measure(txt) > w - 22:
                    while txt and self.font.measure(txt + "…") > w - 22:
                        txt = txt[:-1]
                    txt += "…"
                cy = i * self.ROW_H + (self.ROW_H - lh) / 2
                cx = x + w / 2 if c in self.center else x + 12
                cv.create_text(cx, cy, text=txt, font=self.font, anchor="n",
                               fill=C_TEXT)
            x += w

        # 网格线 (仅可视范围)
        x = 0
        for w in list(widths) + [None]:
            x += w if w else 0
            if left - 2 <= x <= right + 2:
                cv.create_line(min(x, W - 1), top, min(x, W - 1), bot,
                               fill=C_LINE)
        for i in range(r0, r1 + 1):
            y = i * self.ROW_H
            cv.create_line(left, y, min(right, W), y, fill=C_LINE)
        cv.create_line(left, H - 1, min(right, W), H - 1, fill=C_LINE)

        # 冻结表头: 始终画在可视区顶部, 数据行从下面滚过
        hl = self.hfont.metrics("linespace")
        cv.create_rectangle(left, top, right, top + self.HDR_H,
                            fill=C_HEAD, width=0)
        cv.create_line(left, top + self.HDR_H, right, top + self.HDR_H,
                       fill=C_CTRL_BD)
        x = 1
        for c, w in enumerate(widths):
            if x > right or x + w < left:
                x += w
                continue
            cv.create_text(x + 12, top + (self.HDR_H - hl) / 2,
                           text=self.headers[c], font=self.hfont,
                           fill=C_TEXT, anchor="nw")
            x += w
        cv.create_rectangle(left, top, min(right, W) - 1, min(bot, H) - 1,
                            outline=C_LINE)

    # ---- 交互
    def _on_click(self, e):
        if not self.rows:
            return
        cy = self.canvas.canvasy(e.y)
        if cy < 0:
            return
        row = int(cy // self.ROW_H)
        if 0 <= row < len(self.rows):
            self.sel_row = row
            self._draw()

    def _yview(self, *a):
        self.canvas.yview(*a)

    def _xview(self, *a):
        self.canvas.xview(*a)

    def _yscroll(self, first, last):
        self.vs.set(first, last)
        self._draw()

    def _xscroll(self, first, last):
        self.hs.set(first, last)
        self._draw()

    def _bind_wheel(self, _e=None):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_wheel_h)

    def _unbind_wheel(self, _e=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Shift-MouseWheel>")

    def _on_wheel(self, e):
        self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def _on_wheel_h(self, e):
        self.canvas.xview_scroll(-1 if e.delta > 0 else 1, "units")


# =================================================================== 主界面
class AssetMatcherApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self._set_window_icon()
        self.configure(bg=C_BG)
        self.geometry("1280x840")
        self.minsize(1040, 700)

        self.upload_path = tk.StringVar()
        self.inv_path = tk.StringVar(
            value=DEFAULT_INVENTORY if os.path.exists(DEFAULT_INVENTORY) else "")

        self.up_sheet = tk.StringVar()
        self.up_hdr_row = tk.StringVar(value="1")
        self.src_header = tk.StringVar()
        self.inv_sheet = tk.StringVar(value=ALL_SHEETS)
        self.key_header = tk.StringVar()
        self.val_header = tk.StringVar()
        self.target_header = tk.StringVar()
        self.mode_label = tk.StringVar(value=list(MODE_LABELS)[0])
        self.fill_value = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="就绪 — 选择台账与资产清单，指定列后开始匹配")

        self.up_sheets, self.inv_sheets = [], []
        self.result = None
        self.last_out_dir = None
        self._last_target = ""
        self.match_logs = []

        self._build_style()
        self._build_ui()

        if self.inv_path.get():
            self._load_inventory(self.inv_path.get())

    # -------------------------------------------------------------- 样式
    def _set_window_icon(self):
        """窗口图标: 标题栏 + 任务栏。优先多尺寸 ico, 缺失时退回 png"""
        if os.path.exists(ICON_ICO):
            try:
                self.iconbitmap(default=ICON_ICO)
                return
            except Exception:
                pass
        if os.path.exists(ICON_PNG):
            try:
                self._icon_png = tk.PhotoImage(file=ICON_PNG)
                self.iconphoto(True, self._icon_png)
            except Exception:
                pass

    def _put_header_logo(self, parent, size=40):
        """标题左侧显示程序 logo (logo_ui.png, 与 exe 图标同一图源)"""
        if not os.path.exists(LOGO_UI):
            return
        try:
            img = tk.PhotoImage(file=LOGO_UI)
        except Exception:
            return
        self._logo_img = img                      # 持有引用, 防止被回收
        tk.Label(parent, image=img, bd=0, bg=C_BG).pack(side="left", padx=(0, 10))

    def _build_style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        # 输入框 / 下拉框
        st.configure("TEntry", fieldbackground=C_PANEL, foreground=C_TEXT,
                     bordercolor=C_CTRL_BD, lightcolor=C_PANEL, darkcolor=C_PANEL,
                     borderwidth=1, padding=(6, 5), relief="flat",
                     insertcolor=C_TEXT)
        st.map("TEntry", bordercolor=[("focus", C_ACCENT)],
               lightcolor=[("focus", C_ACCENT)], darkcolor=[("focus", C_ACCENT)])
        st.configure("TSpinbox", fieldbackground=C_PANEL, foreground=C_TEXT,
                     bordercolor=C_CTRL_BD, lightcolor=C_PANEL, darkcolor=C_PANEL,
                     borderwidth=1, padding=(6, 5), relief="flat", arrowsize=11)
        st.map("TSpinbox", bordercolor=[("focus", C_ACCENT)])
        st.configure("TCombobox", fieldbackground=C_PANEL, foreground=C_TEXT,
                     background=C_PANEL, bordercolor=C_CTRL_BD, lightcolor=C_PANEL,
                     darkcolor=C_PANEL, borderwidth=1, padding=(6, 5), relief="flat",
                     arrowsize=13)
        st.map("TCombobox", bordercolor=[("focus", C_ACCENT), ("readonly", C_CTRL_BD)],
               fieldbackground=[("readonly", C_PANEL)],
               background=[("readonly", C_PANEL)],
               foreground=[("readonly", C_TEXT)])
        # 表格
        st.configure("Treeview", background=C_PANEL, fieldbackground=C_PANEL,
                     foreground=C_TEXT, rowheight=26, borderwidth=0, font=FONT_UI_S)
        st.configure("Treeview.Heading", background=C_HEAD, foreground=C_TEXT,
                     font=FONT_BOLD, relief="flat", borderwidth=0, padding=(8, 6))
        st.map("Treeview.Heading",
               background=[("active", "#f0f0f0")],
               relief=[("active", "flat")])
        st.map("Treeview", background=[("selected", C_SEL)],
               foreground=[("selected", C_TEXT)])
        # 滚动条: 浅灰凹槽 + 深灰滑块, 保证清晰可辨
        for orient in ("Vertical", "Horizontal"):
            st.configure("%s.TScrollbar" % orient,
                         background="#98a1ad", troughcolor="#eef0f4",
                         bordercolor=C_LINE, darkcolor="#98a1ad",
                         lightcolor="#98a1ad", arrowcolor="#57606f",
                         borderwidth=1, arrowsize=15, relief="flat")
            st.map("%s.TScrollbar" % orient,
                   background=[("active", "#6f7887"), ("pressed", "#57606f")],
                   arrowcolor=[("active", "#2f3640")])

    def _card(self, parent, pad=14, radius=12, follow=True):
        """柔和圆角卡片: Canvas 画圆角矩形 + 内嵌白色 Frame
        follow=True  高度跟随内容(紧凑配置卡)
        follow=False 填满父容器分配的空间(表格主面板)"""
        cv = tk.Canvas(parent, bg=C_BG, highlightthickness=0, bd=0)
        inner = tk.Frame(cv, bg=C_PANEL, bd=0)
        win = cv.create_window(pad, pad, window=inner, anchor="nw")

        def redraw(_e=None):
            cw, ch = cv.winfo_width(), cv.winfo_height()
            if cw < 8:
                return
            if follow:
                ih = inner.winfo_reqheight()
                if ih > 4:
                    cv.configure(height=ih + 2 * pad)
                cv.itemconfigure(win, width=max(cw - 2 * pad, 1))
            else:
                cv.itemconfigure(win, width=max(cw - 2 * pad, 1),
                                 height=max(inner.winfo_reqheight(), ch - 2 * pad))
            cv.delete("card")
            _round_rect(cv, 0, 0, cw - 1, ch - 1, radius,
                        fill=C_PANEL, outline=C_LINE, tag="card")
            cv.tag_lower("card")

        cv.bind("<Configure>", redraw)
        inner.bind("<Configure>", redraw)
        return cv, inner

    def _btn(self, parent, text, command, kind="normal", width=13):
        """圆角按钮: 主按钮实心蓝, 次按钮白底描边"""
        return RoundBtn(parent, text, command, kind=kind)

    def _entry(self, parent, var, width=None):
        e = ttk.Entry(parent, textvariable=var, font=FONT_UI_S, width=width)
        return e

    def _combo(self, parent, var, values=None, cmd=None, width=None):
        cb = ttk.Combobox(parent, textvariable=var, values=values or [],
                          state="readonly", font=FONT_UI_S, width=width)
        if cmd:
            cb.bind("<<ComboboxSelected>>", cmd)
        return cb

    def _field(self, parent, r, label, widget, note=None):
        """堆叠式表单行: 小标签在上, 控件在下"""
        tk.Label(parent, text=label, bg=C_PANEL, fg=C_MUTED, font=FONT_SM,
                 anchor="w").grid(row=r, column=0, columnspan=2, sticky="w",
                                  pady=(0, 3))
        widget.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        if note:
            tk.Label(parent, text=note, bg=C_PANEL, fg=C_MUTED, font=FONT_SM,
                     anchor="w").grid(row=r + 2, column=0, columnspan=2,
                                      sticky="w", pady=(0, 10))
            return r + 3
        return r + 2

    def _group_card(self, parent, col, index, title, subtitle):
        frame, body = self._card(parent, pad=14)
        frame.grid(row=0, column=col, sticky="nsew",
                   padx=(0 if col == 0 else 8, 8 if col < 2 else 0))
        head = tk.Frame(body, bg=C_PANEL)
        head.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        tk.Label(head, text=str(index), bg=C_ACCENT, fg="#ffffff",
                 font=("Microsoft YaHei UI", 8, "bold"), width=2, height=1,
                 padx=4, pady=1).pack(side="left", padx=(0, 8))
        box = tk.Frame(head, bg=C_PANEL)
        box.pack(side="left")
        tk.Label(box, text=title, bg=C_PANEL, fg=C_TEXT, font=FONT_BOLD,
                 anchor="w").pack(anchor="w")
        tk.Label(box, text=subtitle, bg=C_PANEL, fg=C_MUTED, font=FONT_SM,
                 anchor="w").pack(anchor="w", pady=(1, 0))
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        return body

    # -------------------------------------------------------------- 布局
    def _build_ui(self):
        # ---- 底部状态栏(先 pack, 避免被上方组件挤掉)
        sbar = tk.Frame(self, bg=C_BG)
        sbar.pack(fill="x", side="bottom")
        tk.Frame(sbar, bg=C_LINE, height=1).pack(fill="x")
        self.bottom = tk.Label(sbar, text="就绪 — 选择台账与资产清单，指定列后开始匹配",
                               bg=C_BG, fg=C_MUTED, font=FONT_SM, anchor="w")
        self.bottom.pack(fill="x", padx=20, pady=6)

        # ---- 标题栏
        head = tk.Frame(self, bg=C_BG)
        head.pack(fill="x", padx=18, pady=(14, 12))
        self._put_header_logo(head)
        tk.Label(head, text="IT资产台账匹配工具", bg=C_BG, fg=C_TEXT,
                 font=FONT_TITLE).pack(side="left")
        tk.Label(head, text=VERSION, bg=C_BG, fg=C_MUTED,
                 font=FONT_SM).pack(side="left", padx=(8, 0), pady=(5, 0))

        # ---- 配置区: 三列卡片
        cfg = tk.Frame(self, bg=C_BG)
        cfg.pack(fill="x", padx=18, pady=(0, 10))
        for i in range(3):
            cfg.columnconfigure(i, weight=1, uniform="col")

        def frow(body, r, label, widget, note=None):
            """紧凑表单行: 小标签在左, 控件在右"""
            tk.Label(body, text=label, bg=C_PANEL, fg=C_MUTED, font=FONT_SM,
                     anchor="w").grid(row=r, column=0, sticky="w", pady=4,
                                      padx=(0, 10))
            widget.grid(row=r, column=1, sticky="ew", pady=4)
            if note:
                tk.Label(body, text=note, bg=C_PANEL, fg=C_MUTED, font=FONT_SM,
                         anchor="w").grid(row=r + 1, column=1, sticky="w",
                                          pady=(0, 2))
                return r + 2
            return r + 1

        def ffile(body, r, label, var, cmd):
            holder = tk.Frame(body, bg=C_PANEL)
            holder.columnconfigure(0, weight=1)
            self._entry(holder, var).grid(row=0, column=0, sticky="ew")
            self._btn(holder, "浏览", cmd, kind="primary",
                      width=6).grid(row=0, column=1, padx=(6, 0))
            return frow(body, r, label, holder)

        # 卡片 ①  待匹配台账
        b1 = self._group_card(cfg, 0, 1, "待匹配台账", "要被回填的文件")
        r = ffile(b1, 1, "文件", self.upload_path, self.choose_upload)
        self.cb_up_sheet = self._combo(b1, self.up_sheet, [], self.on_up_sheet)
        r = frow(b1, r, "工作表", self.cb_up_sheet)
        spin = ttk.Spinbox(b1, from_=1, to=20, width=6, textvariable=self.up_hdr_row,
                           command=self.on_hdr_row, font=FONT_UI_S)
        spin.bind("<Return>", lambda e: self.on_hdr_row())
        r = frow(b1, r, "表头行", spin)
        self.cb_src = self._combo(b1, self.src_header, [])
        frow(b1, r, "源列（比对）", self.cb_src)

        # 卡片 ②  资产清单
        b2 = self._group_card(cfg, 1, 2, "资产清单", "提供匹配依据的文件")
        r = ffile(b2, 1, "文件", self.inv_path, self.choose_inventory)
        self.cb_inv_sheet = self._combo(b2, self.inv_sheet, [ALL_SHEETS], self.on_inv_sheet)
        r = frow(b2, r, "工作表", self.cb_inv_sheet)
        self.cb_key = self._combo(b2, self.key_header, [])
        r = frow(b2, r, "查找列（比对）", self.cb_key)
        self.cb_val = self._combo(b2, self.val_header, [])
        frow(b2, r, "返回列（取值）", self.cb_val)

        # 卡片 ③  输出设置
        b3 = self._group_card(cfg, 2, 3, "输出设置", "回填到哪里、怎么匹配")
        self.cb_target = self._combo(b3, self.target_header, [], self.on_target_selected)
        r = frow(b3, 1, "回填到", self.cb_target, note="末项「新建列…」可新建")
        self.cb_mode = self._combo(b3, self.mode_label, list(MODE_LABELS))
        r = frow(b3, r, "匹配模式", self.cb_mode)
        frow(b3, r, "未匹配填充", self._entry(b3, self.fill_value, width=18),
             note="留空则不填")

        # ---- 底部面板: 操作按钮 + 匹配后台账表格 (一个通栏卡片)
        tcard, tinner = self._card(self, pad=2, radius=12, follow=False)
        tcard.pack(fill="both", expand=True, padx=18, pady=(0, 10))

        bar = tk.Frame(tinner, bg=C_PANEL)
        bar.pack(fill="x", padx=12, pady=(10, 8))
        self._btn(bar, "开始匹配", self.do_match, kind="primary", width=12).pack(side="left")
        self.btn_export = self._btn(bar, "导出最终台账", self.do_export, width=12)
        self.btn_export.pack(side="left", padx=8)
        self.btn_export.configure(state="disabled")
        self._btn(bar, "打开输出目录", self.open_out_dir, width=11).pack(side="left")
        tk.Label(bar, textvariable=self.status_text, bg=C_PANEL, fg=C_MUTED,
                 font=FONT_UI_S, anchor="e").pack(side="right", expand=True, fill="x")
        tk.Frame(tinner, bg=C_LINE, height=1).pack(fill="x", padx=12)

        thead = tk.Frame(tinner, bg=C_PANEL)
        thead.pack(fill="x", padx=12, pady=(8, 8))
        tk.Label(thead, text="匹配后的台账内容", bg=C_PANEL, fg=C_TEXT,
                 font=FONT_BOLD).pack(side="left")
        self.tree_info = tk.Label(thead, text="", bg=C_PANEL, fg=C_MUTED, font=FONT_SM)
        self.tree_info.pack(side="left", padx=(10, 0), pady=(1, 0))
        legend = tk.Frame(thead, bg=C_PANEL)
        legend.pack(side="right")
        for color, txt in ((C_OK_DOT, "已匹配"), (C_NO_DOT, "未匹配")):
            tk.Label(legend, text="", bg=color, width=1, height=1
                     ).pack(side="left", padx=(0, 5), pady=(3, 0))
            tk.Label(legend, text=txt, bg=C_PANEL, fg=C_MUTED,
                     font=FONT_SM).pack(side="left", padx=(0, 12))
        tk.Frame(tinner, bg=C_LINE, height=1).pack(fill="x")

        wrap = tk.Frame(tinner, bg=C_PANEL)
        wrap.pack(fill="both", expand=True)
        self.table = GridView(wrap)
        self.table.pack(fill="both", expand=True)

        self.placeholder = tk.Label(
            wrap, text="选择文件并点击「开始匹配」，结果将显示在这里",
            bg=C_PANEL, fg=C_MUTED, font=FONT_UI)
        self.placeholder.place(relx=0.5, rely=0.5, anchor="center")

    # -------------------------------------------------------------- 文件加载
    def choose_upload(self):
        path = filedialog.askopenfilename(
            title="选择待匹配台账",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.upload_path.set(path)
            self._load_upload(path)

    def choose_inventory(self):
        path = filedialog.askopenfilename(
            title="选择资产清单",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.inv_path.set(path)
            self._load_inventory(path)

    def _load_upload(self, path):
        try:
            self.up_sheets = scan_workbook(path)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        usable = [s for s in self.up_sheets if s["has_header"]] or self.up_sheets
        self.cb_up_sheet["values"] = [s["name"] for s in usable]
        first = usable[0]
        self.up_sheet.set(first["name"])
        self.up_hdr_row.set(str(first["hdr_row"]))
        self._refresh_upload_columns()

    def _refresh_upload_columns(self):
        name = self.up_sheet.get()
        sheet = next((s for s in self.up_sheets if s["name"] == name), None)
        if not sheet:
            return
        try:
            row = int(self.up_hdr_row.get())
        except ValueError:
            row = sheet["hdr_row"]
        wb = openpyxl.load_workbook(self.upload_path.get(), data_only=True)
        ws = wb[name]
        _r, names = read_header(ws, row)
        wb.close()
        sheet["headers"] = names
        sheet["hdr_row"] = row
        cols = [n for n, c in sorted(names.items(), key=lambda kv: kv[1])]

        self.cb_src["values"] = cols
        if not self.src_header.get() or self.src_header.get() not in cols:
            self.src_header.set(pick_default(names, ["IP", "IP地址"]) or (cols[0] if cols else ""))

        tgt_vals = cols + [NEW_COL]
        self.cb_target["values"] = tgt_vals
        if not self.target_header.get() or self.target_header.get() not in tgt_vals:
            # 识别不到「业务系统」类列时默认新建列, 绝不误填到序号等无关列
            self.target_header.set(pick_target(names) or NEW_COL)

    def _load_inventory(self, path):
        try:
            self.inv_sheets = scan_workbook(path)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        self.inv_sheet.set(ALL_SHEETS)
        self._refresh_inventory_columns()

    def _refresh_inventory_columns(self):
        sel = self.inv_sheet.get()
        sheets = self.inv_sheets if sel == ALL_SHEETS else \
            [s for s in self.inv_sheets if s["name"] == sel]
        union = _union_headers(sheets)
        cols = [n for n, c in sorted(union.items(), key=lambda kv: kv[1])]
        self.cb_key["values"] = cols
        self.cb_val["values"] = cols
        cur_k = self.key_header.get()
        cur_v = self.val_header.get()
        if not cur_k or cur_k not in cols:
            self.key_header.set(pick_default(union, ["IP地址", "IP"]) or (cols[0] if cols else ""))
        if not cur_v or cur_v not in cols:
            self.val_header.set(
                pick_default(union, TARGET_KEYWORDS) or (cols[-1] if cols else ""))

    # -------------------------------------------------------------- 事件
    def on_up_sheet(self, _e=None):
        self._refresh_upload_columns()

    def on_hdr_row(self):
        self._refresh_upload_columns()

    def on_inv_sheet(self, _e=None):
        self._refresh_inventory_columns()

    def on_target_selected(self, _e=None):
        if self.target_header.get() != NEW_COL:
            self._last_target = self.target_header.get()
            return
        name = simpledialog.askstring("新建列", "请输入新列名：", parent=self)
        vals = [v for v in self.cb_target["values"] if v != NEW_COL]
        if name and name.strip():
            name = name.strip()
            if name not in vals:
                vals.append(name)
            self.cb_target["values"] = vals + [NEW_COL]
            self.target_header.set(name)
            self._last_target = name
        else:
            self.target_header.set(self._last_target or (vals[-1] if vals else ""))

    # -------------------------------------------------------------- 匹配
    def do_match(self):
        up, inv = self.upload_path.get().strip(), self.inv_path.get().strip()
        if not up or not os.path.exists(up):
            messagebox.showinfo("提示", "请先选择待匹配台账文件")
            return
        if not inv or not os.path.exists(inv):
            messagebox.showinfo("提示", "请先选择资产清单文件")
            return
        if not self.src_header.get():
            messagebox.showinfo("提示", "请选择源列")
            return
        if not self.key_header.get() or not self.val_header.get():
            messagebox.showinfo("提示", "请选择资产清单的查找列与返回列")
            return
        if not self.target_header.get():
            messagebox.showinfo("提示", "请选择回填列")
            return

        self.status_text.set("正在匹配…")
        self.update_idletasks()

        sel = self.inv_sheet.get()
        cfg = {
            "upload_path": up, "inventory_path": inv,
            "upload_sheet": self.up_sheet.get(),
            "upload_hdr_row": self.up_hdr_row.get(),
            "src_header": self.src_header.get(),
            "key_header": self.key_header.get(),
            "value_header": self.val_header.get(),
            "target_header": self.target_header.get(),
            "inventory_sheets": [] if sel == ALL_SHEETS else [sel],
            "mode": MODE_LABELS.get(self.mode_label.get(), "auto"),
            "delims": DEFAULT_DELIMS,
            "unmatched_fill": self.fill_value.get(),
        }
        try:
            logs = []
            result = run_match_v2(cfg, log=logs.append)
        except Exception as exc:
            messagebox.showerror("匹配失败", "%s" % exc)
            self.status_text.set("匹配失败")
            return

        self.result = result
        self.match_logs = logs
        self._render_table(result)

        total = len(result["rows"])
        ok = sum(1 for r in result["rows"] if r["matched"])
        self.status_text.set("共 %d 行 · 已匹配 %d · 未匹配 %d" % (total, ok, total - ok))
        self.btn_export.configure(state="normal")
        self.last_out_dir = os.path.dirname(os.path.abspath(up))
        un = "、".join(result["unmatched"][:12])
        more = " …等%d个" % (len(result["unmatched"]) - 12) if len(result["unmatched"]) > 12 else ""
        self.bottom.configure(text=
            "映射: 台账[%s] → 清单[%s] → 取值[%s] → 回填[%s(%s%s)]  |  清单命中 %d 个表，索引 %d 个键%s"
            % (cfg["src_header"], cfg["key_header"], cfg["value_header"],
               cfg["target_header"], col_letter(result["target_col"]),
               "，新建" if result["is_new_col"] else "",
               len(result["detail"]), len(result["index"]),
               ("  |  未匹配: " + un + more) if un else "  |  全部匹配成功"))

        if result["unmatched"]:
            messagebox.showwarning(
                "存在未匹配项",
                "共 %d 条未在资产清单中找到对应值:\n\n%s\n\n已按「未匹配填充」规则处理。"
                % (len(result["unmatched"]), "\n".join(result["unmatched"][:20])))

    def _render_table(self, result):
        self.placeholder.place_forget()
        total = len(result["rows"])
        ok = sum(1 for r in result["rows"] if r["matched"])
        self.tree_info.configure(
            text="%d 行 · 命中 %d · 未命中 %d" % (total, ok, total - ok))
        cols = list(result["columns"]) + [(-1, "匹配状态")]
        headers = [name + (" ◀回填" if c == result["target_col"] else "")
                   for c, name in cols]
        rows, tags = [], []
        for item in result["rows"]:
            rows.append(list(item["values"]) +
                        ["已匹配" if item["matched"] else "未匹配"])
            tags.append("ok" if item["matched"] else "no")
        self.table.set_data(headers, rows, row_tags=tags,
                            center_cols=(len(cols) - 1,))

    def do_export(self):
        if not self.result:
            messagebox.showinfo("提示", "请先执行匹配")
            return
        base = os.path.splitext(os.path.basename(self.result["upload_path"]))[0]
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default = os.path.join(os.path.dirname(os.path.abspath(self.result["upload_path"])),
                               "%s_已匹配_%s.xlsx" % (base, stamp))
        out = filedialog.asksaveasfilename(
            title="导出最终台账", defaultextension=".xlsx",
            initialfile=os.path.basename(default),
            initialdir=os.path.dirname(default),
            filetypes=[("Excel 文件", "*.xlsx")])
        if not out:
            return
        try:
            path = export_result_v2(self.result, out)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.last_out_dir = os.path.dirname(path)
        self.status_text.set("已导出: %s" % os.path.basename(path))
        if messagebox.askyesno("导出成功", "已导出:\n%s\n\n是否在资源管理器中打开？" % path):
            try:
                os.startfile(os.path.dirname(path))
            except Exception:
                pass

    def open_out_dir(self):
        d = self.last_out_dir or os.path.dirname(
            os.path.abspath(self.upload_path.get() or "."))
        if d and os.path.isdir(d):
            try:
                os.startfile(d)
            except Exception:
                pass


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ("--cli", "-c"):
        sys.exit(cli_main(sys.argv[2:]))
    _bind_app_id()
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    AssetMatcherApp().mainloop()


if __name__ == "__main__":
    main()
