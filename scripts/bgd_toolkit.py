# -*- coding: utf-8 -*-
"""图纸和变更量自动化处理 —— VBA 宏的 Python 等价实现（不运行任何宏）。

对应 VBA 模块 ``图纸和变更量自动化处理.bas``（V2026.11.08）的全部业务功能：

    CreateBianGengdan  -> create_change_order()   生成「*-变更单打印」表
    Xls2Gldon          -> export_gldon()          生成「广联达导入」表
    Xls2Shenjims       -> export_shenji()         生成「神机导入」表
    LoadX              -> load_quota_db()         载入替换表 + 套价数据库（走 JSON 缓存）
    XChange            -> XchangeDB.simplify()    精简物料描述（给导出用）
    XChangeAll         -> simplify_worksheet()    精简整表 + 去掉结尾下划线
    ReplaceX           -> XchangeDB.normalize()   分隔符规范化
    HrdeX              -> match_quota()           华润定额匹配（含单位换算）
    ifhasNext          -> _has_column()           模板是否含第 n 列

VBA 里的 ``UpdateThisMod`` / ``UpdateCode``（从共享盘导入 .bas 自我更新）属于
VBA 专用维护手段，Python 版不需要，故不移植；模块的入口菜单 ``Xls2Shenji``
（MsgBox 询问广联达/神机）改为命令行参数 ``gldon`` / ``shenji``。

所有 Excel 写操作仍走 Excel COM——.xls 的合并单元格、打印区域、分页符、
字体加粗等格式只能靠 Excel 本体才能保证与宏输出逐格一致。
"""
from __future__ import print_function

import math
import os
import sys
import time

try:
    import win32com.client
except ImportError:                                   # pragma: no cover
    win32com = None

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xchange_db

# ---- Excel 常量
XL_UP = -4162
XL_DOWN = -4161
XL_TO_LEFT = -4159
XL_FORMULAS = -4123
XL_WHOLE = 1
XL_PART = 2
XL_NORMAL_VIEW = 1
XL_PAGE_BREAK_PREVIEW = 2

TEMPLATE_SHEET = "@变更单EXCEL模板"
GDLD_SHEET = "广联达导入"
SHENJI_SHEET = "神机导入"
CVER = "VER - 2026.11.08 by ACECKT"


# ------------------------------------------------------------------ 小工具

def _num(v, default=0.0):
    """VBA 里 ``Range.Value`` 的空值是 Empty（参与运算即为 0），这里统一成数值。"""
    if v is None or v == "":
        return default
    if isinstance(v, bool):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(" ", ""))
    except ValueError:
        return default


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return str(v)


def _is_numeric(v):
    if v is None or v == "":
        return False
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    try:
        float(str(v))
        return True
    except ValueError:
        return False


def _find(ws, what, after_cell=None):
    """等价 ``Cells.Find(What)``；找不到返回 None（VBA 里是 Nothing）。

    pywin32 的后期绑定对 Find 的关键字参数支持不好（会静默返回 None），
    必须按 9 个位置参数完整传入。
    """
    try:
        cell = after_cell if after_cell is not None else ws.Cells(1, 1)
        return ws.Cells.Find(what, cell, XL_FORMULAS, XL_PART, 1, 1, False, False, False)
    except Exception:
        return None


def _replace(rng, what, repl, lookat=XL_PART):
    """等价 ``Range.Replace(What, Replacement)``（同样必须位置传参）。"""
    try:
        rng.Replace(what, repl, lookat, 1, False, False, False, False)
        return True
    except Exception:
        return False


def _find_from_row(ws, what, row):
    """从 (row, 1) 之后开始查找，语义同 ``Find(What, Cells(row, 1))``。"""
    return _find(ws, what, ws.Cells(row, 1))


# ------------------------------------------------------------------ 会话

class Session(object):
    """一次 Excel COM 会话：打开工作簿 → 干活 → 保存并退出。"""

    def __init__(self, path, visible=False):
        if win32com is None:
            raise RuntimeError("需要 pywin32：pip install pywin32")
        self.path = os.path.abspath(path)
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = visible
        self.app.DisplayAlerts = False
        self.app.AutomationSecurity = 1          # 允许宏运行（本脚本自己并不调用宏）
        self.wb = self.app.Workbooks.Open(self.path)
        self.app.CalculateFullRebuild()

    def __enter__(self):
        return self

    def source_sheet(self, name=None):
        """确定源工作表：显式名字 > 「2-汇总」 > 当前活动表。"""
        if name:
            return self.wb.Worksheets(name)
        for s in self.wb.Worksheets:
            if s.Name.strip() == "2-汇总":
                return s
        return self.wb.ActiveSheet

    def save(self, out_path=None):
        if out_path:
            self.wb.SaveAs(os.path.abspath(out_path))
        else:
            self.wb.Save()

    def __exit__(self, exc_type, exc, tb):
        try:
            self.wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            self.app.Quit()
        finally:
            self.wb = None
            self.app = None
        return False


def load_quota_db(force_rebuild=False, path=None):
    """等价 ``LoadX``：载入「替换表」与「数据库」。源文件更新时自动重建 JSON 缓存。"""
    return xchange_db.get_db(path=path, force_rebuild=force_rebuild, quiet=True)


# ------------------------------------------------------------------ 定额匹配

def match_quota(db, rows, code_ver=""):
    """等价 ``HrdeX``：把名称/单位换成定额名称、定额单位、综合单价、定额编码。

    ``rows`` 是 7 列的二维列表（1=名称 2=单位 3=数量 4=综合单价 5=编码 6=备注 7=暂定），
    就地修改并返回同一个列表。
    """
    col_desc = db.col("物料描述")
    col_code = db.col("编码" + str(code_ver)) if code_ver != "" else -1
    col_qname = db.col("定额名称")
    col_qunit = db.col("定额单位")
    col_price = db.col("综合单价")
    if min(col_desc, col_qname, col_qunit, col_price) < 0:
        return rows                                  # 数据库缺列：原样返回

    for row in rows:
        name = _text(row[0])
        key = db.normalize(name.lower())
        hit = db.quota_row(key)
        if hit <= 0:
            row[4] = "未匹配"
            continue

        unit_of_quota = _text(db.value(hit, "定额单位"))
        data_unit = _text(row[1])
        qty = _num(row[2])

        # ---- 单位换算（完全照 VBA 的先后与短路关系）
        converted = False
        if not converted and (unit_of_quota == "米" or unit_of_quota.lower() == "m"):
            if data_unit == "个" and any(k in name for k in ("出地", "出土", "0.5米/个", "0.5m/个")):
                row[2] = qty * 0.5
                converted = True
        if not converted and unit_of_quota == "个":
            if data_unit == "米" or data_unit.lower() == "m":
                if any(k in name for k in ("穿墙", "0.3米/个", "0.3m/个")):
                    row[2] = qty / 0.3
                    converted = True
                # VBA 这一行是 (A And B) Or C Or D，照抄其优先级
                if (not converted and any(k in name for k in ("穿楼",))
                        or any(k in name for k in ("0.2米/个", "0.2m/个"))):
                    row[2] = qty / 0.2
                    converted = True
        if not converted and unit_of_quota == "副":
            if data_unit in ("片", "个", "只"):
                row[2] = qty / 2.0
                converted = True
        if not converted and unit_of_quota in ("m", "米"):
            if data_unit in ("根", "个", "只"):
                row[2] = qty / 2.0
                converted = True

        row[0] = _text(db.value(hit, "定额名称"))
        row[1] = unit_of_quota
        row[3] = db.value(hit, "综合单价")
        code = _text(db.value(hit, "编码" + str(code_ver)) if col_code >= 0 else "")
        row[4] = "" if code == "'-" else ("'" + code)
    return rows


# ------------------------------------------------------------------ 整表精简

def simplify_worksheet(db, ws):
    """等价 ``XChangeAll``：按替换表精简整表，再去掉每格结尾的一个下划线。"""
    for what, repl in db.replace_rules:
        if not what:
            continue
        _replace(ws.Cells, what, repl)

    used = ws.UsedRange
    vals = used.Value
    if not isinstance(vals, tuple):
        text = _text(vals)
        if text.endswith("_"):
            used.Cells(1, 1).Value = text[:-1]
        return

    rows_n, cols_n = used.Rows.Count, used.Columns.Count
    changed = []
    new_rows = []
    for r in range(rows_n):
        line = list(vals[r])
        for c in range(cols_n):
            s = line[c]
            if s is None:
                continue
            if isinstance(s, str) and s.endswith("_"):
                line[c] = s[:-1]
                changed.append((r + 1, c + 1))
        new_rows.append(tuple(line))

    if not changed:
        return
    # 整块写回在 pywin32 后期绑定下会静默失效，统一逐格写（值语义与 VBA 相同）
    r0, c0 = used.Row, used.Column
    for (r, c) in changed:
        ws.Cells(r0 + r - 1, c0 + c - 1).Value = new_rows[r - 1][c - 1]


# ------------------------------------------------------------------ 广联达导入

def export_gldon(db, wb, src):
    """等价 ``Xls2Gldon``：按 C/D/I/J 生成「广联达导入」表。"""
    last = src.UsedRange.Rows.Count
    data = src.Range("C3:J%d" % last).Value
    if not isinstance(data, tuple):
        data = ((data,),)

    drawing, add, reduce = [], [], []
    simplified = ""
    for row in data:                             # data 就是第 3 行起的 C..J
        name = _text(row[0])
        unit = row[1]
        qty_d = _num(row[6])
        qty_c = _num(row[7])
        if qty_d != 0 or qty_c > 0 or qty_c < 0:
            simplified = db.simplify(name)
        if qty_d != 0:
            drawing.append([simplified, unit, qty_d, "", "", simplified, ""])
        if qty_c > 0:
            add.append([simplified, unit, qty_c, "", "", simplified, ""])
        elif qty_c < 0:
            reduce.append([simplified, unit, qty_c, "", "", simplified, ""])

    if len(drawing) + len(add) + len(reduce) == 0:
        return 0

    _clear_or_add_sheet(wb, GDLD_SHEET, src)

    var_r2 = src.Range("R2").Value
    var_s2 = src.Range("S2").Value
    code_ver_r2 = _text(var_r2)

    dst = wb.Worksheets(GDLD_SHEET)
    for col, head in zip(range(1, 7), ("名称", "单位", "数量", "综合单价", "编码", "备注")):
        dst.Cells(1, col).Formula = head
    dst.Cells(1, 7).Formula = ""

    r = 2
    if drawing:
        if var_r2 != "" and var_r2 is not None and _is_numeric(var_s2):
            match_quota(db, drawing, _text(var_r2))
        elif var_s2 != "" and var_s2 is not None and _is_numeric(var_s2):
            match_quota(db, drawing, _text(var_s2))
        else:
            match_quota(db, drawing, "2022")
        dst.Cells(r, 1).Formula = "图纸工程量"
        dst.Cells(r, 7).Formula = ""
        _write_block(dst, r + 1, drawing)
        r = r + len(drawing) + 1
    if add or reduce:
        dst.Cells(r, 1).Formula = "设计变更"
        dst.Cells(r, 7).Formula = ""
        r += 1
    if add:
        match_quota(db, add, code_ver_r2)
        _write_block(dst, r, add)
        r += len(add)
    if reduce:
        match_quota(db, reduce, code_ver_r2)
        _write_block(dst, r, reduce)
        r += len(reduce)

    dst.UsedRange.EntireColumn.AutoFit()
    return len(drawing) + len(add) + len(reduce)


# ------------------------------------------------------------------ 神机导入

def export_shenji(db, wb, src):
    """等价 ``Xls2Shenjims``：生成「神机导入」表（不匹配定额库）。"""
    last = src.UsedRange.Rows.Count
    data = src.Range("C3:J%d" % last).Value
    if not isinstance(data, tuple):
        data = ((data,),)

    drawing, add, reduce = [], [], []
    for row in data:                             # data 就是第 3 行起的 C..J
        name = _text(row[0])
        unit = row[1]
        qty_d = _num(row[6])
        qty_c = _num(row[7])
        if qty_d != 0:
            drawing.append([name, unit, qty_d, qty_d, "cx", "dx4", "调拨价"])
        if qty_c > 0:
            add.append([name, unit, qty_c, qty_c, "cx", "dx4", "调拨价"])
        elif qty_c < 0:
            reduce.append([name, unit, qty_c, qty_c, "cx", "dx4", "调拨价"])

    if len(drawing) + len(add) + len(reduce) == 0:
        return 0

    _clear_or_add_sheet(wb, SHENJI_SHEET, src)

    dst = wb.Worksheets(SHENJI_SHEET)
    for col, head in zip(range(1, 8),
                         ("项目名称", "单位", "计算公式", "工程量", "安装特项", "市政特项", "项目编码")):
        dst.Cells(1, col).Formula = head

    r = 2
    if drawing:
        dst.Cells(r, 1).Formula = "图纸工程量"
        dst.Cells(r, 7).Formula = "*"
        _write_block(dst, r + 1, drawing)
        r = r + len(drawing) + 1
    if add or reduce:
        dst.Cells(r, 1).Formula = "设计变更"
        dst.Cells(r, 7).Formula = "*"
        r += 1
    if add:
        _write_block(dst, r, add)
        r += len(add)
    if reduce:
        _write_block(dst, r, reduce)
        r += len(reduce)

    dst.UsedRange.EntireColumn.AutoFit()
    return len(drawing) + len(add) + len(reduce)


def _write_block(ws, row, rows):
    """把二维列表写到 A{row} 起的 7 列区域。

    注意：pywin32 后期绑定下 ``Range.Value = 二维元组`` 会**静默不生效**
    （单元格仍是空），所以这里逐格写，语义与 VBA 的
    ``Range("A" & r).Resize(n, 7) = arr`` 一致。
    """
    if not rows:
        return
    for i, data in enumerate(rows):
        for j, value in enumerate(data):
            ws.Cells(row + i, j + 1).Value = value


def _clear_or_add_sheet(wb, name, after_src):
    """目标表存在则清空，不存在则新建（VBA 同名逻辑）。"""
    exists = False
    for s in wb.Worksheets:
        if s.Name == name:
            s.UsedRange.Clear()
            exists = True
            break
    if not exists:
        wb.Worksheets.Add(None, after_src)       # (Before, After)：同样是位置传参
        wb.ActiveSheet.Name = name


# ------------------------------------------------------------------ 变更单

def _template_marks(ws):
    """扫描「@变更单EXCEL模板」里的 模板n开始/结束 标记，返回 (列表, 列数)。"""
    items = []
    for i in range(1, 11):
        s = _find(ws, "模板%d开始" % i)
        e = _find(ws, "模板%d结束" % i)
        if s is None or e is None:
            break
        items.append({
            "start": s.Row, "end": e.Row, "end_col": e.Column,
            "cap": int(_num(ws.Cells(s.Row + 1, e.Column).Value)),
            "cols": 0,
        })
    return items


def _has_column(ws, n):
    """等价 ``ifhasNext``：模板中是否存在「[第n列开始]」标记。"""
    return _find(ws, "[第%d列开始]" % n) is not None


def _dst_name():
    """沿用 VBA 的命名：Right((YYMMDD&HHMMSS)/19901015, 3) & "-变更单打印"。"""
    concat = time.strftime("%y%m%d%H%M%S")
    num = float(concat) / 19901015.0
    frac = num - int(num)
    n = int(round(frac * 1000))
    return "%d-变更单打印" % n


def _collect_items(src):
    """读 2-汇总 的变更复制区：J>0 → K/L 为增加，J<0 → M/N 为减少。"""
    last = src.UsedRange.Rows.Count
    data = src.Range("C3:N%d" % last).Value
    if not isinstance(data, tuple):
        data = ((data,),)
    add, reduce = [], []
    for row in data:                             # data 就是第 3 行起的 C..N
        qty = _num(row[7])                       # J 列
        if qty > 0:
            add.append([row[8], row[9]])          # K / L
        elif qty < 0:
            reduce.append([row[10], row[11]])     # M / N
    if add:
        add.insert(0, ["增加材料：", ""])
    if reduce:
        reduce.insert(0, ["减少材料：", ""])
    return add, reduce


def create_change_order(db, wb, src, name=None, out_path=None):
    """等价 ``CreateBianGengdan``：生成「*-变更单打印」。

    返回 (表名, 写入了多少条材料) 或 (None, 0)（无变更量）。
    """
    add, reduce = _collect_items(src)
    n_add, n_reduce = len(add), len(reduce)
    if n_add + n_reduce == 0:
        return None, 0

    dst_name = name or _dst_name()

    # 同名表先删掉
    for s in wb.Worksheets:
        if s.Name == dst_name:
            s.Delete()
            break

    tpl = wb.Worksheets(TEMPLATE_SHEET)
    tpl.Visible = -1                             # 取消隐藏
    before = [s.Name for s in wb.Worksheets]
    # 注意：pywin32 后期绑定下 Copy(After=...) 的关键字会静默失效（不复制），
    # 必须按位置传 (Before, After)
    tpl.Copy(None, src)                          # 复制到源表后面
    new_sheet = None
    for s in wb.Worksheets:                      # 定位副本：名字是新增的那个
        if s.Name not in before:
            new_sheet = s
            break
    if new_sheet is None:
        new_sheet = wb.ActiveSheet
    new_sheet.Name = dst_name

    ws = wb.Worksheets(dst_name)
    marks = _template_marks(ws)
    total = n_add + n_reduce

    # ---- 选模板：容量「够用且最小」的那一套（同容量取靠前的），落在奇数序号
    #      VBA 的判断是 A 组「本模板装得下 且 比当前选的更小」先命中，
    #      所以效果等价于「容量 ≥ 条数的最小容量」；都装不下时退到容量最大的那套。
    fits = [(i, m["cap"]) for i, m in enumerate(marks, 1) if m["cap"] >= total]
    n_tpl = 0
    if fits:
        best_cap = min(cap for _, cap in fits)
        idx = min(i for i, cap in fits if cap == best_cap)
    else:
        best_cap = max(m["cap"] for m in marks)
        idx = min(i for i, m in enumerate(marks, 1) if m["cap"] == best_cap)
    n_tpl = idx if idx % 2 == 1 else idx - 1
    ti = marks[n_tpl - 1]
    block = ti["end"] - ti["start"] + 1

    # ---- 删多余模板
    if n_tpl >= 3:
        ws.Rows("1:%d" % (ti["start"] - 1)).Delete()
        drop = ti["start"] - 1
        for m in marks[n_tpl - 1:]:
            m["start"] -= drop
            m["end"] -= drop
        ti = marks[n_tpl - 1]
    last_row = ws.Cells(65536, ti["end_col"]).End(XL_UP).Row
    if n_tpl + 2 <= len(marks):
        ws.Rows("%d:%d" % (marks[n_tpl]["end"] + 1, last_row)).Delete()
    ws.Columns(ti["end_col"]).Delete(XL_TO_LEFT)

    # ---- 分页：一页装不下就复制整页
    pages = int(math.ceil(float(total) / ti["cap"])) if ti["cap"] else 1
    if total > ti["cap"]:
        for _ in range(pages - 1):
            ws.Rows("1:%d" % (block * 2)).Copy()
            ws.Rows("1:1").Insert(XL_DOWN, 0)     # (Shift, CopyOrigin)
            wb.Application.CutCopyMode = False

    # ---- 模板列数（删除后仍存在的列标记数）
    for k in range(1, 6):
        if _has_column(ws, k):
            marks[n_tpl - 1]["cols"] += 1
            marks[n_tpl]["cols"] += 1
    cols = marks[n_tpl - 1]["cols"] or 1

    # ---- 打印设置
    app = wb.Application
    try:
        app.PrintCommunication = False
    except Exception:
        pass
    ws.PageSetup.PrintArea = ws.Range(
        ws.Cells(1, 1), ws.Cells(block * 2 * pages, ti["end_col"] - 1)).Address
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 0

    last_print_row = block * 2 * pages
    targets = []
    for i in range(1, pages + 1):
        targets.append(block * (2 * i - 1) + 1)
        targets.append(block * 2 * i + 1)

    try:
        app.ActiveWindow.View = XL_PAGE_BREAK_PREVIEW   # 分页符必须先切分页预览
        existing = {}
        for j in range(1, ws.HPageBreaks.Count + 1):
            try:
                existing[ws.HPageBreaks(j).Location.Row] = j
            except Exception:
                pass
        for row in targets:
            if row in existing:
                continue
            if row > last_print_row:              # 打印区域之外的断点 Excel 不接受
                continue
            try:
                ws.HPageBreaks.Add(ws.Range("A%d" % row))
                existing[row] = True
            except Exception:
                pass
    finally:
        try:
            app.ActiveWindow.View = XL_NORMAL_VIEW
        except Exception:
            pass
        try:
            app.PrintCommunication = True
        except Exception:
            pass

    # ---- 填充
    _fill(ws, ti, marks[n_tpl] if n_tpl < len(marks) else ti, block, pages, cols,
          add, n_add, reduce, n_reduce)

    # ---- 清理标记 / 精简描述 / 去公式
    for k in range(1, 4):
        for tag in ("[第%d列开始]", "[第%d列结束]", "[数量%d]"):
            _replace(ws.Cells, tag % k, "")
    simplify_worksheet(db, ws)
    _replace(ws.Cells, "`", "_")
    ws.UsedRange = ws.UsedRange.Value

    if out_path:
        wb.SaveAs(os.path.abspath(out_path))
    return dst_name, total


def _remaining(cols, cap, col, n_reduce):
    """VBA 那两处「剩余容量是否够装减少材料」的公式。"""
    if cols <= 0:
        return False
    return ((cols - 1) * 2 + cap) / float(cols) * (cols - col) - (cols - col - 1) * 2 >= n_reduce


def _fill(ws, ta, tb, block, pages, cols, add, n_add, reduce, n_reduce):
    """把增加/减少材料按列、按页填进两套模板（申请单 + 工程量变更单）。"""
    b_fill_reduce = n_add <= 1                   # 只有减少材料时直接从减少开始
    idx = 1
    page = 0
    tip = ""

    while True:
        page += 1
        start_a = block * 2 * (page - 1) + 1
        end_a = block * 2 * (page - 1) + block
        start_b = end_a + 1
        end_b = end_a + block - 1
        ta["start"], ta["end"] = start_a, end_a
        tb["start"], tb["end"] = start_b, end_b

        next_page = False
        col = 0
        while True:
            col += 1
            head = _find_from_row(ws, "[第%d列开始]" % col, ta["start"])
            if head is None:
                return                            # 没有这一列 → 全部结束
            qty_mark = _find_from_row(ws, "[数量%d]" % col, ta["start"])
            if qty_mark is None:
                return
            row = head.Row
            c1 = head.Column
            c2 = qty_mark.Column

            if tip:
                target = (row - ta["start"]) + tb["start"]
                ws.Cells(row, c1).Value = tip
                ws.Cells(row, c1).Font.Bold = True
                ws.Cells(target, c1).Value = tip
                ws.Cells(target, c1).Font.Bold = True
                ws.Cells(row, c2).Value = ""
                ws.Cells(target, c2).Value = ""
                tip = ""
                row += 1

            action = None
            while True:
                target = (row - ta["start"]) + tb["start"]
                if not b_fill_reduce:
                    if idx <= n_add:
                        marker = _text(ws.Cells(row, c1).Value)
                        if marker != "[第%d列结束]" % col or idx == n_add:
                            _put_item(ws, row, target, c1, c2, add[idx - 1])
                            row += 1
                            if idx >= n_add:
                                b_fill_reduce = True
                                idx = 1
                                if _remaining(cols, ta["cap"], col, n_reduce):
                                    action = "col"
                                    break
                            else:
                                idx += 1
                        else:
                            if col < cols:
                                _put_tip(ws, row, target, c1, c2, "(接右侧)")
                                tip = "接左侧增加材料："
                                action = "col"
                                break
                            _put_tip(ws, row, target, c1, c2, "(接第%d页)" % (page + 1))
                            tip = "接第%d页增加材料：" % page
                            action = "page"
                            break
                    else:
                        b_fill_reduce = True
                        idx = 1
                        if _remaining(cols, ta["cap"], col, n_reduce):
                            action = "col"
                            break
                else:
                    if idx <= n_reduce:
                        marker = _text(ws.Cells(row, c1).Value)
                        if marker != "[第%d列结束]" % col or idx == n_reduce:
                            _put_item(ws, row, target, c1, c2, reduce[idx - 1])
                            row += 1
                            if idx >= n_reduce:
                                action = "done"
                                break
                            idx += 1
                        else:
                            if col < cols:
                                _put_tip(ws, row, target, c1, c2, "(接右侧)")
                                tip = "接左侧减少材料："
                                action = "col"
                                break
                            _put_tip(ws, row, target, c1, c2, "(接第%d页)" % (page + 1))
                            tip = "接第%d页减少材料：" % page
                            action = "page"
                            break
                    else:
                        action = "done"
                        break

            if action == "done":
                return
            if action == "page":
                next_page = True
                break
        if not next_page:
            return
        if page > pages:                          # 保险：不应超出算好的页数
            return


def _put_item(ws, row, target, c1, c2, item):
    """同一张单据的两处（申请单 + 工程量变更单）写同样的名称与数量。"""
    name, qty = item[0], item[1]
    name = "" if name is None else name
    qty = "" if qty is None else qty
    ws.Cells(row, c1).Value = name
    ws.Cells(target, c1).Value = name
    ws.Cells(row, c2).Value = qty
    ws.Cells(target, c2).Value = qty
    if qty == "":
        ws.Cells(row, c1).Font.Bold = True
        ws.Cells(target, c1).Font.Bold = True


def _put_tip(ws, row, target, c1, c2, text):
    ws.Cells(row, c1).Value = text
    ws.Cells(target, c1).Value = text
    ws.Cells(row, c2).Value = ""
    ws.Cells(target, c2).Value = ""


# ------------------------------------------------------------------ CLI

USAGE = """用法：
  python bgd_toolkit.py bgd    <材料平差.xls> [--sheet 2-汇总] [--name 表名] [--out 另存.xls]
  python bgd_toolkit.py gldon  <材料平差.xls> [--sheet 2-汇总] [--out 另存.xls]
  python bgd_toolkit.py shenji <材料平差.xls> [--sheet 2-汇总] [--out 另存.xls]
  python bgd_toolkit.py all    <材料平差.xls> [--target gldon+shenji+bgd] [--out 另存.xls]
  python bgd_toolkit.py db     [--rebuild] [--xchange <XchangeSheet.xls>]

说明：
  · 默认就地保存：不加 --out 时，结果写回原 .xls（与宏一样在同一个工作簿里加表）
  · 加 --no-save 只预览、不落盘
  · 套价数据库走 scripts/xchange_cache.json 缓存（加载 ~10ms）；源 .xls 路径、大小或
    修改时间任意一项变了，下一次运行会自动重建缓存
  · db 子命令查看/重建缓存，--xchange 指定源数据库路径
"""


def _opt(args, name, default=None):
    if name in args:
        i = args.index(name)
        if i + 1 < len(args):
            return args[i + 1]
    return default


def main(argv):
    if not argv or argv[0] in ("-h", "--help"):
        print(USAGE)
        return 0 if argv else 2
    cmd = argv[0]
    args = argv[1:]

    # ---- db：只看/重建套价数据库缓存，不需要变更单文件
    if cmd == "db":
        src = _opt(args, "--xchange")
        if "--rebuild" in args:
            db = load_quota_db(force_rebuild=True, path=src)
            st = db.stats()
            print("[OK] 缓存已重建：%d 条替换规则 / %d 行定额 / %d 条分隔符"
                  % (st["rules"], st["db_rows"], st["separators"]))
        else:
            db = load_quota_db(path=src)
            st = db.stats()
            print("[OK] 缓存可用：%d 条替换规则 / %d 行定额 / %d 条分隔符（建于 %s）"
                  % (st["rules"], st["db_rows"], st["separators"], st.get("built_at", "?")))
            print("     源文件：%s" % (db.data.get("source", {}).get("path", "?")))
        return 0

    path = None
    for a in args:
        if a.lower().endswith((".xls", ".xlsx", ".xlsm")):
            path = a
            break
    if not path:
        print("缺少变更单文件（.xls）")
        return 2

    sheet = _opt(args, "--sheet")
    out = _opt(args, "--out")
    name = _opt(args, "--name")
    no_save = "--no-save" in args
    db = load_quota_db(path=_opt(args, "--xchange"))

    with Session(path) as s:
        src = s.source_sheet(sheet)
        if cmd == "bgd":
            dst, n = create_change_order(db, s.wb, src, name=name, out_path=out)
            if not dst:
                print("[INFO]「%s」里没有变更量，未生成变更单" % src.Name)
                return 0
            print("[OK] 已生成「%s」，共 %d 条材料" % (dst, n))
        elif cmd == "gldon":
            n = export_gldon(db, s.wb, src)
            if not n:
                print("[INFO] 没有数据（I 列图纸量、J 列变更量都为 0）")
                return 0
            print("[OK] 已生成「%s」，共 %d 条" % (GDLD_SHEET, n))
        elif cmd == "shenji":
            n = export_shenji(db, s.wb, src)
            if not n:
                print("[INFO] 没有数据（I 列图纸量、J 列变更量都为 0）")
                return 0
            print("[OK] 已生成「%s」，共 %d 条" % (SHENJI_SHEET, n))
        elif cmd == "all":
            targets = (_opt(args, "--target") or "bgd").split("+")
            if "gldon" in targets:
                print("[OK] 广联达导入 %d 条" % export_gldon(db, s.wb, src))
            if "shenji" in targets:
                print("[OK] 神机导入 %d 条" % export_shenji(db, s.wb, src))
            if "bgd" in targets:
                dst, n = create_change_order(db, s.wb, src, name=name)
                print("[OK] 变更单「%s」%d 条" % (dst, n) if dst else "[INFO] 没有变更量")
        else:
            print(USAGE)
            return 2

        if not no_save and not out:
            s.save()
            print("[OK] 已保存：%s" % path)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
