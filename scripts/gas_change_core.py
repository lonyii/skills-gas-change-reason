# -*- coding: utf-8 -*-
"""
燃气工程设计变更单 —— 变更原因读写核心库

知识来源：南昌燃气工程设计变更单历史归档
  E:\\2024\\2024结算相关\\1.材料平差\\设计变更\\hk  (136 份 .xls 变更单)

本模块提供：
  1. ChangeOrder  —— 变更单读写（xlrd 读 / Excel COM 写）
  2. Corpus       —— 历史语料库加载与相似案例检索
  3. ReasonBuilder—— 依据历史规律合成变更原因

环境要求：
  - Python 3.12/3.13
  - xlrd>=2.0 (读 .xls)
  - pywin32 (写 .xls，需本机安装 Excel)
"""

import os
import re
import sys
import glob
import json
import shutil
import tempfile
import time

try:
    import xlrd
except ImportError:  # pragma: no cover
    raise SystemExit("缺少依赖 xlrd，请执行: pip install xlrd")

if hasattr(sys.stdout, "buffer"):
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


# ---------------------------------------------------------------- 常量

#: 历史变更单归档根目录（可多个）
#:
#: 取值优先级：
#:   1) 环境变量 GAS_CHANGE_ARCHIVE（多个路径用 ; 分隔）
#:   2) 本文件同级的 config.json 中的 "archive_roots"（数组）或 "archive_root"（字符串）
#:   3) 下面的默认值（本机路径）
#:
#: 用 `python fill_change_reason.py reindex` 可重建语料缓存。
_DEFAULT_ARCHIVE = [
    r"E:\2024\2024结算相关\1.材料平差\设计变更\hk",
    r"D:\Documents\桌面\历史变更单",
]


def _resolve_archive_roots():
    """解析归档根目录列表：环境变量 > config.json > 默认值。

    显式设置即使路径不存在也予以采纳，这样拼错路径时 doctor 能明确报错，
    而不是静默回退到默认路径。
    """
    roots = []
    env = os.environ.get("GAS_CHANGE_ARCHIVE")
    if env and env.strip():
        roots = [p for p in re.split(r"[;|]", env.replace("\n", ";")) if p.strip()]
    else:
        cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(cfg):
            try:
                import json
                with open(cfg, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                v = data.get("archive_roots")
                if isinstance(v, list):
                    roots = [str(x) for x in v if str(x).strip()]
                elif data.get("archive_root"):
                    roots = [str(data["archive_root"])]
            except Exception:
                pass
    if not roots:
        roots = list(_DEFAULT_ARCHIVE)
    return [os.path.abspath(r.strip()) for r in roots]


ARCHIVE_ROOTS = _resolve_archive_roots()

#: 兼容旧接口：单一根目录（第一个存在的）
ARCHIVE_ROOT = next((r for r in ARCHIVE_ROOTS if os.path.isdir(r)),
                    ARCHIVE_ROOTS[0] if ARCHIVE_ROOTS else "")

#: 变更单打印表名关键字（排除作废/旧版）
PRINT_SHEET_KEY = "变更单打印"

#: 变更原因占位符
PLACEHOLDER = "[变更原因]"

#: 申请单结尾语（314 份语料中 213 份使用）
TAIL_APP = "兹提出设计变更申请单，请予以审批。"

#: 工程量变更单结尾语
TAIL_QTY = "兹提出工程量变更单。"

#: 罕见结尾语（历史上是笔误，勿模仿）
TAIL_APP_RARE = "拟调整如下材料，兹提出变更申请单。"

#: 申请单结尾语的已知变体（按 314 份语料频次排序）
#: 老归档 99% 用标准句；新归档 40% 用「兹提出变更申请单，请予以审批。」
TAIL_APP_VARIANTS = [
    "兹提出设计变更申请单，请予以审批。",   # 213
    "兹提出变更申请单，请予以审批。",       # 72
    "兹提出变更单申请，请予以审批。",       # 15
    "兹提出工程量变更申请单，请予以审批。",  # 4
    "兹提出设计变更清单，请予以审批。",     # 1
]

#: 识别正文是否已自带结尾语（避免重复拼接）
_TAIL_RE = re.compile(r"(兹提出[^。；]*?(?:申请单|变更单|变更清单|工程量变更申请单)[^。；]*?。)\s*$")


def detect_tail(text):
    """识别文本结尾是否已是结尾语；是则返回该结尾语，否则空串。"""
    m = _TAIL_RE.search(str(text or "").strip())
    return m.group(1) if m else ""


# ---------------------------------------------------------------- 工具

def _norm(name):
    """材料名归一化：去括号注释、去规格后缀、小写。

    '20#钢无缝管_D76_4mm_双层熔结环氧粉末防腐' -> '20#钢无缝管_d76'
    '钢制弯头_20#碳钢_DN80_4mm_1.5D_90°'      -> '钢制弯头_dn80_90°'
    """
    s = str(name).strip()
    s = re.sub(r"（[^）]*）", "", s)
    s = re.sub(r"\([^)]*\)", "", s)
    # 去尾部规格说明：_4mm_... / _1.6MPa_...
    s = re.sub(r"_\d+(?:\.\d+)?mm.*$", "", s)
    s = re.sub(r"_\d+(?:\.\d+)?MPa.*$", "", s)
    return s.strip().lower()


def _short_spec(name):
    """截断材料名到规格段，用于正文引用（对齐历史措辞）。

    '钢制单头牙_DN15_外牙_50cm' -> '钢制单头牙_DN15'
    '钢制弯头_DN80_90°'         -> '钢制弯头_DN80_90°'
    '20#钢无缝管_D76'            -> '20#钢无缝管_D76'
    """
    s = str(name).strip()
    parts = s.split("_")
    for i, p in enumerate(parts[1:], start=1):
        if re.match(r"^(DN|dn|D|d)\d+$", p):
            keep = parts[: i + 1]
            if i + 1 < len(parts) and "°" in parts[i + 1]:
                keep.append(parts[i + 1])
            return "_".join(keep)
    return s


def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _norm_name(name):
    """工程名称归一化（用于识别「同一个工程」）：去括号后缀、空格、标点。"""
    s = _clean(name)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = re.sub(r"[\s\-—_、，,·．.]+", "", s)
    s = re.sub(r"(工程|项目)$", "", s)
    return s


# ---------------------------------------------------------------- 变更单读取

class ChangeOrder(object):
    """一个设计变更单工作簿的只读视图。

    典型工作表：'1-初步比对' / '2-汇总' / '587-变更单打印' /
                 '4-汇总打印版 (新版)' / '5-变更估价' / '@变更单EXCEL模板'
    """

    def __init__(self, path):
        self.path = os.path.abspath(path)
        self.name = os.path.basename(self.path)
        self._wb = None
        self._print_sheets = None
        self._print_sheet = None
        self._merged_cache = {}
        self._orders = None

    # -- 基础访问 ------------------------------------------------------

    @property
    def workbook(self):
        if self._wb is None:
            self._wb = xlrd.open_workbook(self.path, formatting_info=True)
        return self._wb

    @property
    def sheet_names(self):
        return [s.name for s in self.workbook.sheets()]

    def sheet(self, name):
        try:
            return self.workbook.sheet_by_name(name)
        except Exception:
            return None

    @property
    def print_sheets(self):
        """全部可作为变更单打印的工作表（排除作废/旧版）。

        一个工作簿常有多张 '*变更单打印' 表（如「（地下）」「（地上）」各一张），
        每张表内还可能纵向排列多张变更单。
        """
        if self._print_sheets is None:
            self._print_sheets = [
                s for s in self.workbook.sheets()
                if PRINT_SHEET_KEY in s.name and "作废" not in s.name and "旧版" not in s.name
            ]
        return self._print_sheets

    @property
    def print_sheet(self):
        """第一张 '*变更单打印' 表（兼容旧接口）。"""
        if self._print_sheet is not None:
            return self._print_sheet
        self._print_sheet = self.print_sheets[0] if self.print_sheets else None
        return self._print_sheet

    def merged_for(self, sh):
        """取某张表的合并区（xlrd 的 merged_cells 未按行排序，勿依赖顺序）。"""
        if sh is None:
            return []
        cached = self._merged_cache.get(sh.name)
        if cached is None:
            cached = list(getattr(sh, "merged_cells", []) or [])
            self._merged_cache[sh.name] = cached
        return cached

    @property
    def merged(self):
        return self.merged_for(self.print_sheet)

    # -- 字段提取 ------------------------------------------------------

    def orders(self):
        """枚举本工作簿中的全部变更单（多 sheet / 同 sheet 多张）。

        一张变更单 = 「设计变更申请单」+「工程量变更单」两个单据。
        版式特征：
          · 申请单：A 列「工程编号」→「设计变更原因」→ 材料区 →「变更预算」
          · 工程量单：A 列「工程量变更单」标题 →「工程编号」→ 整行合并的原因区
        同一 sheet 内出现多个「设计变更原因」即为多张变更单（模板默认 4 组）。
        """
        if self._orders is not None:
            return self._orders
        out = []
        for sh in self.print_sheets:
            anchors = [r for r in range(sh.nrows)
                       if _clean(sh.cell_value(r, 0)) == "设计变更原因"]
            if anchors:
                bounds = [(anchors[i], anchors[i + 1] if i + 1 < len(anchors) else sh.nrows)
                          for i in range(len(anchors))]
            else:
                # 老版式可能只有工程量变更单，没有申请单
                title = next((r for r in range(sh.nrows)
                              if "工程量变更单" in _clean(sh.cell_value(r, 0))), None)
                bounds = [(None, sh.nrows)] if title is not None else []
            for i, (start, end) in enumerate(bounds):
                order = self._order_at(sh, start, end, len(out) + 1, i + 1)
                if order:
                    out.append(order)
        self._orders = out
        return out

    def _order_at(self, sh, app_row, end_row, seq, index):
        """提取单张变更单：原因格位置、编号名称、增减材料、现状。"""
        merged = self.merged_for(sh)
        qty_title = None
        for r in range(0 if app_row is None else app_row + 1,
                       min(end_row, sh.nrows)):
            if "工程量变更单" in _clean(sh.cell_value(r, 0)):
                qty_title = r
                break

        # 申请单原因格：B 列起的合并区
        app_slot = None
        if app_row is not None:
            for (rlo, rhi, clo, chi) in merged:
                if rlo == app_row and clo == 1:
                    app_slot = (sh.name, rlo, clo)
                    break
            if app_slot is None:
                app_slot = (sh.name, app_row, 1)

        # 工程量单原因格：标题行之下、行号最小的整行合并区
        # （xlrd 的 merged_cells 未按行排序，必须取 min(rlo)）
        qty_slot = None
        if qty_title is not None:
            cands = [(rlo, clo) for (rlo, rhi, clo, chi) in merged
                     if clo == 0 and chi >= 5 and qty_title < rlo < end_row]
            if cands:
                rlo, clo = min(cands)
                qty_slot = (sh.name, rlo, clo)

        header_row, code, name = self._order_header(sh, app_row, qty_title)
        if app_row is None and qty_title is None:
            return None

        mat_start = header_row if header_row is not None else (app_row or qty_title or 0)
        mat_end = qty_title if qty_title is not None else end_row
        add, red = self._order_materials(sh, mat_start, mat_end)

        app_val = _clean(sh.cell_value(app_slot[1], app_slot[2])) if app_slot else ""
        qty_val = _clean(sh.cell_value(qty_slot[1], qty_slot[2])) if qty_slot else ""
        has_draft = bool((app_val and app_val != PLACEHOLDER)
                         or (qty_val and qty_val != PLACEHOLDER))
        return {
            "sheet": sh.name,
            "index": index,
            "seq": seq,
            "code": code,
            "name": name,
            "header_row": header_row,
            "app_slot": app_slot,
            "qty_slot": qty_slot,
            "app_addr": _addr(app_slot[1], app_slot[2]) if app_slot else "",
            "qty_addr": _addr(qty_slot[1], qty_slot[2]) if qty_slot else "",
            "add": add,
            "red": red,
            "draft_app": app_val,
            "draft_qty": qty_val,
            "has_draft": has_draft,
            "needs_fill": not has_draft,
        }

    def _order_header(self, sh, app_row, qty_title):
        """该单据的 (工程编号所在行, 工程编号, 工程名称)。

        「工程编号」紧邻「设计变更原因」上方；同 sheet 多张单据时，
        向上扫描遇到上一张单据的「设计变更原因」即停，避免取到别人的编号。
        """
        start = app_row if app_row is not None else (qty_title or 0)
        r = start
        while r >= 0 and start - r <= 30:
            a = _clean(sh.cell_value(r, 0))
            if a == "工程编号":
                return r, _clean(sh.cell_value(r, 2)), _clean(sh.cell_value(r, 7))
            if a == "设计变更原因" and r != app_row:
                break
            r -= 1
        for r in range(start, min(sh.nrows, start + 30)):
            if _clean(sh.cell_value(r, 0)) == "工程编号":
                return r, _clean(sh.cell_value(r, 2)), _clean(sh.cell_value(r, 7))
        return None, "", ""

    def _order_materials(self, sh, start_row, end_row):
        """读取该单据的增/减材料。

        关键：**不能按列判定增/减**。有的模板把「减少材料」标签放在 A 列、
        「增加材料」放在 H 列，也有 A 列整块只有「减少材料」的。必须按标签
        单元格的实际文本决定该侧归属，否则会把减少项误读成增加项。
        """
        add, red = [], []
        side = {0: None, 7: None}
        for r in range(max(0, start_row), min(end_row, sh.nrows)):
            a = _clean(sh.cell_value(r, 0))
            h = _clean(sh.cell_value(r, 7))
            for col, txt in ((0, a), (7, h)):
                t = txt.replace("：", ":")
                if t == "增加材料:":
                    side[col] = "add"
                elif t == "减少材料:":
                    side[col] = "red"
            if a.startswith(("变更预算", "施工单位", "监理单位")):
                break
            for col in (0, 7):
                if side[col] is None:
                    continue
                nm = _clean(sh.cell_value(r, col))
                if not nm or nm.replace("：", ":") in ("增加材料:", "减少材料:"):
                    continue
                if nm.startswith(("变更预算", "施工单位", "监理单位",
                                  "预算人员", "日 期", "签 章")):
                    continue
                item = {"name": nm,
                        "qty": _clean(sh.cell_value(r, 6 if col == 0 else 13)),
                        "norm": _norm(nm)}
                (add if side[col] == "add" else red).append(item)
        return add, red

    # -- 兼容旧接口：全部指向第一张变更单 --------------------------------

    def _first_order(self):
        orders = self.orders()
        return orders[0] if orders else None

    def header(self):
        """返回 (工程编号, 工程名称) —— 第一张变更单。"""
        o = self._first_order()
        return (o["code"], o["name"]) if o else ("", "")

    def materials(self):
        """返回 (增加材料 list, 减少材料 list) —— 第一张变更单。"""
        o = self._first_order()
        return (o["add"], o["red"]) if o else ([], [])

    def reason_slots(self):
        """第一张变更单的原因格位置。

        返回 dict:
          app : (sheet_name, row, col)  申请单「设计变更原因」右侧合并区左上角
          qty : (sheet_name, row, col)  工程量变更单「[变更原因]」合并区左上角
        """
        o = self._first_order()
        out = {}
        if o:
            if o["app_slot"]:
                out["app"] = o["app_slot"]
            if o["qty_slot"]:
                out["qty"] = o["qty_slot"]
        return out

    def current_reason(self):
        """读取第一张变更单已有原因（申请单侧）；无则空串。"""
        o = self._first_order()
        return o["draft_app"] if o else ""

    def needs_fill(self, order=None):
        """True 表示变更原因尚未填写。"""
        o = order if isinstance(order, dict) else self._first_order()
        return bool(o["needs_fill"]) if o else True

    def read_draft(self, order=None):
        """读取两处变更原因单元格的现有内容（用户草稿）。

        返回 dict:
          app : 申请单现状文本（无则空串）
          qty : 工程量变更单现状文本（无则空串）
          app_addr / qty_addr : 对应 A1 地址
          has_draft : 是否存在可作为提示词的草稿
        """
        o = order if isinstance(order, dict) else self._first_order()
        if not o:
            return {"app": "", "qty": "", "app_addr": "", "qty_addr": "", "has_draft": False}
        return {"app": o["draft_app"], "qty": o["draft_qty"],
                "app_addr": o["app_addr"], "qty_addr": o["qty_addr"],
                "has_draft": o["has_draft"]}


# ---------------------------------------------------------------- 历史语料

class Corpus(object):
    """历史变更原因语料库。

    支持多个归档目录（root 可为字符串或列表）。
    """

    def __init__(self, root=None, cache=None):
        if root is None:
            roots = list(ARCHIVE_ROOTS)
        elif isinstance(root, (list, tuple)):
            roots = [os.path.abspath(r) for r in root if str(r).strip()]
        else:
            roots = [os.path.abspath(str(root))]
        self.roots = roots
        self.root = roots[0] if roots else ""
        self.cache = cache
        self.records = []
        if cache and os.path.exists(cache):
            with open(cache, "r", encoding="utf-8") as fh:
                self.records = json.load(fh)
        else:
            self.records = self._scan()
            if cache:
                try:
                    with open(cache, "w", encoding="utf-8") as fh:
                        json.dump(self.records, fh, ensure_ascii=False, indent=1)
                except OSError:
                    pass

    # -- 扫描 ----------------------------------------------------------

    def _scan(self):
        recs = []
        seen = set()
        files = []
        for root in self.roots:
            if not root or not os.path.isdir(root):
                continue
            for f in sorted(glob.glob(os.path.join(root, "**", "*.xls"), recursive=True)):
                if ".workbuddy" in f:
                    continue
                # 写入前的自动备份副本（与用户自建备份目录）不能进语料，
                # 否则同一工程会被重复计数
                if "_变更单备份" in f or ".bak." in os.path.basename(f):
                    continue
                if "备份" in os.path.basename(os.path.dirname(f)):
                    continue
                key = os.path.normcase(os.path.abspath(f))
                if key in seen:
                    continue
                seen.add(key)
                files.append(f)
        for f in files:
            try:
                co = ChangeOrder(f)
                sh = co.print_sheet
            except Exception:
                continue
            if sh is None:
                continue
            try:
                code, name = co.header()
                add, red = co.materials()
                reason = co.current_reason()
            except Exception:
                continue
            recs.append({
                "file": os.path.basename(f),
                "path": f,
                "code": code,
                "name": name,
                "reason": reason,
                "add": add,
                "red": red,
                "add_norm": sorted(set(x["norm"] for x in add)),
                "red_norm": sorted(set(x["norm"] for x in red)),
            })
        return recs

    # -- 检索 ----------------------------------------------------------

    @staticmethod
    def jaccard(a, b):
        A, B = set(a), set(b)
        if not (A | B):
            return 0.0
        return len(A & B) / len(A | B)

    def similar(self, add_norm, red_norm, top=8, min_score=0.0):
        """按材料增删集合相似度返回最相近的历史案例。

        分数 = 0.5 * Jaccard(增加) + 0.5 * Jaccard(减少)
        """
        scored = []
        for r in self.records:
            if not r["reason"] or r["reason"] == PLACEHOLDER:
                continue
            s = (0.5 * self.jaccard(add_norm, r["add_norm"])
                 + 0.5 * self.jaccard(red_norm, r["red_norm"]))
            if s >= min_score:
                scored.append((s, r))
        scored.sort(key=lambda x: -x[0])
        return scored[:top]

    def by_code(self, code):
        """按工程编号找历史记录（同名工程的其他标段可作参照）。"""
        return [r for r in self.records if r["code"] == code]

    def project_match(self, code="", name="", add_norm=None, red_norm=None):
        """找同项目的历史定稿原因（同工程编号 或 同工程名称）。

        返回 (材料相似度, 记录) 或 None。
        编号相同但名称不同（如同一工程的地下段/地上段）不算同一张单据，
        避免把另一段的定稿口径套到这一段。
        """
        add_norm = add_norm or []
        red_norm = red_norm or []
        nm = _norm_name(name)
        cands = []
        for r in self.records:
            if not r["reason"] or r["reason"] == PLACEHOLDER:
                continue
            rn = _norm_name(r.get("name"))
            same = False
            if code and r.get("code") and r["code"] == code:
                same = (not nm or not rn or nm == rn or nm in rn or rn in nm)
            elif nm and rn and (nm == rn or (len(nm) >= 6 and (nm in rn or rn in nm))):
                same = True
            if not same:
                continue
            score = (0.5 * self.jaccard(add_norm, r["add_norm"])
                     + 0.5 * self.jaccard(red_norm, r["red_norm"]))
            cands.append((score, r))
        if not cands:
            return None
        cands.sort(key=lambda x: -x[0])
        return cands[0]

    # -- 润色场景的增强检索 --------------------------------------------

    #: 中文停用词（用于从草稿抽关键词）
    _STOP = set("因故该工程原定我部经与及其有关进行需要根据实际情况现场核实商议后决定"
                "造成相关材料有所增减减少增加如下问题出现情况时间工作要求以上"
                "本次变更理由关于对于同时并且然后因此所以但是"
                "一二三四五六七八九十的了在是有和就不人都一个"
                "予以提出审批兹申请单工程量变更设计变更原因"
                "相关对应配套管材管件材料变动有所" )

    @classmethod
    def _blocks(cls, text):
        """切分出语义块（不做块内滑窗）——这些词适合展示。"""
        t = str(text or "")
        t = re.sub(r"[（(][^）)]*[）)]", " ", t)          # 去括号注释
        t = re.sub(r"\d+(?:\.\d+)?\s*(?:米|m|个|片|套|处|根|副|件|支)", " ", t)
        t = re.sub(r"[^一-鿿A-Za-z_]+", " ", t)
        parts = re.split(r"[，,。;；、\s]+|"
                         r"(?:因|故|经|的|及|并|后|时|而|就|也|还|又|再)", t)
        out = set()
        for seg in parts:
            seg = seg.strip()
            if len(seg) < 3 or any(ch.isdigit() for ch in seg):
                continue
            if all(ch in cls._STOP for ch in seg):
                continue
            out.add(seg)
        return out

    @classmethod
    def _keywords(cls, text):
        """匹配用关键词 = 语义块 + 块内子串。

        块内滑窗能产生「接入口」「灶具接入口」这类更短的匹配词
        （历史原因的切块上下文与草稿不同，长块常匹配不上），
        但滑窗碎片不适合展示，故展示时另用 _blocks。
        """
        out = set(cls._blocks(text))
        for seg in cls._blocks(text):
            for n in (5, 4, 3):
                for i in range(len(seg) - n + 1):
                    w = seg[i:i + n]
                    if any(ch.isdigit() for ch in w):
                        continue
                    if all(ch in cls._STOP for ch in w):
                        continue
                    out.add(w)
        return out

    @classmethod
    def _display_words(cls, text, add=None, red=None):
        """可展示的命中词。

        = 语义块 + 块内 4–6 字子串 + 材料基类名，
          全部剔除程序性套话（报告公司…现场核实…等）。
        块内子串保证「矮立管移位」这类具体表述能被历史原文匹配到。
        """
        out = set()
        for seg in cls._blocks(text):
            if len(seg) >= 4 and not cls._is_proc(seg):
                out.add(seg)
            for n in (6, 5, 4):
                for i in range(len(seg) - n + 1):
                    w = seg[i:i + n]
                    if any(ch.isdigit() for ch in w):
                        continue
                    if cls._is_proc(w):
                        continue
                    if all(ch in cls._STOP for ch in w):
                        continue
                    out.add(w)
        for w in cls._material_words(add) | cls._material_words(red):
            if len(w) >= 3 and not cls._is_proc(w):
                out.add(w)
        return out

    @classmethod
    def _material_words(cls, items):
        """从材料清单提取基类名（如 '镀锌弯头_DN40_90°_内丝' → '镀锌弯头'）。

        草稿通常会提到具体材料，用材料名做关键词比中文分词可靠得多。
        """
        out = set()
        for x in items or []:
            name = x["name"] if isinstance(x, dict) else str(x)
            base = re.split(r"[_（(\s]", name)[0]
            if len(base) >= 2 and not any(ch.isdigit() for ch in base):
                out.add(base)
        return out

    #: 程序性套话：可作为「文风一致」的弱信号，但不具区分度，展示时排除
    _PROC = ("报告公司", "相关代表", "人员代表", "现场核实", "商议后决定", "核实后决定",
             "兹提出", "申请单", "请予以审批", "工程量变更单", "变更申请单",
             "公司相关", "代表到", "关代表", "表到现", "到现场", "场核实")

    @classmethod
    def _is_proc(cls, word):
        return any(p in word for p in cls._PROC)

    def related(self, add_norm, red_norm, code="", draft="", top=8,
                add=None, red=None):
        """为「润色」场景检索相关历史案例。

        与 similar() 的区别：多叠加两个维度，
        避免材料差异大时相似度全为 0 而返回无关案例——
          ① 材料 Jaccard（权重 3）
          ② 同类别工程（细分类型码相同，权重 1）
          ③ 草稿关键词与历史原因的命中（权重 1.5）

        关键词权重按 IDF 折减：像「报告公司相关代表到现场核实」这种
        历史语料里人人都有���通用套话，DF 极高、区分度为零，应当降权；
        而「灶具接入口位置」这类具体表述 DF 低，权重高。

        关键词来源：草稿语义块 + 增/减材料基类名。
        返回 (综合分, 材料分, 命中理由, 记录) 列表。
        """
        kw = self._keywords(draft)
        kw |= self._material_words(add) | self._material_words(red)
        cat = (code or "")[4:5] if code and len(code) >= 5 else ""
        n_all = max(1, len([r for r in self.records if r["reason"]]))
        scored = []
        for r in self.records:
            if not r["reason"] or r["reason"] == PLACEHOLDER:
                continue
            mat = (0.5 * self.jaccard(add_norm, r["add_norm"])
                   + 0.5 * self.jaccard(red_norm, r["red_norm"]))
            same_cat = bool(cat) and (r["code"] or "")[4:5] == cat
            hits = []
            hit_idf = 0.0
            for k in kw:
                if k in r["reason"]:
                    hits.append(k)
                    hit_idf += self._idf(k, n_all)
            total_idf = sum(self._idf(k, n_all) for k in kw) or 1.0
            kw_ratio = hit_idf / total_idf
            score = mat * 3.0 + (1.0 if same_cat else 0.0) + kw_ratio * 3.0
            if score <= 0:
                continue
            why = []
            if mat > 0:
                why.append("材料相似 %.2f" % mat)
            if same_cat:
                why.append("同类工程")
            # 展示用词：语义块 + 材料基类名（不含滑窗碎片与程序性套话）
            disp = self._display_words(draft, add, red)
            good = [k for k in disp if k in r["reason"] and self._idf(k, n_all) >= 0.3]
            good.sort(key=lambda k: (-len(k), -self._idf(k, n_all)))
            picked = []
            for k in good:
                if not any(k in p for p in picked):
                    picked.append(k)
            if picked:
                why.append("关键词命中：" + "、".join(picked[:3]))
            elif hits:
                why.append("文风与同类工程一致")
            scored.append((score, mat, "；".join(why), r))
        scored.sort(key=lambda x: -x[0])
        return scored[:top]

    #: 关键词文档频率缓存
    _df_cache = None

    def _doc_freq(self, word):
        """关键词在历史原因原文中的文档频率（缓存）。

        直接对**原始原因文本**做子串判断，而不是对转换后的词集——
        否则「报告公司相关代表到现场核实」这类套话因切分方式不同
        而统计失真，无法被 IDF 识别为低区分度词。
        """
        if Corpus._df_cache is None:
            df = {}
            for r in self.records:
                if not r["reason"] or r["reason"] == PLACEHOLDER:
                    continue
                txt = r["reason"]
                for w in set(self._keywords(txt)) | self._material_words(r["add"]) \
                        | self._material_words(r["red"]):
                    df[w] = df.get(w, 0) + 1
            Corpus._df_cache = df
        return Corpus._df_cache.get(word, 0)

    def _raw_doc_freq(self, word):
        """在历史原因**原文**中直接统计包含该词的文档数。

        必须基于原文子串匹配，不能只查 _keywords 词集：
        不同文本的切块结果不同（"报告公司相关代表…" vs "报告公司相关人员代表…"），
        若按词集统计，套话会被误判为"罕见词"而获得最高权重。
        """
        cache = getattr(Corpus, "_raw_df_cache", None)
        if cache is None:
            cache = {}
            for r in self.records:
                if not r["reason"] or r["reason"] == PLACEHOLDER:
                    continue
                txt = r["reason"]
                for w in self._keywords(txt):
                    if w in txt:
                        cache[w] = cache.get(w, 0) + 1
            Corpus._raw_df_cache = cache
        if word in cache:
            return cache[word]
        # 缓存未覆盖（来自草稿的切块）→ 直接遍历原文统计
        n = 0
        for r in self.records:
            if r["reason"] and word in r["reason"]:
                n += 1
        cache[word] = n
        return n

    def _idf(self, word, n_all):
        """IDF：越少见的词权重越高。

        用原始文本的文档频率计算，能正确识别「人人都有」的通用套话。
        """
        df = self._raw_doc_freq(word) / float(n_all)
        return 1.0 - min(df, 0.95)

    def stats(self):
        filled = [r for r in self.records if r["reason"] and r["reason"] != PLACEHOLDER]
        return {
            "total": len(self.records),
            "filled": len(filled),
            "pending": len(self.records) - len(filled),
        }


# ---------------------------------------------------------------- 原因合成

#: 触发因子 -> (关键词, 原因模板)
#: 模板中 {n} 为数量占位，{q} 会被替换为实际数量
FACTORS = [
    {
        "key": "pe_to_steel",
        "name": "PE管改钢管（材质转换）",
        "material_kw": ["pe管", "pe警示带", "pe管_dn", "鞍型", "热熔弯头_dn", "电熔"],
        "steel_kw": ["钢无缝管", "钢制弯头", "钢制三通", "钢制大小头", "焊接法兰", "热镀锌钢管", "镀锌"],
        "template": ("因原设计{old}在该处无法满足敷设及安装要求，经报告公司相关代表到现场核实并商议后决定，"
                     "改为采用{new}敷设，配套管件及连接方式相应调整，故造成相关管材管件有所增减"),
    },
    {
        "key": "instrument",
        "name": "仪表/阀门连接件",
        "material_kw": ["单头牙", "压力表", "铜球阀", "四氟垫", "金属缠绕垫", "过滤网", "跨接线"],
        "template": "因{what}需要{conn}，故增加{item}",
    },
    {
        "key": "flange_blind",
        "name": "法兰盲板封堵预留",
        "material_kw": ["盲板", "焊接法兰", "丝扣法兰"],
        "template": "因{what}采用焊接法兰加盲板进行封堵预留，故增加{item}",
    },
    {
        "key": "underground",
        "name": "地下管线避让",
        "material_kw": ["弯头", "三通", "套筒", "变径"],
        "template": ("该工程位于{where}，地下管线及其他设施较多，需根据现场开挖实际情况调整管线敷设位置及长度，"
                     "并增加管件绕行避让障碍，故造成相关管材管件有所增减"),
    },
    {
        "key": "protect_groove",
        "name": "埋深不够加保护沟",
        "material_kw": ["保护沟", "钢板"],
        "template": ("因{what}段管道埋深不够，根据公司相关要求需敷设{how}保护沟，"
                     "故造成相关材料有所增减"),
    },
    {
        "key": "collision",
        "name": "碰接方式变更",
        "material_kw": ["鞍型电熔", "不停输", "停气碰接", "碰接"],
        "template": ("原定碰接方式因{why}无法实施，经报告公司相关代表到现场核实商议后决定，"
                     "改为{how}碰接，故造成相关材料有所增减"),
    },
    {
        "key": "reserved",
        "name": "增加预留",
        "material_kw": ["预留", "端帽", "球阀阀门井"],
        "template": "因{why}需增加预留，故增加{item}，造成相关材料有所增加",
    },
]

#: 结尾语
TAILS = {
    "app": TAIL_APP,
    "qty": TAIL_QTY,
}


class ReasonBuilder(object):
    """依据材料清单生成变更原因。

    主路径走「材料级推理」（material_reason 模块）：
    从增/减材料反推**工程动作**，再按历史文风组装。

    保留本类作为兼容入口；material_reason 不可用时退化为模板兜底。
    """

    def __init__(self, corpus):
        self.corpus = corpus

    # -- 特征识别 ------------------------------------------------------

    @staticmethod
    def detect(add, red):
        """识别触发因子，返回命中列表。"""
        try:
            from material_reason import derive_actions
            return sorted(set(a.kind for a in derive_actions(add, red, "", "")))
        except ImportError:
            pass
        add_n = [x["norm"] for x in add]
        red_n = [x["norm"] for x in red]
        all_n = add_n + red_n
        hits = []
        if (any(("pe管" in n or "pe警示" in n) for n in red_n)
                and any(("钢无缝管" in n or "钢制" in n or "镀锌" in n) for n in add_n)):
            hits.append("pe_to_steel")
        if any(k in n for n in all_n
               for k in ("单头牙", "压力表", "铜球阀", "四氟垫", "金属缠绕垫", "过滤网", "跨接线")):
            hits.append("instrument")
        if any("盲板" in n for n in add_n):
            hits.append("flange_blind")
        if any("保护沟" in n for n in all_n):
            hits.append("protect_groove")
        return hits

    # -- 正文合成 ------------------------------------------------------

    def build(self, add, red, name="", code="", location="", extra=None,
              reuse_project=True):
        """生成变更原因正文（不含结尾语）。

        返回 (正文, 命中因子列表, 参考案例 list, 同项目定稿 dict|None)

        同项目定稿（同工程编号/同工程名称的历史已填原因）优先复用：
        它承载了材料推不出的现场事实，比纯材料推理更贴近项目实际。
        """
        extra = (extra or "").strip()
        add_norm = sorted(set(x["norm"] for x in add))
        red_norm = sorted(set(x["norm"] for x in red))
        refs = self.corpus.similar(add_norm, red_norm, top=5)

        project = None
        if reuse_project:
            hit = self.corpus.project_match(code=code, name=name,
                                            add_norm=add_norm, red_norm=red_norm)
            if hit:
                project = {"score": hit[0], "record": hit[1]}

        try:
            from material_reason import compose, derive_actions, reuse_project_reason
        except ImportError:
            return (self._build_fallback(add, red, name, location, extra),
                    self.detect(add, red), refs, None)

        if project:
            body = reuse_project_reason(project["record"]["reason"], add, red,
                                        name=name, code=code, extra=extra)
            hits = sorted(set(a.kind for a in derive_actions(add, red, name, code)))
            return body, hits, refs, project

        body, acts = compose(add, red, name=name, code=code,
                             location=location, extra=extra)
        return body, sorted(set(a.kind for a in acts)), refs, None

    # -- 兜底（material_reason 不可用时）-------------------------------

    def _build_fallback(self, add, red, name, location, extra):
        segs = []
        if extra:
            segs.append(extra.rstrip("。；; ") + "，")
        loc = location or self.infer_location(name)
        segs.append(self._site_clause(loc))
        body = "".join(segs)
        return body if body.endswith("；") else body + "；"

    @staticmethod
    def infer_location(name):
        """从工程名称推断现场场景。"""
        try:
            from material_reason import infer_scene
            s = infer_scene(name)
            if s:
                return s + "内"
        except ImportError:
            pass
        return ""

    @staticmethod
    def _site_clause(loc):
        """室内 / 室外兜底句（与 material_reason.fallback_clause 共用实现）。"""
        try:
            from material_reason import fallback_clause
            return fallback_clause(location=loc)
        except ImportError:
            pass
        return ("因该工程地下管线及其他设施较多，需根据现场开挖实际情况调整管线敷设位置"
                "及长度，并增加管件绕行避让障碍，故造成相关管材管件有所增减；")


# ---------------------------------------------------------------- 写入 (.xls)

class ExcelWriter(object):
    """通过 Excel COM 写入 .xls（xlrd 只能读）。

    .xls 老格式必须走 Excel COM；写入采用「先复制到临时文件 → 改 → 覆盖原文件」
    的方式，避免 Excel 独占锁导致原文件损坏。
    """

    def __init__(self):
        self._app = None

    def _excel(self):
        if self._app is None:
            import win32com.client
            self._app = win32com.client.DispatchEx("Excel.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = False
        return self._app

    def close(self):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def write(self, path, sheet_name, cells):
        """写入单个工作表（兼容旧接口）。"""
        return self.write_many(path, {sheet_name: cells})

    def write_many(self, path, sheet_cells):
        """一次 COM 会话向多个工作表批量写入。

        sheet_cells: {sheet_name: {"B5": "文本", "A49": "文本"}}
        返回 True 表示成功。
        """
        src = os.path.abspath(path)
        tmp_dir = tempfile.mkdtemp(prefix="wbx_")
        tmp = os.path.join(tmp_dir, os.path.basename(src))
        shutil.copy2(src, tmp)
        ok = False
        wb = None
        try:
            app = self._excel()
            wb = app.Workbooks.Open(tmp)
            for sheet_name, cells in sheet_cells.items():
                ws = wb.Worksheets(sheet_name)
                for addr, val in cells.items():
                    ws.Range(addr).Value = val
            wb.Save()
            wb.Close()
            wb = None
            shutil.copy2(tmp, src)
            ok = True
        finally:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return ok


def backup_file(path, backup_dir=None):
    """写入前把原件复制到 `_变更单备份/`，返回备份路径；失败返回空串。

    默认放在与原件同级的 `_变更单备份` 子目录，文件名带时间戳，
    避免同名覆盖；该目录已被语料扫描排除。
    """
    src = os.path.abspath(path)
    d = backup_dir or os.path.join(os.path.dirname(src), "_变更单备份")
    try:
        if not os.path.isdir(d):
            os.makedirs(d)
        ts = time.strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(os.path.basename(src))
        dst = os.path.join(d, "%s.%s.bak%s" % (base, ts, ext))
        shutil.copy2(src, dst)
        return dst
    except Exception:
        return ""


def compose_texts(app_text, qty_text=None, existing_app=""):
    """正文 → (申请单全文, 工程量变更单全文)，结尾语规范化。

    结尾语沿用该文件已有写法（同一单位/同批单据可能用不同变体）。
    """
    app_text = (app_text or "").strip()
    tail_app = detect_tail(existing_app) or TAIL_APP
    if app_text.endswith(tail_app):
        app_full = app_text
    else:
        core = app_text.rstrip("；;。 ")
        app_full = (core + "；" + tail_app) if core else tail_app

    if qty_text:
        q = qty_text.strip()
        if not q.endswith(TAIL_QTY):
            qcore = q.rstrip("；;。 ")
            q = (qcore + "；" + TAIL_QTY) if qcore else TAIL_QTY
        qty_full = q
    else:
        body = app_full[: -len(tail_app)].rstrip("；; ") if app_full.endswith(tail_app) else app_full
        qty_full = body + "；" + TAIL_QTY
    return app_full, qty_full


def fill_orders(path, items, overwrite=False, backup=True, backup_dir=None):
    """按单据写入变更原因（一次 Excel 会话写完全部 sheet）。

    items: [{
        "label":     日志标识，如「767-变更单打印（地下）#1」
        "app_slot":  (sheet, row, col) 申请单原因格；可为 None
        "qty_slot":  (sheet, row, col) 工程量单原因格；可为 None
        "app_text":  申请单正文（不含结尾语亦可）
        "qty_text":  工程量单正文；None 时由申请单正文派生
        "existing":  该单据已有内容（安全校验用，非空且未授权则跳过）
    }]

    返回 dict：ok / written / skipped / backup
    """
    written, skipped = [], []
    cells_by_sheet = {}

    for it in items:
        if it.get("existing") and not overwrite:
            skipped.append({"label": it.get("label", ""),
                            "reason": "变更原因已有内容，未授权覆盖"})
            continue
        app_slot, qty_slot = it.get("app_slot"), it.get("qty_slot")
        if not app_slot and not qty_slot:
            skipped.append({"label": it.get("label", ""), "reason": "未定位到原因格"})
            continue
        app_full, qty_full = compose_texts(it.get("app_text") or "",
                                          it.get("qty_text"),
                                          it.get("existing_app") or "")
        cells = {}
        if app_slot:
            cells[_addr(app_slot[1], app_slot[2])] = app_full
        if qty_slot:
            cells[_addr(qty_slot[1], qty_slot[2])] = qty_full
        sheet = (app_slot or qty_slot)[0]
        cells_by_sheet.setdefault(sheet, {}).update(cells)
        written.append({"label": it.get("label", ""), "sheet": sheet, "cells": cells})

    if not written:
        return {"ok": False, "written": [], "skipped": skipped, "backup": ""}

    bak = backup_file(path, backup_dir) if backup else ""
    w = ExcelWriter()
    try:
        ok = w.write_many(path, cells_by_sheet)
    finally:
        w.close()
    return {"ok": ok, "written": written, "skipped": skipped, "backup": bak}


def fill_reason(path, app_text, qty_text=None, overwrite=False,
                order_index=0, backup=True, backup_dir=None):
    """把变更原因写入一张变更单的申请单与工程量变更单两个位置。

    app_text   : 申请单「设计变更原因」正文（不含结尾语亦可，脚本会补全）
    qty_text   : 工程量变更单正文；None 时自动由 app_text 派生
    overwrite  : 是否授权覆盖目标格已有内容。
                 False（默认）——目标格已有内容时**中止**，避免用自动生成的
                 通用文案覆盖掉人工写好的真实原因。
                 True——允许覆盖（CLI 的 --body 参数即视为授权）。
    order_index: 该工作簿内的第几张变更单（0 起）。多张变更单时必须逐张指定。
    backup     : 写入前是否自动备份原件（默认是）。

    返回 (ok, cells, slots, backup_path)
    """
    co = ChangeOrder(path)
    orders = co.orders()
    if not orders:
        raise RuntimeError("未找到任何变更单（既无『设计变更原因』也无『工程量变更单』）：%s" % path)
    if order_index < 0 or order_index >= len(orders):
        raise RuntimeError("变更单序号超出范围：该文件共 %d 张变更单，请求第 %d 张"
                           % (len(orders), order_index + 1))
    o = orders[order_index]
    if not o["app_slot"]:
        raise RuntimeError("未找到申请单变更原因单元格位置：%s（%s）" % (path, o["sheet"]))

    existing = o["draft_app"] if o["draft_app"] and o["draft_app"] != PLACEHOLDER else ""
    if existing and not overwrite:
        raise RuntimeError(
            "该变更单的变更原因已填写，为避免覆盖已有内容已中止。\n"
            "  位置：%s!%s\n  已有内容：%s\n"
            "  若是用户草稿需润色：请用 polish 命令取素材，再以 --body 写入润色结果；\n"
            "  若确需用新正文重写：请以 --body 指定新正文（该参数授权覆盖）。"
            % (o["sheet"], o["app_addr"], existing[:40])
        )

    if existing:
        sys.stderr.write("[警告] %s!%s 原有内容将被覆盖：%s\n"
                         % (o["sheet"], o["app_addr"], existing[:50]))

    res = fill_orders(path, [{
        "label": "%s#%d" % (o["sheet"], o["index"]),
        "app_slot": o["app_slot"],
        "qty_slot": o["qty_slot"],
        "app_text": app_text,
        "qty_text": qty_text,
        "existing": existing,
        "existing_app": existing,
    }], overwrite=True, backup=backup, backup_dir=backup_dir)

    cells = res["written"][0]["cells"] if res["written"] else {}
    slots = {"app": o["app_slot"]}
    if o["qty_slot"]:
        slots["qty"] = o["qty_slot"]
    return res["ok"], cells, slots, res["backup"]


def _addr(r, c):
    """0-based (row, col) -> A1 地址。"""
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return "%s%d" % (s, r + 1)
