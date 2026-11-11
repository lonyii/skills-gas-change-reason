# -*- coding: utf-8 -*-
"""XchangeSheet.xls 的 JSON 缓存层（替换表 + 套价数据库）。

设计
----
* ``xchange_cache.json`` 由本模块 ``build`` 生成，随技能一起分发；
  运行时 ``load()`` **只读 JSON**，不依赖 xlrd、不开 Excel，速度快且零依赖。
* 源文件（XchangeSheet.xls）比缓存新（大小或修改时间变了）时，``load()`` 自动重建缓存；
  源文件不在本机时继续用缓存，不报错。
* 缓存写入是原子的（先写 .tmp 再 os.replace），不会留下半个文件。

语义对齐原 VBA（``图纸和变更量自动化处理.bas``）
--------------------------------------------
* ``simplify(text)``      —— XChange：**先删全部空格**，再按「替换表」顺序逐条替换（忽略大小写）
* ``simplify_cell(text)`` —— XChangeAll：**不删空格**，按「替换表」顺序逐条替换（忽略大小写），
                             再去掉**一个**结尾下划线
* ``normalize(text)``     —— ReplaceX：转小写后按「物料描述分隔符 / 替换为」逐条替换
* ``quota_row(key)``      —— 规范化描述 → 首次出现的行号（1 基，0 表示未匹配）

用法
----
    python xchange_db.py build [--source <XchangeSheet.xls>] [--force]
    python xchange_db.py status
    python xchange_db.py simplify "PE管_PE100_SDR11_dn63_黑色"
"""
from __future__ import print_function

import json
import os
import sys
import time

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(SKILL_DIR, "scripts")
CONFIG_FILE = os.path.join(SCRIPTS_DIR, "config.json")
CACHE_FILE = os.path.join(SCRIPTS_DIR, "xchange_cache.json")
CACHE_VERSION = 2

#: 源文件候选路径（按顺序尝试，先命中者为准）
DEFAULT_SOURCES = [
    r"F:\Scripts\Xchange\XchangeSheet.xls",
    "XchangeSheet.xls",
]

REPLACE_SHEET = "替换表"
DB_SHEET = "数据库"
SEP_HEADER = "物料描述分隔符"


# --------------------------------------------------------------- 配置 / 路径

def _read_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                return json.load(fh) or {}
        except Exception:
            return {}
    return {}


def source_path(explicit=None):
    """解析源 xls 路径：参数 > 配置 > 环境变量 > 默认候选。找不到返回 None。"""
    if explicit:
        return explicit if os.path.exists(explicit) else explicit
    cfg = _read_config()
    p = (cfg.get("xchange_file") or "").strip()
    if p and os.path.exists(p):
        return p
    p = (os.environ.get("GAS_XCHANGE_FILE") or "").strip()
    if p and os.path.exists(p):
        return p
    for cand in DEFAULT_SOURCES:
        if os.path.isabs(cand):
            if os.path.exists(cand):
                return cand
        else:
            for base in (os.getcwd(), SCRIPTS_DIR, SKILL_DIR):
                full = os.path.join(base, cand)
                if os.path.exists(full):
                    return full
    return None


def _stat(path):
    st = os.stat(path)
    return {"path": os.path.abspath(path), "size": st.st_size, "mtime": round(st.st_mtime, 3)}


# --------------------------------------------------------------- 抽取

def _cell_text(v):
    """把单元格值转成 VBA UsedRange.Value 里那种「文本或数字」。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return int(v) if v.is_integer() else v
    return v


def extract(src):
    """从 xls 抽出替换表 + 数据库（含分隔符规则）。返回缓存 dict。"""
    try:
        import xlrd
    except ImportError:
        raise RuntimeError("重建缓存需要 xlrd：pip install xlrd")
    bk = xlrd.open_workbook(src, on_demand=True)
    try:
        # ---- 替换表
        sh = bk.sheet_by_name(REPLACE_SHEET)
        rules = []
        for r in range(1, sh.nrows):          # 跳过表头
            what = _cell_text(sh.cell_value(r, 0))
            repl = _cell_text(sh.cell_value(r, 1)) if sh.ncols > 1 else ""
            if what == "" or what is None:
                continue
            rules.append([what, "" if repl is None else repl])

        # ---- 数据库
        sh = bk.sheet_by_name(DB_SHEET)
        headers = [_cell_text(sh.cell_value(0, c)) for c in range(sh.ncols)]
        rows = []
        for r in range(1, sh.nrows):
            row = [_cell_text(sh.cell_value(r, c)) for c in range(sh.ncols)]
            if all((x == "" or x is None) for x in row):
                continue                      # 整行空白不入缓存
            rows.append(row)

        # ---- 分隔符规则（表内「物料描述分隔符 / 替换为」两列，遇到空行即止）
        sep_col = -1
        for i, h in enumerate(headers):
            if str(h).strip() == SEP_HEADER:
                sep_col = i
                break
        separators = []
        if sep_col >= 0:
            for row in rows:
                sep = row[sep_col] if sep_col < len(row) else ""
                if sep == "" or sep is None:
                    break
                repl = row[sep_col + 1] if sep_col + 1 < len(row) else ""
                separators.append([sep, "" if repl is None else repl])

        return {
            "cache_version": CACHE_VERSION,
            "built_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "source": _stat(src),
            "replace_rules": rules,
            "separators": separators,
            "db": {"headers": headers, "rows": rows},
        }
    finally:
        try:
            bk.release_resources()
        except Exception:
            pass


# --------------------------------------------------------------- 建缓存

def build(src=None, cache_file=CACHE_FILE, verbose=True):
    src = source_path(src)
    if not src or not os.path.exists(src):
        raise RuntimeError("找不到 XchangeSheet.xls：%s" % (src or "（未配置）"))
    data = extract(src)
    tmp = cache_file + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, cache_file)
    if verbose:
        size_mb = os.path.getsize(cache_file) / 1048576.0
        print("[OK] 缓存已重建：%s（替换规则 %d 条 / 数据库 %d 行 / 分隔符 %d 条，%.2f MB）" %
              (cache_file, len(data["replace_rules"]), len(data["db"]["rows"]),
               len(data["separators"]), size_mb))
    return data


def is_stale(cache, src=None):
    """缓存是否比源文件旧 / 不匹配。源文件不存在时视为「不旧」。"""
    src = source_path(src)
    if not src or not os.path.exists(src):
        return False
    meta = (cache or {}).get("source") or {}
    if not meta.get("path") or os.path.abspath(meta["path"]) != os.path.abspath(src):
        return True
    try:
        st = os.stat(src)
    except OSError:
        return False
    if int(meta.get("size") or -1) != st.st_size:
        return True
    return abs(float(meta.get("mtime") or 0) - round(st.st_mtime, 3)) > 0.001


def load(force_rebuild=False, cache_file=CACHE_FILE, src=None, quiet=False):
    """返回缓存 dict：优先读 JSON；源文件更新过则先重建；无缓存则重建。"""
    cache = None
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as fh:
                cache = json.load(fh)
            if cache.get("cache_version") != CACHE_VERSION:
                cache = None
        except Exception:
            cache = None

    if cache is not None and not force_rebuild and not is_stale(cache, src):
        return cache

    if cache is not None and not quiet:
        print("[INFO] 检测到 XchangeSheet.xls 有更新，正在重建缓存 …")
    try:
        return build(src=src, cache_file=cache_file, verbose=not quiet)
    except Exception as e:
        if cache is not None:
            if not quiet:
                print("[WARN] 缓存重建失败（沿用旧缓存）：%s" % e)
            return cache
        raise


# --------------------------------------------------------------- 查询对象

class XchangeDB(object):
    """缓存的内存视图：规则、分隔符、数据库与匹配索引。"""

    def __init__(self, data):
        self.data = data
        self.replace_rules = [(r[0], r[1]) for r in data.get("replace_rules", [])]
        self.separators = [(s[0], s[1]) for s in data.get("separators", [])]
        headers = data.get("db", {}).get("headers", [])
        self.headers = headers
        self.rows = data.get("db", {}).get("rows", [])
        self._col = {str(h).strip(): i for i, h in enumerate(headers) if str(h).strip()}
        self._index = None
        self._simplify_cache = {}

    # ---- 基础查询
    def col(self, name):
        return self._col.get(name, -1)

    def code_col(self, ver):
        return self.col("编码%s" % ver)

    def value(self, row_no, col_name):
        i = self.col(col_name)
        if i < 0 or row_no < 1 or row_no > len(self.rows):
            return ""
        row = self.rows[row_no - 1]
        return row[i] if i < len(row) else ""

    # ---- 文本处理（对齐 VBA）
    def simplify(self, text):
        """XChange：删全部空格 → 按替换表逐条替换（忽略大小写），带结果缓存。"""
        s = "" if text is None else str(text)
        if s in self._simplify_cache:
            return self._simplify_cache[s]
        out = s.replace(" ", "")
        for what, repl in self.replace_rules:
            if not what:
                continue
            out = _ireplace(out, what, repl)
        self._simplify_cache[s] = out
        return out

    def simplify_cell(self, text):
        """XChangeAll 的单格语义：不删空格，逐条替换后去掉一个结尾下划线。"""
        s = "" if text is None else str(text)
        for what, repl in self.replace_rules:
            if not what:
                continue
            s = _ireplace(s, what, repl)
        if s.endswith("_"):
            s = s[:-1]
        return s

    def normalize(self, text):
        """ReplaceX：转小写后按分隔符规则逐条替换。"""
        s = "" if text is None else str(text).lower()
        for sep, repl in self.separators:
            if not sep:
                continue
            s = _ireplace(s, sep, repl)
        return s

    def quota_row(self, key):
        """规范化描述 → 首次出现的行号（1 基）；0 表示未匹配。"""
        if self._index is None:
            idx = {}
            mcol = self.col("物料描述")
            if mcol >= 0:
                for i, row in enumerate(self.rows, 1):
                    k = self.normalize(row[mcol] if mcol < len(row) else "")
                    if k not in idx:
                        idx[k] = i
            self._index = idx
        return self._index.get(key, 0)

    def stats(self):
        return {
            "rules": len(self.replace_rules),
            "separators": len(self.separators),
            "db_rows": len(self.rows),
            "headers": self.headers,
            "built_at": self.data.get("built_at"),
        }


def _ireplace(text, old, new):
    """忽略大小写的全量替换（等价 VBA Replace(..., vbTextCompare)）。

    先做一次「命中判断」直接返回——替换表有 200 多条规则，
    逐条跑 str.find 在整张表上会明显拖慢。
    """
    if not old:
        return text
    o = old.lower()
    low = text.lower()
    if o not in low:
        return text
    out = []
    i = 0
    n = len(o)
    while True:
        j = low.find(o, i)
        if j < 0:
            out.append(text[i:])
            break
        out.append(text[i:j])
        out.append(new)
        i = j + n
    return "".join(out)


def get_db(force_rebuild=False, quiet=True, path=None):
    """便捷入口：拿到可直接查询的 XchangeDB。"""
    return XchangeDB(load(force_rebuild=force_rebuild, quiet=quiet, src=path))


# --------------------------------------------------------------- CLI

def main(argv):
    cmd = argv[0] if argv else "status"
    args = argv[1:]
    if cmd == "build":
        src = None
        forced = "--force" in args
        if "--source" in args:
            src = args[args.index("--source") + 1]
        build(src=src)
        return 0
    if cmd == "status":
        src = source_path()
        print("源文件：%s" % (src or "（未找到，将只用缓存）"))
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            print("缓存：%s（%.2f MB，生成于 %s）" %
                  (CACHE_FILE, os.path.getsize(CACHE_FILE) / 1048576.0, data.get("built_at")))
            print("内容：替换规则 %d 条 / 数据库 %d 行 / 分隔符 %d 条" %
                  (len(data.get("replace_rules", [])), len(data.get("db", {}).get("rows", [])),
                   len(data.get("separators", []))))
            print("是否过期：%s" % ("是（下次运行会自动重建）" if is_stale(data, src) else "否"))
        else:
            print("缓存：尚未生成")
        return 0
    if cmd in ("simplify", "simplify_cell", "normalize"):
        text = args[0] if args else ""
        db = get_db(quiet=False)
        print(getattr(db, cmd)(text))
        return 0
    print("用法: xchange_db.py [build|status|simplify|simplify_cell|normalize] ...")
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
