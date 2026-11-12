# -*- coding: utf-8 -*-
"""
燃气工程变更单 —— 变更原因生成 CLI

用法:
  # 1) 只分析，不写文件（推荐先看结果）
  python fill_change_reason.py analyze "<变更单.xls>"

  # 2) 分析并写入
  python fill_change_reason.py fill "<变更单.xls>" [--location "南昌县老城区"] [--extra "现场补充事实"]

  # 3) 单元格已有用户草稿 → 输出润色素材包，AI 润色后回写
  python fill_change_reason.py polish "<变更单.xls>"
  python fill_change_reason.py fill "<变更单.xls>" --body "<AI润色后的正文>"

  # 4) 重建历史语料库缓存
  python fill_change_reason.py reindex

  # 5) 检索历史相似案例
  python fill_change_reason.py search "<变更单.xls>" [--top 8]

  # 6) 检查 / 安装技能更新（从 GitHub 仓库）
  python fill_change_reason.py update [--check] [--rollback]
"""

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import os
import re
import sys
import json
import glob
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gas_change_core import (  # noqa: E402
    ChangeOrder, Corpus, ReasonBuilder, fill_reason, fill_orders,
    ARCHIVE_ROOT, ARCHIVE_ROOTS, PLACEHOLDER, TAIL_APP, TAIL_QTY,
)

CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "corpus_cache.json")


def cmd_polish(args):
    """输出「润色素材包」：用户草稿 + 材料推导 + 历史案例 + 文风规范。

    供 AI 参照历史语料把草稿润色成规范文案，再以 fill --body 写回。
    """
    co = ChangeOrder(args.file)
    orders = co.orders()
    if not orders:
        print("[错误] 未找到 '*变更单打印' 工作表或变更单。")
        return 2
    idx = max(1, args.order or 1)
    if idx > len(orders):
        print("[错误] 该文件共 %d 张变更单，--order=%d 超出范围" % (len(orders), idx))
        return 2
    o = orders[idx - 1]
    code, name = o["code"], o["name"]
    add, red = o["add"], o["red"]
    draft = co.read_draft(o)

    print("=" * 70)
    print("润色素材包 —— %s  %s" % (code, name))
    print("工作表 %s（该文件共 %d 张变更单，本张为第 %d 张）"
          % (o["sheet"], len(orders), idx))
    print("=" * 70)

    # 1) 用户草稿（提示词）
    print("\n【1】用户草稿（作为提示词，润色时须保留其中的现场事实）")
    if draft["has_draft"]:
        if draft["app"]:
            print("  申请单 %s：" % draft["app_addr"])
            print("    %s" % draft["app"])
        if draft["qty"] and draft["qty"] != draft["app"]:
            print("  工程量单 %s：" % draft["qty_addr"])
            print("    %s" % draft["qty"])
    else:
        print("  （无草稿 —— 单元格为空或为占位符，可直接用 fill 自动生成）")

    # 2) 工程编号分类
    print("\n【2】工程分类")
    try:
        from project_code import describe, parse_code
        p = parse_code(code)
        if p:
            print("  %s（%s） %s" % (p["cat_name"], p["group"], describe(code)))
    except ImportError:
        pass

    # 3) 材料增删
    print("\n【3】材料变动（润色时须与之一致）")
    print("  增加 %d 项：" % len(add))
    for x in add:
        print("    + %-42s %s" % (x["name"], x["qty"]))
    print("  减少 %d 项：" % len(red))
    for x in red:
        print("    - %-42s %s" % (x["name"], x["qty"]))

    # 4) 推导出的工程动作
    print("\n【4】材料 → 工程动作（自动推理，供交叉验证草稿）")
    try:
        from material_reason import derive_actions
        acts = derive_actions(add, red, name, code)
        if acts:
            for a in acts:
                print("    [%s] %s" % (a.kind, a.text))
        else:
            print("    （未命中既有动作类型 —— 草稿信息为主要依据）")
    except ImportError:
        pass

    # 5) 相似历史案例（润色的措辞参照）
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    add_n = sorted(set(x["norm"] for x in add))
    red_n = sorted(set(x["norm"] for x in red))
    draft_text = draft["app"] or draft["qty"] or ""
    print("\n【5】相关历史案例（润色的措辞参照）")
    hits = corpus.related(add_n, red_n, code=code, draft=draft_text,
                          top=args.top, add=add, red=red)
    if not hits:
        print("  （未找到相关案例，请以草稿事实 + 文风规范为准）")
    for score, mat, why, r in hits:
        print("  [%.2f] %s" % (score, r["file"]))
        print("      关联：%s" % (why or "同类"))
        print("      原文：%s" % r["reason"])

    # 6) 文风规范
    print("\n【6】文风规范（历史 132 份语料归纳）")
    for line in _STYLE_RULES:
        print("  " + line)

    print("\n" + "=" * 70)
    print("下一步：AI 参照以上素材润色出正文，然后写回——")
    print('  python fill_change_reason.py fill "%s" --order %d --body "<润色后的正文>"'
          % (args.file, idx))
    print("=" * 70)
    return 0


_STYLE_RULES = [
    "① 结构：分句以「；」分隔，末句以「故造成…有所增减/减少/增加」收尾，结尾语前保留「；」",
    "② 结尾语（314 份语料）：申请单 69% 用「兹提出设计变更申请单，请予以审批。」；",
    "   23% 用「兹提出变更申请单，请予以审批。」——若同一单位/同批单据用的是后者，就沿用后者；",
    "   工程量变更单用「兹提出工程量变更单。」",
    "③ 每句以「因 / 该工程 / 原定 / 我部」开头，陈述事实后再给结论",
    "④ 少量点状专项材料要点名并带数量（如「故增加钢制单头牙_DN15-4个」）；",
    "   管长类总量变化用概括表述（「故造成相关管材管件有所增减」），不要罗列全部材料",
    "⑤ 决策过程写「经报告公司相关代表到现场核实并商议后决定」",
    "⑥ 长度 38–234 字（历史区间），多数 70–150 字",
    "⑦ 保留草稿中的现场事实（甲方要求 / 阻工 / 库房缺货 / 设备变动等），只改措辞不改事实",
    "",
    "新语料补充的常见句式：",
    "  · 因图纸/设计与现场不符：「因设计图碰接点位置与现场不符…」",
    "  · 物资供应问题：「由于物资供应部不再提供支架，需施工单位自行购买…」",
    "  · 配合施工方要求：「该工程为道路配合工程，我部需配合…」",
    "  · 结论收尾变体：「综上所述，兹提出变更申请单，请予以审批。」",
    "  · 变更清单式：「…造成相关材料增减如下；拟调整如下材料，兹提出变更申请单。」",
]


PLACEHOLDER_TXT = PLACEHOLDER


def _is_backup_path(path):
    """自动备份 / 用户自建备份目录里的副本不参与批量处理。"""
    parts = os.path.normpath(path).split(os.sep)
    if any("备份" in p for p in parts[:-1]):
        return True
    return ".bak." in os.path.basename(path)


def _collect_files(paths):
    """把参数（文件 / 目录）展开成 .xls 列表；自动跳过备份副本。"""
    out, seen = [], set()
    for p in paths:
        ap = os.path.abspath(p)
        if os.path.isdir(ap):
            for f in sorted(glob.glob(os.path.join(ap, "**", "*.xls"), recursive=True)):
                if _is_backup_path(f):
                    continue
                key = os.path.normcase(os.path.abspath(f))
                if key not in seen:
                    seen.add(key)
                    out.append(f)
        elif os.path.isfile(ap):
            key = os.path.normcase(ap)
            if key not in seen:
                seen.add(key)
                out.append(ap)
        else:
            print("[忽略] 路径不存在：%s" % p)
    return out


def _status_text(o):
    """原因格现状：空缺 / 占位符 / 已填（草稿）。"""
    if not o["has_draft"]:
        return "空缺"
    if (o["draft_app"] or "").strip() == PLACEHOLDER_TXT:
        return "占位符"
    return "已填"


def _order_title(o):
    return "第%d张[%s]" % (o["seq"], o["sheet"])


def _category_text(code):
    try:
        from project_code import describe, parse_code
        p = parse_code(code)
        if p:
            label = p["cat_name"] if p["group"] == p["cat_name"] \
                else "%s（%s）" % (p["cat_name"], p["group"])
            return "%s %s" % (label, describe(code))
    except ImportError:
        pass
    return ""


def cmd_analyze(args):
    """分析一个或多个文件 / 目录，逐张变更单预览原因与风险提示。"""
    files = _collect_files(args.files)
    if not files:
        print("[错误] 没有可分析的 .xls 文件")
        return 2
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    rb = ReasonBuilder(corpus)
    total_orders = 0
    fallback_cnt = 0

    for path in files:
        co = ChangeOrder(path)
        orders = co.orders()
        print("=" * 78)
        print("文件: %s" % os.path.basename(path))
        if not orders:
            print("  [错误] 未找到变更单（无「设计变更原因」也无「工程量变更单」）")
            print("         现有工作表：%s" % ", ".join(co.sheet_names))
            continue
        print("  变更单 %d 张" % len(orders))
        built = []
        for o in orders:
            total_orders += 1
            body, hits, refs, proj = rb.build(o["add"], o["red"], name=o["name"],
                                              code=o["code"],
                                              location=args.location,
                                              extra=args.extra)
            # 沿用同项目定稿时即便没有材料动作，也不属于「兜底句」
            fallback = (not hits) and not proj
            if fallback:
                fallback_cnt += 1
            built.append((o, body, hits, refs, proj, fallback))
            proj_txt = "无"
            if proj:
                proj_txt = ("同项目(%.2f)" % proj["score"]) if proj["score"] > 0 else "同项目(定稿)"
            print("  第%d张［%s］%s %s  增%d/减%d  现状:%s %s  定稿:%s  因子:%s%s" % (
                o["seq"], o["sheet"], o["code"] or "-", o["name"] or "-",
                len(o["add"]), len(o["red"]), _status_text(o),
                o["app_addr"] if o["has_draft"] else "",
                proj_txt, ",".join(hits) if hits else "无",
                "  ⚠兜底" if fallback else ""))

        for (o, body, hits, refs, proj, fallback) in built:
            if args.summary:
                continue
            print("\n" + "-" * 78)
            print("▍第%d张  %s!" % (o["seq"], o["sheet"]), end="")
            print("%s  %s" % (o["app_addr"], o["qty_addr"]))
            print("工程编号: %s    工程名称: %s" % (o["code"] or "-", o["name"] or "-"))
            cat = _category_text(o["code"])
            if cat:
                print("项目分类: %s" % cat)
            if args.detail:
                print("\n增加材料 (%d):" % len(o["add"]))
                for x in o["add"]:
                    print("   + %-45s %s" % (x["name"], x["qty"]))
                print("减少材料 (%d):" % len(o["red"]))
                for x in o["red"]:
                    print("   - %-45s %s" % (x["name"], x["qty"]))
                try:
                    from material_reason import derive_actions
                    acts = derive_actions(o["add"], o["red"], o["name"], o["code"])
                    if acts:
                        print("\n材料 → 工程动作 推理:")
                        for a in acts:
                            print("   [%s] %s" % (a.kind, a.text))
                except ImportError:
                    pass

            if proj:
                print("\n【同项目历史定稿】%s" % proj["record"]["file"])
                print("  %s" % proj["record"]["reason"])
                print("  → 本次正文沿用该口径，仅补入定稿未提及的本次材料动作")
                if not hits:
                    print("    （本次增减材料未推出新的工程动作，正文即为定稿口径 + 结论句）")
            if fallback:
                print("\n⚠ 未命中任何工程动作，正文回落到兜底句 —— 正式单据前请用")
                print("   --location / --extra 补充现场事实，或用 --body 指定正文。")

            print("\n【申请单正文】")
            print(body.rstrip("；") + "；" + TAIL_APP)
            print("\n【工程量变更单正文】")
            print(body.rstrip("；") + "；" + TAIL_QTY)

            if args.detail or args.summary:
                continue
            print("\n参考历史案例:")
            for s, r in refs:
                print("  [%.3f] %s" % (s, r["file"]))
                print("        %s" % (r["reason"][:100] + ("..." if len(r["reason"]) > 100 else "")))

    print("\n" + "=" * 78)
    print("合计 %d 个文件 / %d 张变更单 | 兜底句 %d 张" % (len(files), total_orders, fallback_cnt))
    if fallback_cnt:
        print("提示：标注 ⚠ 的变更单材料推不出工程动作，自动正文只是兜底，建议人工核校。")
    return 0


def cmd_fill(args):
    """批量写入：一个文件内的所有变更单各按自己的材料生成并写回。"""
    files = _collect_files(args.files)
    if not files:
        print("[错误] 没有可写入的 .xls 文件")
        return 2
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    rb = ReasonBuilder(corpus)
    rc = 0

    for path in files:
        co = ChangeOrder(path)
        orders = co.orders()
        print("=" * 78)
        print("文件: %s" % os.path.basename(path))
        if not orders:
            print("  [跳过] 未找到变更单")
            rc = 1
            continue

        if args.order is not None:
            if not (1 <= args.order <= len(orders)):
                print("  [跳过] --order=%d 超出范围（共 %d 张）" % (args.order, len(orders)))
                rc = 1
                continue
            targets = [orders[args.order - 1]]
        else:
            targets = orders

        if args.body and len(orders) > 1 and args.order is None:
            print("  [跳过] 该文件有 %d 张变更单，--body 必须配合 --order 指定写哪一张"
                  % len(orders))
            rc = 1
            continue

        items = []
        for o in targets:
            if args.body:
                body = args.body
            else:
                body, hits, refs, proj = rb.build(o["add"], o["red"], name=o["name"],
                                                  code=o["code"],
                                                  location=args.location,
                                                  extra=args.extra)
                if proj:
                    print("  · %s 沿用同项目定稿口径（相似度 %.2f）"
                          % (_order_title(o), proj["score"]))
                elif not hits:
                    print("  ⚠ %s 未命中工程动作，正文为兜底句，建议先 analyze 人工核校"
                          % _order_title(o))
            existing = o["draft_app"] if o["draft_app"] and o["draft_app"] != PLACEHOLDER_TXT else ""
            items.append({
                "label": _order_title(o),
                "app_slot": o["app_slot"],
                "qty_slot": o["qty_slot"],
                "app_text": body,
                "qty_text": None,
                "existing": existing,
                "existing_app": existing,
            })

        res = fill_orders(path, items,
                          overwrite=bool(args.body) or args.overwrite,
                          backup=bool(args.backup_dir) and not args.no_backup,
                          backup_dir=args.backup_dir or None)

        if res["skipped"]:
            for s in res["skipped"]:
                print("  · %s 跳过：%s" % (s["label"], s["reason"]))
        for w in res["written"]:
            print("  → %s 写入 %s" % (w["sheet"],
                                     ", ".join(w["cells"].keys())))
            for addr, val in w["cells"].items():
                print("     %-5s ← %s" % (addr, val[:70] + ("..." if len(val) > 70 else "")))
        if res["backup"]:
            print("  备份: %s" % res["backup"])
        print("  结果: %s" % ("成功" if res["ok"] else "失败/无写入"))
        if not res["ok"]:
            rc = 1
    return rc


def cmd_reindex(args):
    cache = args.out or CACHE
    corpus = Corpus(ARCHIVE_ROOTS, None)   # 强制重扫
    with open(cache, "w", encoding="utf-8") as fh:
        json.dump(corpus.records, fh, ensure_ascii=False, indent=1)
    st = corpus.stats()
    print("重建语料库: %s" % cache)
    print("  文件总数 %d | 已填原因 %d | 待填 %d" % (st["total"], st["filled"], st["pending"]))
    return 0


def cmd_doctor(args):
    """环境自检：依赖、归档路径、Excel 可用性。"""
    ok = True
    print("=" * 62)
    print("燃气工程变更原因生成 —— 环境自检")
    print("=" * 62)

    print("\n[1] Python 版本")
    v = sys.version_info
    print("    %d.%d.%d  %s" % (v.major, v.minor, v.micro, sys.executable))
    if v < (3, 8):
        print("    ✗ 需要 Python 3.8 及以上")
        ok = False
    else:
        print("    ✓ 满足要求")

    print("\n[2] 依赖库")
    try:
        import xlrd
        print("    ✓ xlrd %s" % xlrd.__VERSION__)
    except ImportError:
        print("    ✗ 缺少 xlrd   → pip install xlrd")
        ok = False
    try:
        import win32com.client  # noqa: F401
        print("    ✓ pywin32（写入 .xls 需要）")
    except ImportError:
        print("    ✗ 缺少 pywin32 → pip install pywin32")
        ok = False
    except Exception as exc:
        print("    ⚠ pywin32 已安装但导入异常：%s" % str(exc)[:60])
        print("      写入功能可能受限，请确认已安装 Microsoft Excel")

    print("\n[3] Excel（写入 .xls 需要本机安装 Excel）")
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        print("    ✓ Excel %s" % app.Version)
        app.Quit()
    except Exception as exc:
        print("    ✗ 不可用：%s" % str(exc)[:70])
        print("      仅查看/生成原因不受影响；写入需安装 Microsoft Excel")
        ok = False

    print("\n[4] 历史变更单归档目录")
    total_xls = 0
    for root in ARCHIVE_ROOTS:
        print("    · %s" % root)
        if not os.path.isdir(root):
            print("      ✗ 目录不存在")
            ok = False
            continue
        n = len(glob.glob(os.path.join(root, "**", "*.xls"), recursive=True))
        total_xls += n
        if n:
            print("      ✓ 找到 %d 份 .xls" % n)
        else:
            print("      ⚠ 目录存在但没有 .xls 文件")
    if not ARCHIVE_ROOTS:
        print("    ✗ 未配置任何归档目录")
        ok = False
    elif total_xls == 0 and ok:
        print("    ⚠ 所有归档目录中都没有 .xls —— 检索相似案例将不可用")
        print("      请设置环境变量 GAS_CHANGE_ARCHIVE（多个路径用 ; 分隔），")
        print('      或在 scripts/config.json 写入 {"archive_roots": ["路径1", "路径2"]}')

    print("\n[5] 语料库缓存")
    if os.path.exists(CACHE):
        try:
            import json
            d = json.load(open(CACHE, encoding="utf-8"))
            print("    ✓ %s（%d 条）" % (os.path.basename(CACHE), len(d)))
            print("      若归档目录有变动，请运行：python fill_change_reason.py reindex")
        except Exception as exc:
            print("    ⚠ 缓存无法读取：%s" % str(exc)[:50])
    else:
        print("    ⚠ 无缓存，请运行：python fill_change_reason.py reindex")

    print("\n" + "=" * 62)
    print("结论: %s" % ("环境就绪 ✓" if ok else "存在问题，请按上方 ✗ 项处理"))
    print("=" * 62)
    return 0 if ok else 1


def cmd_search(args):
    corpus = Corpus(ARCHIVE_ROOTS, CACHE)
    co = ChangeOrder(args.file)
    orders = co.orders()
    if not orders:
        print("[错误] 未找到变更单")
        return 2
    idx = max(1, args.order or 1)
    if idx > len(orders):
        print("[错误] 该文件共 %d 张变更单，--order=%d 超出范围" % (len(orders), idx))
        return 2
    o = orders[idx - 1]
    add, red = o["add"], o["red"]
    add_n = sorted(set(x["norm"] for x in add))
    red_n = sorted(set(x["norm"] for x in red))
    print("检索：%s 第%d张 %s %s" % (o["sheet"], idx, o["code"], o["name"]))
    for s, r in corpus.similar(add_n, red_n, top=args.top):
        print("[%.3f] %s" % (s, r["file"]))
        print("   原因: %s" % r["reason"])
        print()
    return 0


def cmd_update(args):
    """检查 / 安装技能更新。

    优先使用独立的 ``update.py``；若该文件不存在（旧版本），
    则回退到本文件内置的精简实现——否则旧版本永远无法自更新，
    陷入「要先更新才能更新」的死循环。
    """
    try:
        import update as upd
        if args.rollback:
            return upd.rollback()
        if args.check:
            remote, local, _has = upd.do_check()
            return 0 if remote else 2
        return upd.do_update()
    except ImportError:
        pass
    return _builtin_update(args)


# ---------------------------------------------------------------- 内置更新兜底
# 供「尚未包含 update.py 的旧版本」自更新使用。
# 逻辑与 update.py 一致但更精简：检查版本 → 下载 → 校验 → 逐文件覆盖 → 保留配置。

def _builtin_update(args):
    import shutil
    import tempfile
    import zipfile
    try:
        from urllib.request import Request, urlopen
        from urllib.error import HTTPError, URLError
    except ImportError:
        from urllib2 import Request, urlopen, HTTPError, URLError

    REPO = "lonyii/skills-gas-change-reason"
    RAW = "https://raw.githubusercontent.com/%s/main/VERSION" % REPO
    ZIP = "https://codeload.github.com/%s/zip/refs/heads/main" % REPO
    UA = "gas-change-reason-skill-updater"

    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ver_file = os.path.join(skill_dir, "VERSION")
    cfg_file = os.path.join(skill_dir, "scripts", "config.json")
    backup_dir = os.path.join(os.path.dirname(skill_dir),
                              ".%s.update_backup" % os.path.basename(skill_dir))

    def local_ver():
        try:
            with open(ver_file, "r", encoding="utf-8") as fh:
                return fh.read().strip() or "0.0.0"
        except Exception:
            return "0.0.0"

    def vt(v):
        p = re.findall(r"\d+", str(v or ""))
        while len(p) < 3:
            p.append("0")
        return tuple(int(x) for x in p[:3])

    print("[INFO] 本地版本: %s" % local_ver())
    _local_ver_cached = local_ver()
    try:
        req = Request(RAW, headers={"User-Agent": UA})
        with urlopen(req, timeout=20) as resp:
            remote = resp.read().decode("utf-8", "replace").strip()
    except Exception as e:
        print("[ERROR] 无法读取远端版本：%s" % e)
        print("       请手动下载：https://github.com/%s" % REPO)
        return 2
    print("[INFO] 远端版本: %s" % remote)

    if vt(remote) <= vt(local_ver()):
        print("[OK] 已是最新版本")
        return 0
    if args.check:
        print("[有更新] %s → %s" % (local_ver(), remote))
        return 0

    tmp = tempfile.mkdtemp(prefix="gcr_upd_")
    try:
        zp = os.path.join(tmp, "s.zip")
        try:
            req = Request(ZIP, headers={"User-Agent": UA})
            with urlopen(req, timeout=120) as resp, open(zp, "wb") as fh:
                shutil.copyfileobj(resp, fh)
        except Exception as e:
            print("[ERROR] 下载失败：%s" % e)
            return 2
        new_dir = os.path.join(tmp, "new")
        os.makedirs(new_dir, exist_ok=True)
        try:
            with zipfile.ZipFile(zp) as z:
                names = [n for n in z.namelist() if n and not n.startswith("__MACOSX")]
                root = set(n.split("/")[0] for n in names).pop()
                for n in names:
                    rel = n[len(root):].lstrip("/")
                    if not rel:
                        continue
                    dest = os.path.join(new_dir, rel.replace("/", os.sep))
                    if n.endswith("/"):
                        if not os.path.isdir(dest):
                            os.makedirs(dest)
                        continue
                    d = os.path.dirname(dest)
                    if d and not os.path.isdir(d):
                        os.makedirs(d)
                    with z.open(n) as s, open(dest, "wb") as o:
                        shutil.copyfileobj(s, o)
        except Exception as e:
            print("[ERROR] 解压失败：%s" % e)
            return 2

        # 校验必需文件
        for must in ("SKILL.md", "scripts/gas_change_core.py",
                     "scripts/material_reason.py", "scripts/fill_change_reason.py",
                     "scripts/bgd_toolkit.py", "scripts/gas_flow.py",
                     "scripts/xchange_db.py", "scripts/xchange_cache.json"):
            if not os.path.exists(os.path.join(new_dir, must.replace("/", os.sep))):
                print("[ERROR] 新版本缺少必需文件 %s，更新中止" % must)
                return 2
        print("[OK] 新版本校验通过")

        # 备份旧版（逐文件复制，避免 Windows 目录占用）
        cfg = None
        if os.path.exists(cfg_file):
            try:
                import json as _json
                with open(cfg_file, "r", encoding="utf-8") as fh:
                    cfg = _json.load(fh)
            except Exception:
                cfg = None
        try:
            if os.path.isdir(backup_dir):
                shutil.rmtree(backup_dir, ignore_errors=True)
            for r, ds, fs in os.walk(skill_dir):
                ds[:] = [d for d in ds if d != "__pycache__"]
                rel = os.path.relpath(r, skill_dir)
                td = backup_dir if rel == "." else os.path.join(backup_dir, rel)
                if not os.path.isdir(td):
                    os.makedirs(td)
                for f in fs:
                    if f == "config.json" or f.endswith((".pyc", ".pyo")):
                        continue
                    try:
                        shutil.copy2(os.path.join(r, f), os.path.join(td, f))
                    except Exception:
                        pass
        except Exception as e:
            print("[ERROR] 备份失败：%s" % e)
            return 2

        # 逐文件覆盖
        for r, ds, fs in os.walk(new_dir):
            ds[:] = [d for d in ds if d != "__pycache__"]
            for f in fs:
                if f == "config.json" or f.endswith((".pyc", ".pyo")):
                    continue
                src = os.path.join(r, f)
                dst = os.path.join(skill_dir, os.path.relpath(src, new_dir))
                d = os.path.dirname(dst)
                if d and not os.path.isdir(d):
                    os.makedirs(d)
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    try:
                        if os.path.exists(dst):
                            os.remove(dst)
                        shutil.copy2(src, dst)
                    except Exception as e:
                        print("[ERROR] 覆盖失败 %s：%s" % (f, e))
                        return 1
        if cfg:
            try:
                with open(cfg_file, "w", encoding="utf-8") as fh:
                    import json as _json
                    _json.dump(cfg, fh, ensure_ascii=False, indent=2)
                print("[OK] 已保留本地归档目录配置")
            except Exception:
                pass
        print("[OK] 更新完成：%s → %s" % (_local_ver_cached, remote))
        return 0
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="燃气工程变更单变更原因生成")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("analyze", help="分析变更单并预览原因（支持文件/目录批量）")
    p.add_argument("files", nargs="+", help="变更单 .xls 文件或目录（可多个）")
    p.add_argument("--location", default="", help="现场位置描述，如「南昌县老城区」")
    p.add_argument("--extra", default="", help="补充的现场事实描述")
    p.add_argument("--summary", action="store_true", help="只看总览，不打印正文")
    p.add_argument("--detail", action="store_true", help="附材料明细与动作推理")
    p.set_defaults(func=cmd_analyze)

    p = sub.add_parser("fill", help="生成并写入变更原因（支持文件/目录批量、多张变更单）")
    p.add_argument("files", nargs="+", help="变更单 .xls 文件或目录（可多个）")
    p.add_argument("--location", default="")
    p.add_argument("--extra", default="")
    p.add_argument("--body", default="", help="直接指定正文，覆盖自动生成")
    p.add_argument("--order", type=int, default=None,
                   help="只处理该文件中的第 N 张变更单（1 起）")
    p.add_argument("--overwrite", action="store_true",
                   help="授权覆盖已有原因（--body 本身即授权）")
    p.add_argument("--no-backup", action="store_true",
                   help="显式关闭备份（默认已关闭；用于覆盖 --backup-dir）")
    p.add_argument("--backup-dir", default="",
                   help="备份目录；指定后才开启备份（默认不生成备份文件）")
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser("reindex", help="重建历史语料库")
    p.add_argument("--out", default="")
    p.set_defaults(func=cmd_reindex)

    p = sub.add_parser("polish", help="输出润色素材包（草稿+材料+案例+文风），供 AI 润色")
    p.add_argument("file")
    p.add_argument("--order", type=int, default=1, help="该文件中的第 N 张变更单（1 起）")
    p.add_argument("--top", type=int, default=5)
    p.set_defaults(func=cmd_polish)

    p = sub.add_parser("doctor", help="环境自检（依赖/归档目录/Excel）")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("search", help="检索相似历史案例")
    p.add_argument("file")
    p.add_argument("--order", type=int, default=1, help="该文件中的第 N 张变更单（1 起）")
    p.add_argument("--top", type=int, default=8)
    p.set_defaults(func=cmd_search)

    p = sub.add_parser("update", help="检查/安装技能更新（从 GitHub 仓库）")
    p.add_argument("--check", action="store_true", help="只检查是否有新版本，不更新")
    p.add_argument("--rollback", action="store_true", help="回滚到上一个版本")
    p.set_defaults(func=cmd_update)

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
