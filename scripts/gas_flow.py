#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""燃气工程统一入口：从 DWG 图纸或材料平差表出发，一路做到「变更原因」填好。

设计要点
--------
1. **多入口自动判断**：
   - 输入 .dwg/.dxf（或装着图纸的目录）→ 先调用 dwg-material-to-xls 生成材料平差表；
   - 输入材料平差 .xls → 直接用。
2. **已有变更单就跳过生成**：平差表里若已有 `*-变更单打印`，不再新建，直接拿它去填原因。
3. **默认就地**：所有改动（新增变更单 sheet + 变更原因）都写回同一份材料平差表，
   不额外生成「变更单 xls 文件」。只有显式 `--standalone` 才导出只含变更单的独立文件。
4. 三步顺序不变：材料平差 → 变更单 → 变更原因。

用法
----
    python gas_flow.py <DWG/目录/材料平差.xls> [选项]

选项
----
    --out-dir DIR      DWG 转平差表时的成品目录（默认图纸目录下的「材料平差输出」）
    --standalone       额外导出只含变更单的独立 xls（默认不导出）
    --out FILE         独立变更单文件路径（配合 --standalone，默认同目录 <原名>-变更单.xls）
    --sheet NAME       指定源数据表（默认「2-汇总」）
    --name NAME        指定新建变更单 sheet 名
    --extra "...",     补充材料推断不出的现场事实
    --location "..."   覆盖自动推断的现场位置
    --no-save          只预览，不写文件

依赖：本技能的 bgd_toolkit.py / fill_change_reason.py，以及 dwg-material-to-xls 技能（仅 DWG 入口需要）。
"""
from __future__ import print_function

__author__ = "lonyii"
__email__ = "lonyii@qq.com"

import glob
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))

DWG_EXT = (".dwg", ".dxf")
XLS_EXT = (".xls", ".xlsx", ".xlsm")
CHANGE_SHEET_MARK = "变更单打印"


def out(msg):
    print(msg)
    sys.stdout.flush()


def err(msg):
    print("[错误] %s" % msg, file=sys.stderr)
    sys.stderr.flush()


def run(cmd, label, capture=False, allow_fail=False):
    """跑子进程；capture 时收集输出以便判断，allow_fail 时不直接抛错。"""
    out("")
    out("---- %s ----" % label)
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    if capture:
        p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             env=env, universal_newlines=True,
                             encoding="utf-8", errors="replace")
        text = p.communicate()[0] or ""
        out(text.rstrip())
        if p.returncode and not allow_fail:
            raise RuntimeError("%s 失败（退出码 %d）" % (label, p.returncode))
        return p.returncode, text
    rc = subprocess.call(cmd, env=env)
    if rc and not allow_fail:
        raise RuntimeError("%s 失败（退出码 %d）" % (label, rc))
    return rc, ""


def find_dwg_exporter():
    """定位 dwg-material-to-xls 的 auto_export.py。"""
    env = os.environ.get("GAS_DWG_SKILL_DIR")
    candidates = []
    if env:
        candidates.append(os.path.join(env, "scripts", "auto_export.py"))
    parent = os.path.dirname(HERE)
    candidates.append(os.path.join(parent, "dwg-material-to-xls", "scripts", "auto_export.py"))
    home = os.path.expanduser("~")
    candidates.append(os.path.join(home, ".workbuddy", "skills", "dwg-material-to-xls",
                                   "scripts", "auto_export.py"))
    for c in candidates:
        if os.path.isfile(c):
            return c
    err("找不到 dwg-material-to-xls 技能（auto_export.py）。"
        "请先安装，或用环境变量 GAS_DWG_SKILL_DIR 指定其目录。")
    return None


def sheet_names(path):
    """只读列出工作表名（.xls 走 xlrd，.xlsx/.xlsm 走 zip+xml）。"""
    if path.lower().endswith(".xls"):
        try:
            import xlrd
        except ImportError:
            raise RuntimeError("需要 xlrd：pip install xlrd")
        return list(xlrd.open_workbook(path).sheet_names())
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    return re.findall(r'<sheet[^>]*\bname="([^"]+)"', xml)


def has_change_sheet(path):
    return any(CHANGE_SHEET_MARK in n for n in sheet_names(path))


def expand_inputs(inputs):
    """目录展开为图纸/表格清单；返回 (dwgs, xls)。"""
    dwgs, xls = [], []
    for p in inputs:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                full = os.path.join(p, name)
                if not os.path.isfile(full):
                    continue
                low = name.lower()
                if low.endswith(DWG_EXT):
                    dwgs.append(full)
                elif low.endswith(XLS_EXT) and not low.startswith("~$"):
                    xls.append(full)
        else:
            low = p.lower()
            if low.endswith(DWG_EXT):
                dwgs.append(p)
            elif low.endswith(XLS_EXT):
                xls.append(p)
            else:
                err("不支持的输入类型：%s" % p)
    return dwgs, xls


def new_xls_in(directory, before):
    """找出本轮新增的 .xls（排除 Excel 临时文件）。"""
    pattern = os.path.join(directory, "*.xls*")
    after = set(p for p in glob.glob(pattern) if not os.path.basename(p).startswith("~$"))
    added = sorted(after - before)
    return added


def default_standalone_path(xls_path, out=None):
    if out:
        return os.path.abspath(out)
    base = os.path.splitext(os.path.basename(xls_path))[0]
    return os.path.join(os.path.dirname(os.path.abspath(xls_path)), base + "-变更单.xls")


def parse(argv):
    """轻量参数解析：选项与其取值成对消费，剩余才是输入路径。"""
    if not argv or argv[0] in ("-h", "--help"):
        return None
    value_opts = ("--out-dir", "--out", "--sheet", "--name", "--extra", "--location")
    flags = ("--standalone", "--no-save")
    values = {}
    switches = {}
    inputs = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in value_opts:
            if i + 1 < len(argv):
                values[a] = argv[i + 1]
                i += 2
            else:
                i += 1
        elif a in flags:
            switches[a] = True
            i += 1
        elif a.startswith("-"):
            i += 1
        else:
            inputs.append(a)
            i += 1
    return {
        "inputs": inputs,
        "out_dir": values.get("--out-dir"),
        "standalone": switches.get("--standalone", False),
        "out": values.get("--out"),
        "sheet": values.get("--sheet"),
        "name": values.get("--name"),
        "extra": values.get("--extra"),
        "location": values.get("--location"),
        "no_save": switches.get("--no-save", False),
    }


def main(argv):
    args = parse(argv)
    if args is None:
        out(__doc__)
        return 0 if argv else 2
    if not args["inputs"]:
        err("缺少输入：请给 DWG/目录/材料平差 .xls")
        return 2

    dwgs, xls = expand_inputs(args["inputs"])
    if not dwgs and not xls:
        err("没有找到 .dwg/.dxf 或材料平差 .xls")
        return 2

    py = sys.executable
    bgd_tool = os.path.join(HERE, "bgd_toolkit.py")
    fill_tool = os.path.join(HERE, "fill_change_reason.py")

    # ---- 入口 1：DWG → 先生成材料平差表
    if dwgs:
        exporter = find_dwg_exporter()
        if not exporter:
            return 2
        out_dir = os.path.abspath(args["out_dir"] or
                                  os.path.join(os.path.dirname(os.path.abspath(dwgs[0])),
                                               "材料平差输出"))
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
        before = set(new_xls_in(out_dir, set()))
        out("[1/3] 从 %d 张 DWG 生成材料平差表 → %s" % (len(dwgs), out_dir))
        run([py, exporter] + dwgs + ["--out-dir", out_dir], "DWG → 材料平差")
        added = new_xls_in(out_dir, before)
        if len(added) != 1:
            err("本轮生成了 %d 份平差表，无法确定用哪一份：%s\n"
                "请用 --out-dir 指定独立目录，或直接把材料平差 .xls 传给本命令。"
                % (len(added), ", ".join(added)))
            return 2
        xls_path = added[0]
    else:
        xls_path = os.path.abspath(xls[0])
        if len(xls) > 1:
            out("[提示] 检测到 %d 份平差表，本次处理：%s" % (len(xls), xls_path))

    out("[输入] %s" % xls_path)
    standalone_path = default_standalone_path(xls_path, args["out"]) if args["standalone"] else None

    # ---- 入口 2：平差表（有/无变更单都行）
    exists = has_change_sheet(xls_path)
    if exists:
        out("[2/3] 已有变更单 sheet，跳过生成，直接用于填写变更原因")
        if standalone_path:
            run([py, bgd_tool, "export", xls_path, "--out", standalone_path],
                "导出独立变更单")
            work = standalone_path
        else:
            work = xls_path
    else:
        out("[2/3] 未发现变更单，先生成（默认写回原平差表，不另建文件）")
        cmd = [py, bgd_tool, "bgd", xls_path]
        for flag, key in (("--sheet", "sheet"), ("--name", "name")):
            if args[key]:
                cmd += [flag, args[key]]
        if standalone_path:
            cmd += ["--standalone", "--out", standalone_path]
        if args["no_save"]:
            cmd.append("--no-save")
        run(cmd, "生成变更单")
        work = standalone_path if standalone_path else xls_path

    # ---- 最后一步：填变更原因
    out("[3/3] 填写变更原因")
    cmd = [py, fill_tool, "fill", work]
    for flag, key in (("--extra", "extra"), ("--location", "location")):
        if args[key]:
            cmd += [flag, args[key]]
    rc, text = run(cmd, "填写变更原因", capture=True, allow_fail=True)
    if rc and ("已有内容" in text or "未授权覆盖" in text or "无写入" in text):
        out("[INFO] 变更原因已填写，无需处理（要润色请用 polish → fill --body）")
        rc = 0
    if rc:
        raise RuntimeError("填写变更原因 失败（退出码 %d）" % rc)

    out("")
    out("[完成] 成品：%s" % work)
    if standalone_path and not exists:
        out("        原平差表保持原样；独立变更单：%s" % standalone_path)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except RuntimeError as exc:
        err(str(exc))
        sys.exit(1)
    except KeyboardInterrupt:
        err("已中断")
        sys.exit(130)
