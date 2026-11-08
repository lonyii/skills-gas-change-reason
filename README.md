# 燃气工程变更原因生成

> **首次安装：把 `gas-change-reason.zip` 交给你的 AI，说一句「帮我安装这个技能包」即可自动装好（依赖安装、路径配置、自检校验全部自动完成）；之后每次把变更单发给它，说「帮我填写这份变更单的变更原因」即可。**

给一份**只有变更内容、变更原因为空**的燃气工程设计变更单，本技能参照公司 314 份历史变更单的写法，
生成规范的变更原因并写入 `*-变更单打印` 表的 `[变更原因]` 单元格（申请单 B5 + 工程量变更单 A49）。

---

## 给 AI：安装这个技能包

> 用户已把 `gas-change-reason.zip` 交给你。请按以下步骤完成安装，**不要让用户手动操作**。

### 步骤 1 — 解压

把 zip 解压到工作区临时目录（如 `<tmp>/gas-change-reason/`）。zip 内已含顶层目录 `gas-change-reason/`。

### 步骤 2 — 一键安装

```bash
python <解压路径>/gas-change-reason/scripts/install.py
```

脚本会自动完成：Python 版本检查 → 缺失依赖的 pip 安装 → 复制到用户级 skills 目录
（`C:\Users\<用户名>\.workbuddy\skills\gas-change-reason\`）→ 自动探测历史变更单归档目录并写入配置
→ 运行 doctor 自检。

**若脚本输出 `ERROR`**：按提示处理后重跑。常见情况：

| 情况 | 处理 |
|---|---|
| 找不到 `python` | 试 `py -3` 或 `python3`；仍不行则查找已安装的 Python |
| pip 安装失败 | 用 `python -m pip install xlrd pywin32` 手动装，然后 `python install.py --no-deps` |
| 未能定位归档目录 | 脚本会用随包语料，功能正常。如需指定：`python install.py --archive "<路径>"` |
| Excel 不可用 | 仅「写入」功能受阻（`.xls` 需 Excel COM）。生成与分析不受影响 |

指定归档目录或安装到别处：

```bash
python install.py --archive "D:\你的路径\设计变更单归档"
python install.py --target "D:\skills\gas-change-reason"
```

### 步骤 4 — 处理变更单

收到变更单后，**先看变更原因单元格有没有内容**：

```bash
python <安装路径>/scripts/fill_change_reason.py analyze "<变更单.xls>"
```

- **单元格为空 / 是 `[变更原因]` 占位符** → 自动生成：
  ```bash
  python fill_change_reason.py fill "<变更单.xls>"
  ```
- **单元格已有用户手填的原因** → **当作提示词，由 AI 润色后写回**（不要直接覆盖）：
  ```bash
  # 1) 取润色素材包
  python fill_change_reason.py polish "<变更单.xls>"
  # 2) 参照素材包润色（保留草稿中的现场事实，只改措辞）
  # 3) 写回
  python fill_change_reason.py fill "<变更单.xls>" --body "<润色后的正文>"
  ```

### 步骤 5 — 确认

安装脚本末尾会打印 `安装完成 ✓` 和实际安装路径。若自检有 ✗ 项，重跑：

```bash
python <安装路径>/scripts/fill_change_reason.py doctor
```

### 步骤 6 — 交给用户

告知：装好了，以后直接把变更单路径发给你，说「帮我填写这份变更单的变更原因」即可。

---

## 给使用者：怎么用

装好后，**直接把变更单发给 Agent**：

> 帮我填写这份变更单的变更原因：D:\xxx\变更单.xls

Agent 会自动判断走哪条路径：

| 单元格状态 | Agent 的动作 |
|---|---|
| 空 / 是 `[变更原因]` 占位符 | 自动生成 → 预览 → 写入 |
| **已有你手填的原因** | **当作提示词**，参照历史语料润色 → 写回 |

或用命令行：

```bash
cd <安装路径>/scripts

# 先看看，不写文件
python fill_change_reason.py analyze "变更单.xls"

# 确认后写入（直接覆盖原文件，不留备份）
python fill_change_reason.py fill "变更单.xls"

# 补充材料推断不出的现场原因（最影响准确度）
python fill_change_reason.py fill "变更单.xls" --extra "因厂方要求将调压柜移位"
```

### 命令速查

| 命令 | 作用 | 改文件 |
|---|---|---|
| `doctor` | 环境自检 | 否 |
| `analyze <文件>` | 预览原因 + 工程动作推理过程 | 否 |
| `polish <文件>` | **输出润色素材包**（草稿+材料+案例+文风），供 AI 润色 | 否 |
| `fill <文件>` | 生成并写入变更原因 | **是**（直接覆盖，不备份） |
| `search <文件>` | 检索相似历史案例 | 否 |
| `reindex` | 重建历史语料库缓存 | 改缓存 |

`fill` / `analyze` 可选参数：

| 参数 | 说明 |
|---|---|
| `--extra "..."` | 补充材料推断不出的现场事实（甲方要求、阻工、库房缺货…） |
| `--location "..."` | 覆盖自动推断的现场位置 |
| `--body "..."` | 直接指定正文，覆盖自动生成（**AI 润色回写用这个**） |

### 你自己动手润色的话

```bash
# 1) 拿到素材包（内有你写的草稿、相关历史案例、文风规范）
python fill_change_reason.py polish "变更单.xls"

# 2) 照着规范改一改

# 3) 写回
python fill_change_reason.py fill "变更单.xls" --body "你润色后的正文"
```

> 润色原则：**只改措辞，不改事实**。草稿里的甲方要求、阻工、库房缺货等现场信息要保留，
> 把它改写成历史语料的习惯句式（如「因…，经报告公司相关代表到现场核实并商议后决定，改为…」）。

---

## 它怎么工作

```
增/减材料清单
    ↓  解析规格（DN65 / dn110 / D76 / D89×4.5 等多种命名）
工程动作推导
    ↓  pipe_swap   管段替换    量级相当 + 公称直径接近
    ↓  resize      接口改径    同基类 + 主规格不同 + 1换1
    ↓  branch_add  支管增设    同规格小管件 +≥3 个
    ↓  cap_reserve 预留封堵    法兰+盲板(+垫片)成套出现
    ↓  instrument  仪表连接    出现单头牙 / 过滤网
    ↓  protect     埋深防护    出现保护沟
    ↓  anti_impact 防撞护栏    出现防撞/护栏
    ↓  依据工程编号类别调整措辞 + 长度自适应（≤234 字）
变更原因正文
    ↓  申请单结尾语 / 工程量变更单结尾语
B5 + A49
```

### 工程编号含义

编号结构 `{公司码}{年份}{类别码}{顺序号}`，如 `NC25E075` = 南昌市 · 2025年 · 商业 · 第075号。

类别码取自《工程建设项目分类、编码及命名指引》附表1：

| 码 | 细分类型 | 码 | 细分类型 |
|---|---|---|---|
| A/B/C | 新建/建成小区、民用拆改迁 | H/I | 市政管网、市政管网拆改迁 |
| D | 工业 | J | 零星工程 |
| E | 商业 | K | 高压/次高压线路 |
| F | 公福 | L | 输配/加气/特种场站 |

区域码：`NC` 南昌市、`IC` 南昌县、`ID` 新建区、`IE` 赣江新区、`VX` 经开区。

场景措辞按类别分化：小区→「该小区内」、工业→「该厂区内」、商业→「该商业场所内」、
公福→「该公共建筑内」、市政→「该市政管网沿线」。

> 名称里的具体场景优先于类别码——锅炉类工程横跨工业(D)与商业(E)。

---

## 注意事项

- **写入需要本机安装 Microsoft Excel**（`.xls` 老格式 + Excel COM 写入）。只生成不写入则不需要。
- **材料只能反映「改了什么」，反映不了「为什么改」**。现场真实原因请用 `--extra` 补充，这一步最影响准确度。
- **写入直接覆盖原文件，不生成 `.备份.xls` 副本**。目标格已有内容时自动生成模式会拒绝写入（防误操作）；用 `--body` 显式指定时允许覆盖，此时会在终端打印告警提示原内容将被永久替换。
- 目前针对 `.xls` 老格式设计，`.xlsx` 需另做适配。

## 目录内容

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 技能定义，Agent 据此发现与调用 |
| `README.md` | 本文件 |
| `scripts/install.py` | 一键安装脚本 |
| `scripts/gas_change_core.py` | 变更单读写 + 语料库 |
| `scripts/material_reason.py` | 材料级推理引擎 |
| `scripts/project_code.py` | 工程编号分类映射 |
| `scripts/fill_change_reason.py` | 命令行入口 |
| `scripts/corpus_cache.json` | 历史语料缓存（314 条） |
| `scripts/config.example.json` | 归档目录配置模板 |
| `references/patterns.md` | 310 份历史原因原文，按因子归类 |
| `references/corpus_summary.md` | 语料统计与典型案例 |
