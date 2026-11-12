# 燃气工程变更原因生成

**作者：lonyii（lonyii@qq.com）**

> **首次安装（二选一）**：
> ① 把 `gas-change-reason.zip` 交给你的 AI，说一句「帮我安装这个技能包」；
> ② 只给仓库地址 `https://github.com/lonyii/skills-gas-change-reason`，让 AI 从 GitHub 拉取安装。
> 之后每次把变更单发给它，说「帮我填写这份变更单的变更原因」即可。

给一份**只有变更内容、变更原因为空**的燃气工程设计变更单，本技能参照公司 314 份历史变更单的写法，
生成规范的变更原因并写入 `*-变更单打印` 表的 `[变更原因]` 单元格（每张单据两处：申请单 + 工程量变更单）。
一个 Excel 里有多张变更单（跨 sheet 或同一 sheet 纵向多张）时会逐张识别、逐张出文、逐张写入。
默认直接覆盖、不生成备份文件；需要留存原件时用 `--backup-dir` 指定目录主动开启。

| | |
|---|---|
| ✍️ 作者 | lonyii（lonyii@qq.com） |
| 📦 仓库 | <https://github.com/lonyii/skills-gas-change-reason> |
| 🔄 自动更新 | `python scripts/fill_change_reason.py update` |
| 📌 版本 | 见仓库根的 `VERSION` 文件 |

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

### 步骤 3 — 直接从 GitHub 安装（用户没给 zip 时）

```bash
git clone --depth 1 https://github.com/lonyii/skills-gas-change-reason.git
python skills-gas-change-reason/scripts/install.py --from-github
```

或者不克隆、直接让安装脚本自己拉：

```bash
python install.py --from-github
```

`--from-github` 会先从仓库下载源码再安装，其余流程（依赖安装、归档探测、自检）与本地安装一致。

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

> ⚠️ 直接跑 `fill`（不带 `--body`）遇到已有内容会**中止并报错**——这是保护机制，
> 不要绕过它。要覆盖必须走 `--body`，那是"明确指定新正文"的语义。

### 步骤 5 — 确认

安装脚本末尾会打印 `安装完成 ✓` 和实际安装路径。若自检有 ✗ 项，重跑：

```bash
python <安装路径>/scripts/fill_change_reason.py doctor
```

### 步骤 6 — 交给用户

告知：装好了，以后直接把变更单路径发给你，说「帮我填写这份变更单的变更原因」即可。
需要更新版本时可以说「更新一下变更原因技能」。

---

## 版本更新

```bash
cd <安装路径>/scripts

# 检查是否有新版本（不改任何文件）
python fill_change_reason.py update --check

# 更新到最新版
python fill_change_reason.py update

# 更新出问题可回滚
python fill_change_reason.py update --rollback
```

更新走 GitHub 公开仓库，**不需要账号或 Token**。机制要点：

- 版本号在仓库根的 `VERSION` 文件
- 下载后先校验必需文件齐全，再原子替换；失败自动回滚
- **本地归档目录配置（`config.json`）会保留**，升级不会丢
- 已是最新版本时不做任何写入

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

# 先看看，不写文件（也可以直接给一个目录，批量预览）
python fill_change_reason.py analyze "变更单.xls"
python fill_change_reason.py analyze "D:\桌面\待办" --summary

# 确认后写入（默认不生成备份；需备份时加 --backup-dir "目录"）
python fill_change_reason.py fill "变更单.xls"

# 文件里有多张单据时，只看/只改其中一张
python fill_change_reason.py fill "变更单.xls" --order 2

# 补充材料推断不出的现场原因（最影响准确度）
python fill_change_reason.py fill "变更单.xls" --extra "因厂方要求将调压柜移位"

# 检查 / 安装更新
python fill_change_reason.py update --check
python fill_change_reason.py update
```

### 统一入口（推荐，DWG 或平差表都能直接喂）

```bash
cd <安装路径>/scripts

# 从 DWG 开始：自动转材料平差表 → 生成变更单 → 填变更原因
python gas_flow.py "<施工图.dwg>" "<竣工图.dwg>"

# 从材料平差表开始：已有变更单就跳过生成，直接填原因
python gas_flow.py "<材料平差.xls>"

# 只有用户明确说「单独生成变更单」时才导出独立 xls
python gas_flow.py "<材料平差.xls>" --standalone

# 原因已填时：标准结尾语 → 无需处理(rc=0)；非标准 → 逐张输出润色素材包(rc=10，提示 AI 用 fill --body 回写)
```

默认所有改动都写回**同一份材料平差表**（新增 sheet + 填原因），不另建变更单文件。
`--out-dir` 可指定 DWG 转平差的成品目录；DWG 入口需要 `dwg-material-to-xls` 技能已安装。

### 也可以单独跑生成脚本（不用再跑 VBA 宏）

材料平差表里本来要运行的宏（`CreateBianGengdan` / 广联达 / 神机导出）已全部改写成
Python，**不用开宏、不用宏安全性**：

```bash
cd <安装路径>/scripts

# 生成「<编号>-变更单打印」——替代原来的 CreateBianGengdan 宏
python bgd_toolkit.py bgd "材料平差表.xls"

# 需要广联达 / 神机导入表时
python bgd_toolkit.py gldon  "材料平差表.xls"
python bgd_toolkit.py shenji "材料平差表.xls"

# 三个一起做
python bgd_toolkit.py all "材料平差表.xls" --target gldon+shenji+bgd

# 看套价数据库缓存状态（默认直读缓存，源 XchangeSheet.xls 更新后自动重建）
python bgd_toolkit.py db
```

| 要点 | 说明 |
|---|---|
| 数据来源 | 「2-汇总」表的 J 列（变量）/ K-L（增加）/ M-N（减少），与宏一致 |
| 模板 | 只在运行时复制一份 `@变更单EXCEL模板` 工作表，**模板文件本身不动** |
| 速度 | 套价数据库走 `xchange_cache.json`，加载约 10ms（原来每次要开 Excel 读 3.9MB 的 .xls） |
| 等价性 | 与宏逐格比对：6 组变更单用例 + 4 组导出用例，全部 0 差异（含打印区域与分页符） |
| 默认写回 | 结果直接写回原平差表（新增 sheet），不另建变更单 xls 文件 |
| 独立导出 | 显式 `--standalone` 或 `export` 子命令才导出只含变更单的独立 xls |
| 原因已填判定 | 按结尾语：全标准 → 无需处理(rc=0)；有非标准 → 输出润色素材包(rc=10) |

### 命令速查

| 命令 | 作用 | 改文件 |
|---|---|---|
| `doctor` | 环境自检 | 否 |
| `analyze <文件或目录>` | 预览原因 + 每张单据的现状/因子/是否命中同项目定稿 | 否 |
| `polish <文件>` | **输出润色素材包**（草稿+材料+案例+文风），供 AI 润色 | 否 |
| `fill <文件或目录>` | 生成并写入变更原因（逐张单据） | **是**（直接覆盖，默认不备份） |
| `search <文件>` | 检索相似历史案例 | 否 |
| `reindex` | 重建历史语料库缓存 | 改缓存 |
| `update [--check]` | 检查 / 安装技能更新 | **是**（替换技能目录，保留本地配置） |
| `update --rollback` | 回滚到上一个版本 | **是** |

`fill` / `analyze` 可选参数：

| 参数 | 说明 |
|---|---|
| `--extra "..."` | 补充材料推断不出的现场事实（甲方要求、阻工、库房缺货…） |
| `--location "..."` | 覆盖自动推断的现场位置 |
| `--body "..."` | 直接指定正文，覆盖自动生成（**AI 润色回写用这个**；多张单据须配 `--order`） |
| `--order N` | 只处理文件里的第 N 张变更单（1 起） |
| `--overwrite` | 授权覆盖已有原因 |
| `--backup-dir "..."` | 指定备份目录并开启备份（默认不备份）；`--no-backup` 显式关闭 |
| `--summary` / `--detail` | `analyze` 只看总览 / 附材料明细与动作推理 |

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
    ↓  先看有没有同项目历史定稿（有就沿用其口径，只补本次新材料动作）
    ↓  依据工程编号类别调整措辞 + 长度自适应（≤234 字）
变更原因正文
    ↓  申请单结尾语 / 工程量变更单结尾语
每张单据的两处原因格（各写各的）
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
- **备份默认关闭**：默认直接覆盖、不生成备份文件。需要留存原件时显式
  `--backup-dir "目录"` 才会先复制（文件名带时间戳），备份失败会中止写入。三种情形行为不同：
  - 目标格为空 → `fill` 自动生成并写入
  - 目标格已有内容 → `fill`（不带 `--body`/`--overwrite`）**会跳过该张**并提示改走 polish 流程
  - `polish --order N` 润色后 → `fill --order N --body "<润色结果>"` 授权覆盖
- 目前针对 `.xls` 老格式设计，`.xlsx` 需另做适配。

## 目录内容

| 文件 | 说明 |
|---|---|
| `SKILL.md` | 技能定义，Agent 据此发现与调用 |
| `README.md` | 本文件 |
| `VERSION` | 版本号（更新时据此判断是否有新版） |
| `CHANGELOG.md` | 更新日志（每次版本的可感知变化） |
| `scripts/install.py` | 一键安装脚本（`--from-github` 支持从仓库直接装） |
| `scripts/update.py` | 自动更新：检查版本 → 下载 → 校验 → 原子替换 → 失败回滚 |
| `scripts/gas_change_core.py` | 变更单读写 + 语料库 |
| `scripts/material_reason.py` | 材料级推理引擎 |
| `scripts/project_code.py` | 工程编号分类映射 |
| `scripts/fill_change_reason.py` | 命令行入口 |
| `scripts/gas_flow.py` | **统一入口**：DWG / 平差表自动判断从哪步开始，一路做到填好变更原因 |
| `scripts/bgd_toolkit.py` | **生成变更单 / 广联达导入 / 神机导入**（替代材料平差表里的 VBA 宏） |
| `scripts/xchange_db.py` | 套价数据库 JSON 缓存层（读取/重建/精简/定额查询） |
| `scripts/xchange_cache.json` | 套价数据库缓存（替换规则 253 条 / 定额 5069 行 / 分隔符 20 条） |
| `scripts/corpus_cache.json` | 历史语料缓存（314 条） |
| `scripts/config.example.json` | 归档目录配置模板 |
| `references/patterns.md` | 310 份历史原因原文，按因子归类 |
| `references/corpus_summary.md` | 语料统计与典型案例 |
