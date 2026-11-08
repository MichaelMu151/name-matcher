# 政府部门名称编码与人工审核引擎

这是一个面向广州、深圳及其他地区组织史数据的部门名称标准化与审核程序。它将输入 Excel 中的历史部门名称，与 `codebook.xlsx` 中维护的标准部门名称、别名、部门类别、部门级别和有效期进行比较，并把结果分成自动编码、人工审核、名称有效期问题和多种可能等不同路径。

程序的设计目标不是“猜一个最像的名称”，而是让每一个正式编码都可以追溯到 codebook 中的具体记录，并让不确定结果进入人工审核。

## 主要用途

- 审核输入数据的 `w_name` 是否属于 codebook 的标准名称池。
- 使用 `w_name_o` 召回标准部门名称候选。
- 检查 `cityname`、`distname` 与 codebook 的适用范围。
- 比较输入已有的 `w_name`、`wrank_t`、`s1_sector` 与系统判断。
- 识别错误输入、字段冲突、多种可能和缺失编码。
- 根据 codebook 的有效期发现旧部门名称超出使用年份的情况。
- 处理部门级别、类别和名称在年份或月份节点上的变化。
- 根据最终确定的部门级别调整人员 `post_rank`。
- 将人工选择保存为 alias、反馈记录和后续可追踪的审核日志。

## 处理顺序

程序严格按照以下顺序工作：

```text
输入 Excel 与 codebook
        ↓
系统召回候选并生成初步审核队列
        ↓
人工复核 w_name / s1_sector / wrank_t
        ↓
检查 w_name 是否处于 codebook 有效期
        ↓
处理级别、类别和名称的年份/月度变动
        ↓
根据最终 wrank_t 调整 post_rank
        ↓
导出 reviewed_content、审核队列和修改日志
```

### 1. 人工复核

人工复核以名称组为单位进行。同一个 `w_name_o` 的多条记录会合并在候选卡片中，确认一次即可应用到该名称组的原始行。

每一次选择都会立即写入本地暂存文件 `data/pending_review.json`。只有点击“保存本批复核结果”后，才会更新原始 Excel、`reviewed_content.xlsx` 和 codebook。

人工选择可以是：

- 确认系统建议；
- 选择候选卡片；
- 保留输入中已经存在的编码；
- 手动新增标准名称；
- 无法匹配；
- 跳过并稍后处理。

### 2. 名称有效期检查

如果一个标准名称只在某一时间段有效，程序会把年份超出有效区间的记录单独列为 `w_name超出有效期`。

例如 `专利管理局` 的 codebook 有效期截止到 2001 年，但输入数据在 2002 年仍使用这个名称，程序不会擅自改成新名称，而是把它放入人工审核队列。

相关问题包括：

- `w_name超出有效期`；
- `w_name有效期无法判断`；
- 同一名称和年份对应多个不同编码的 `多种可能`。

### 3. 级别、类别和名称变动

对于处于有效期内的部门，程序依据 `year_valid_from/to` 和 `month_valid_from/to` 选择对应的 codebook 记录：

- 只有年份边界时，按年份填充，不复制年度行；
- 有月份边界时，使用 `styr/stmo` 与 `fiyr/fimo` 判断任职区间；
- 任职跨越月份节点时，在变动审阅结果中展示变化前后区间；
- 级别未变化但 `wrank20 != wrank_t` 时，标记为 `级别没变但是wrank不统一`。

旧名称在有效期外不会被当作普通级别变化自动修正。

### 4. post_rank 联动

只有前面步骤确定了最终 `wrank_t` 后，程序才会调整 `post_rank`。

内部计算使用统一的等级序列：

```text
副股级 → 股级 → 副科级 → 科级 → 副处级 → 处级
→ 副厅级 → 地厅级 → 副部级 → 省部级
```

历史数据中的名称会统一解释：

- `科级` = `正科级`；
- `处级` = `正处级`；
- `厅级`、`地厅级` = `正厅级`；
- `省部级` = `正部级`。

例如部门级别从 `副厅级` 变为 `处级`，变化量为 -1：

- `post_rank=副厅级` → `处级`；
- `post_rank=处级` → `副处级`。

程序会保留：

- `post_rank_original`；
- `post_rank_adjustment_reason`；
- `post_rank_adjustment_status`。

无法识别的职位级别、调整后超出范围的级别不会被强行修改，而会保留原值并标记为待审核。

## 安装

建议使用 Python 3.11–3.13：

```bash
cd /path/to/wname_matcher
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
python -m pip check
```

运行测试：

```bash
python -m pytest -q
```

当前测试覆盖 codebook 审计、候选匹配、导出安全校验和 `post_rank` 联动逻辑。

## 推荐使用方式：PySide6 桌面界面

```bash
python run_qt.py
```

界面主要页面：

1. **开始运行**：选择输入 Excel、`codebook.xlsx`、配置文件和可选模型；
2. **结果预览**：查看系统判断、字段冲突和问题类型；
3. **人工复核**：按名称组打开候选卡片，修改并暂存审核；
4. **变动审阅**：查看 codebook 中经历过名称、类别或级别变化的部门，检查名称有效期问题和 `post_rank` 示例；
5. **使用说明**：查看操作规则和输出文件说明。

首次使用建议：

1. 选择输入数据和 codebook；
2. 点击“开始运行”；
3. 进入“人工复核”，逐组处理候选；
4. 在“本批暂存”表检查选择；
5. 点击“保存本批复核结果”；
6. 再检查“变动审阅”中的名称有效期问题；
7. 查看输出目录中的审核队列和 `post_rank` 调整日志。

如果保存时出现 alias 冲突，程序会保存无冲突记录，并把冲突记录留在暂存区，而不会整批丢失。

## 命令行用法

### 审计 codebook

```bash
python cli.py audit \
  --codebook /path/to/codebook.xlsx \
  --reports-dir reports/codebook_audit
```

### 训练冷启动排序模型

```bash
python cli.py train \
  --codebook /path/to/codebook.xlsx \
  --output models/synthetic_ranker.joblib
```

模型只是候选排序器，不会生成不在 codebook 中的正式字段。

### 批量运行

```bash
python cli.py run \
  --codebook /path/to/codebook.xlsx \
  --input /path/to/input.xlsx \
  --output-dir outputs \
  --config config.yaml \
  --model models/synthetic_ranker.joblib
```

### 启动旧版 Streamlit 界面

```bash
python cli.py review-ui
```

Streamlit 版本保留用于兼容旧工作流。日常人工复核建议使用 PySide6，因为它支持批次暂存、上一条/下一条、候选卡片、alias 写入和变动审阅的连续流程。

## 输入数据要求

常用字段如下：

```text
record_id, year, cityname, distname,
w_name_o, w_name, wrank_t, wrank20, s1_sector,
p_name, post_rank, styr, stmo, fiyr, fimo
```

建议提供稳定的 `record_id`。如果没有，程序会根据源行和重复次序生成内部 ID。

任职时间由 `styr/stmo` 到 `fiyr/fimo` 共同决定。缺少月份时，年度边界按该年的中间月份用于审阅定位；codebook 的有效期则按起始年 1 月、结束年 12 月解释。

## Codebook 要求

`codebook.xlsx` 必须包含：

```text
cityname, distname, dtype, drank,
s1_sector, wrank, w_name, alias, notes,
year_valid_from, month_valid_from,
year_valid_to, month_valid_to
```

`alias` 支持 `；`、`;`、`|`、逗号、空格和换行分隔。标准名称始终是 `w_name`，真实数据中的其他称呼只能作为 alias，不会替代标准名称。

四个有效期字段全部为空，表示该 codebook 记录适用于所有年份。有效期字段存在时，程序会进行时间范围检查。

## DeepSeek API

PySide6 页面可以选择 DeepSeek API，并填写 API Key、模型名称、API 地址和本次调用上限。

API Key 只在本次运行内存中使用，不写入配置文件、日志或 Git。DeepSeek 只能在规则和排序器无法高置信决定时，对已有候选进行二次裁决；它不能生成候选白名单之外的部门名称。详细约束见 `skills/deepseek_judge/SKILL.md`。

## 输出文件

每次批处理会在指定输出目录生成：

```text
reviewed_content.xlsx             输入文件副本及系统结果
review_queue.xlsx                 错误输入、冲突、多种可能等问题
modification_log.xlsx             字段级 before/after 日志
temporal_review_queue.xlsx        名称有效期问题
post_rank_adjustment_log.xlsx     职位级别联动日志
debug/queries.jsonl               查询、候选和判定证据
run_meta.json                     本次运行的版本与配置元数据
```

人工审核保存时，程序还会更新原始输入 Excel、`reviewed_content.xlsx`、codebook 的 alias、`data/human_feedback.csv` 和 `data/pending_review.json`。

## 安全与可追溯性

- 正式字段只能通过 codebook 候选回查，不能由模型自由生成；
- 原始数据保存前会创建备份；
- 原始列顺序、行顺序和非目标单元格受到导出校验保护；
- alias 冲突不会被静默覆盖；
- 每次人工选择都先落盘到暂存批次；
- `post_rank` 调整保留原值、调整原因和状态；
- 无法识别或超出范围的级别进入人工处理，不做强制猜测。

## 目录结构

```text
cli.py                         命令行入口
run_qt.py                      PySide6 启动入口
config.yaml                    运行配置
codebook.xlsx                  本地标准 codebook，不应提交到公共仓库
src/engine/codebook.py         codebook 加载、时间和审计
src/engine/candidates.py       候选召回
src/engine/features.py         匹配特征
src/engine/ranker.py           候选排序
src/engine/policy.py           AUTO/HUMAN/MULTIPLE/NONE 分流
src/engine/llm_judge.py        DeepSeek 安全仲裁
src/engine/pipeline.py         批处理、导出和安全校验
src/engine/qt_app.py           PySide6 人工审核界面
src/engine/rank_adjuster.py    post_rank 等级归一化与调整
tests/                         回归和单元测试
resources/                     城市、后缀和冲突词表
skills/deepseek_judge/         DeepSeek 行为约束
```

运行生成的 `outputs/`、`data/`、`cache/`、`reports/`、`*.xlsx` 和模型二进制文件默认被 `.gitignore` 排除，避免将研究数据、缓存、API 相关信息或大型中间结果提交到 GitHub。

## 当前验证

本地已完成：

```text
22 个自动化测试通过
Python 源码编译检查通过
pip check 通过
```

建议每次修改 codebook 或核心代码后重新运行测试，并先用小样本数据检查审核队列，再运行完整数据集。
