# 政府部门名称编码与组织史数据审核程序

本项目用于把组织史原始 Excel 中的部门名称和部门级别，与研究者维护的 `codebook.xlsx` 对照，并以可复核、可追溯的方式完成部门编码、历史有效期处理和职位级别核查。程序把自动匹配、人工决定、有效期扩充和最终人工修改分成连续阶段。标准部门名称、类别和部门级别以 codebook 为准；任何不确定或冲突的结果都会保留为审核事项。

本说明按实际操作顺序解释每一步：输入什么、程序如何判断、产生什么文件、哪些地方需要人工决定。配套的高清图示见 [`docs/workflow.svg`](docs/workflow.svg)。

## 全流程一览

```text
原始 Excel + codebook
  → 首次配置与匹配
  → 第一遍人工审核 w_name / s1_sector / wrank_t
  → 保存“人工复核完成”阶段文件
  → 重新上传阶段文件并按 codebook 地域范围核对
  → 判断是否存在有效期变化
       ├─ 没有：wrank20 = wrank_t，生成阶段文件
       └─ 有：按有效期填充、必要时扩行、按 2020 有效段填 wrank20，生成阶段文件
  → 上传当前阶段文件，进行 post_rank 预审
  → 人工修改 post_rank 并保存阶段文件
  → 最终一致性检查与最终成稿
```

每一个“阶段文件”都作为下一步的输入，便于保留过程版本和定位差异。原始数据、阶段文件和最终成稿应分别保存；不要把中间结果直接覆盖原始文件。

## 1. 输入和首次运行配置

开始运行前，需要准备原始组织史 Excel 和适用于本批资料的 codebook。开始界面要求选择原始数据路径、codebook 路径，并填写年鉴起点年月和终点年月。年鉴时间用于判断当前数据的观察范围，在有效期扩充时临时补齐缺失的任职起止年月。程序只在内部计算时补值，原始字段缺失状态会在输出前恢复；由有效期节点产生的拆分边界则会保留。若不使用 DeepSeek，可选择本地规则与排序流程；若使用 API，应在界面输入 Key，不要把 Key 写入代码或提交到仓库。

原始数据常用字段为：

```text
record_id, year, cityname, distname,
w_name_o, w_name, s1_sector, wrank_t, wrank20,
p_name, post1, post_rank, styr, stmo, fiyr, fimo
```

其中 `w_name_o` 是原始名称，`w_name` 是标准化部门名称；`s1_sector` 是部门类别；`wrank_t` 是该条记录或有效期段对应的部门级别；`wrank20` 表示 2020 年部门级别；`post1` 是职位名称；`post_rank` 是数据中记录的个人职位级别。稳定的 `record_id` 有助于逐行追溯；没有该字段时，程序依赖源表行号等内部标识定位。

## 2. 读取 codebook、索引和第一遍人工审核

程序读取原始数据和 codebook 后，按标准名称、alias、地域范围、有效期、部门级别和类别组织查找索引。`w_name` 是唯一的标准名称来源；真实资料中出现的其他叫法只能作为 alias。alias 用于召回候选，不能因此把别名当成新的标准名称。地域字段用于限制候选适用范围；codebook 的 `cityname` 或 `distname` 若为空，则该字段不限制范围。多值范围按 codebook 中的分隔格式拆分后逐项匹配。

系统随后根据原始 `w_name_o` 和现有编码生成候选卡片。相同 `w_name_o` 的记录会合并成名称组，以减少重复审核；卡片中仍应展示组内记录的地区、年份、当前编码等上下文，人工需确认这些记录确实适用同一标准编码。研究者逐项确认 `w_name`、`s1_sector` 和 `wrank_t`，可以选择系统候选，也可以确认输入已有值或按界面规则处理无法匹配的记录。选择先进入本批暂存区，审核者可以回看、编辑或删除暂存项；点击“保存本批复核结果”后，程序才写出人工复核阶段文件，并按当前工作流保留审核状态与日志。

第一阶段文件是后续有效期处理的基线。它记录第一遍人工审核后确定的名称、类别和级别，后续步骤应从这个文件继续，而不是重新上传未审核的原始文件。

## 3. 上传人工复核文件并确定审核范围

将第一阶段“人工复核完成”文件加载到变动审核模块后，程序重新读取当前 codebook，并逐行检查数据记录的 `cityname`、`distname` 是否落在 codebook 表示的范围内。只有地域范围匹配的记录进入后续审核；该范围由本次使用的 codebook 决定，不固定为广州或某个预设城市。codebook 中多值地域会逐项匹配，空白地域字段作为该维度的通配条件。

随后程序根据数据中的 `w_name`（必要时结合原始名称 alias）查找适用的 codebook 记录，并判断该部门是否具有有效期信息。有效期字段为 `year_valid_from/to` 和 `month_valid_from/to`。四项均空表示没有时间分段；存在时间边界表示部门至少有一个历史适用段。若同一名称和地域下有多条有效期记录，程序会结合数据年份、任职起止时间及 codebook 边界定位适用段。缺少月份时，年鉴起止时间用于临时补全任职边界；codebook 起始月份缺失按年初解释，结束月份缺失按年末解释。若数据缺少关键年份或有效期段相互冲突，程序应将该记录列为问题供人工判断，不应把不确定结果当作已确认编码。

## 4. 无有效期变化分流：同步 2020 年级别

如果某部门在当前地域范围内没有有效期变化，程序按既定规则把 `wrank20` 设为本条人工审核后的 `wrank_t`，并生成无变化部门阶段文件。`w_name`、`s1_sector`、`wrank_t` 和 `wrank20` 的变动应能通过阶段前后文件及审核日志核对；源文件中的其他研究字段不应因这一步改变。这个分流让没有历史级别分段的部门，其 2020 年部门级别与当前 codebook 级别保持一致。

如 codebook 级别标为 `00待处理` 等待定值，不能把它当作通常等级参加职位级别运算。项目还保留了将 codebook 中待处理部门对应数据行的 `wrank20` 统一标记为 `00待处理` 的阶段处理规则；该处理会生成独立副本，并显示处理行数，不应覆盖上传文件。

## 5. 有效期变化分流：按时间段填充并处理跨节点任职

有有效期记录的部门进入部门扩充模块。程序以年鉴起止年月为观察边界，临时填补缺少的 `styr/stmo/fiyr/fimo`，再用 codebook 的有效期段为每个部门确定当时适用的 `w_name`、`s1_sector` 和 `wrank_t`。如果任职区间完全落在单一有效期段内，使用该段编码；若任职区间跨越明确的月份变化节点，则将跨节点的年度记录拆成变化前后两段。变化节点所在年份的边界采用节点年月：变化前段结束于该年月，变化后段开始于同一年月，因为资料只精确到月份，无法知道该月具体哪一天发生变化。扩行结果保留源记录关联信息，并在拆分日志中记录源行、分段起止时间和对应编码，便于追查每一行的来历。

临时补出的任职日期用于计算后会恢复为原始空值；真实由节点划分出来的边界仍保留，以表达拆分结果。遇到没有明确月份、有效期重叠、任职年月不完整或无法定位有效段的情况，程序应生成问题记录供人工检查。对这些情形，不能仅凭部门名称自动推定历史变化。

当扩充结束后，程序按 codebook 中覆盖 2020 年的有效期段填充 `wrank20`。例如住房公积金管理中心 2009 年 2 月前为副厅级、从 2009 年 2 月起为处级，2020 年落在后段，因此扩充结果的 `wrank20` 为处级。没有有效期变化的部门不通过这个规则选段；其 `wrank20` 已在上一分流按 `wrank_t` 同步。有效期分流结果输出为独立的部门扩充阶段文件，包含处理结果、问题清单和拆分日志等工作表。人工审核者应先查看这些结果，再把该阶段文件上传到下一步。

## 6. post_rank 预审：从职位名称和部门级别推导理论值

上传当前阶段文件后，post_rank 预审只在 codebook 确定的地域范围内运行，并读取当前最终化的 `wrank_t`、职位名称 `post1` 和原始 `post_rank`。程序使用 [`src/engine/post_rank_mapping.py`](src/engine/post_rank_mapping.py) 中从 [`resources/4post_rank.R`](resources/4post_rank.R) 转录维护的四组职位词表，将能识别的职位归入四类：部门党政领导、部门副领导、二级局及类似机构党政领导、二级局及类似机构党政副手。当前映射约定这四类依次对应 `rank1`、`rank2`、`rank3`、`rank4`。

程序将 `wrank_t` 规范到统一的级别序列，再按职位组次序向下推导：`rank1` 等于部门级别，`rank2` 比部门级别低一级，`rank3` 低两级，`rank4` 低三级。例如 `wrank_t=副厅级` 时，四档理论级别依次为副厅级、处级、副处级、科级。历史级别写法会归一后比较，例如“处级”按正处级比较，“地厅级”按正厅级比较；输出显示名称沿用数据所需写法。若推导结果超出当前级别序列下限或 `wrank_t` 无法识别，则结果进入待人工处理状态。

推导出的理论级别与原 `post_rank` 比较。相同的记录标记为一致；不同的记录作为 `post_rank 不一致` 问题展示；职位名称不在词表中的记录列为“职位名称无法识别”；部门级别无法规范化的记录列为“部门级别无法判断”；缺少原 `post_rank` 的记录单独标出。post_rank 预审输出的是核查结果和建议，不自动更改原始 `post_rank`。研究者在 Excel 中根据原始资料人工修改，并把修改后的文件作为下一阶段输入。词表是明示规则，因此职位名称的匹配可追查；新增、合并或修改职位类别时，应同步更新词表并保留版本记录。

## 7. 人工修改与最终一致性检查

研究者在 post_rank 阶段文件中人工处理不一致与未知职位，保存为新的 Excel，再加载到最终检查步骤。最终检查逐项核对 `w_name`、`s1_sector`、`wrank_t`、`wrank20` 和 `post_rank`，确认人工审核与变动扩充后的字段之间没有明显冲突，并保留原始字段、修改字段和问题说明供复查。最终成稿以人工确认的文件为依据；程序输出的“建议”不等于研究者已经批准的修改。

## 8. 运行方法

推荐 Python 3.11–3.13，并在项目目录建立虚拟环境：

```bash
cd /path/to/wname_matcher
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
python run_qt.py
```

启动后依次选择原始 Excel、codebook 和配置文件，填写年鉴起止年月，再运行第一遍匹配与人工复核。保存“人工复核完成”文件后，在变动审核模块加载该文件；检查、导出并保存变动审核阶段结果后，再将其上传到 post_rank 审阅模块。完成 Excel 人工修改后，按最终检查步骤生成成稿。旧版 Streamlit 入口保留用于兼容：

```bash
python cli.py review-ui
```

命令行的 codebook 审计和批量运行入口如下：

```bash
python cli.py audit --codebook /path/to/codebook.xlsx --reports-dir reports/codebook_audit
python cli.py train --codebook /path/to/codebook.xlsx --output models/synthetic_ranker.joblib
python cli.py run --codebook /path/to/codebook.xlsx --input /path/to/input.xlsx --output-dir outputs --config config.yaml --model models/synthetic_ranker.joblib
```

## 9. 输出、日志与追溯

批处理运行目录通常包含：

| 文件 | 内容与用途 |
|---|---|
| `reviewed_content.xlsx` | 输入表副本及程序生成的审核结果字段 |
| `review_queue.xlsx` | 错误输入、字段冲突、多种可能、无法判断等人工队列 |
| `modification_log.xlsx` | 字段级修改前后值及对应来源 |
| `temporal_review_queue.xlsx` | 名称超出有效期或有效期无法判断的记录 |
| `post_rank_adjustment_log.xlsx` | 部门等级变化所对应的职位等级联动建议 |
| `debug/queries.jsonl` | 名称查询、候选与判定证据，供技术排查 |
| `run_meta.json` | 本次运行配置和元数据 |

部门扩充阶段文件另含处理结果、问题清单、拆分日志和汇总。post_rank 预审表逐条显示年份、部门、`post1`、当前 `wrank_t`、推导的职位类别、理论级别、原 `post_rank` 和问题类型。`source_index`、源行号或 `record_id` 应与原数据结合使用，不能把排序后的表格显示行号当作源数据行号。

可追溯性来自四类证据：本批使用的 codebook 记录（标准名称、alias、地域范围、有效期及等级）；源 Excel 的行标识；程序生成的阶段文件与拆分日志；以及人工审核暂存和修改前后日志。分享或归档一次审核时，应一并保存原始输入版本、codebook 版本、程序版本、年鉴范围、阶段文件和最终审核记录。公共 GitHub 仓库不应包含研究数据、API Key、个人缓存或本地临时文件。

## 10. 输入格式与 codebook 约定

codebook 至少应提供以下字段：

```text
cityname, distname, s1_sector, wrank, w_name, alias,
year_valid_from, month_valid_from, year_valid_to, month_valid_to
```

项目其他匹配模块可能还会使用 `dtype`、`drank`、`notes` 等字段。正式运行前应检查 codebook 的列名、标准名称唯一性、alias 冲突、地域范围和有效期边界。`alias` 支持常见分隔符；标准名始终是 `w_name`。有效期四字段为空代表不分历史段；只提供年份边界时按整年解释；提供月份边界时按具体年月解释。若有效期段有重叠或明显断档，先修正或确认 codebook，再处理数据。

## 11. 重要规则与人工责任

- codebook 决定标准部门名称、适用地区、部门类别、部门级别和有效期。
- DeepSeek（若启用）只能帮助判断 codebook 已召回的候选，不能创制正式标准名称；API Key 不应进入代码、配置提交或日志。
- 名称组审核减少重复操作，但只有上下文一致的同名原始记录才能共用同一决定；跨地区、跨年份或跨有效期记录要检查范围。
- 有效期扩充可能增加行数。需通过源行标识和拆分日志确认扩行的前后来源。
- `wrank20` 是 2020 年部门级别字段。无有效期变化时按 `wrank_t` 同步；有变化时取覆盖 2020 年的 codebook 段。
- `4post_rank.R` 词表分类属于可解释规则，不是统计概率。未命中的职位必须人工识别；边界级别无法推导时也应进入人工审核。
- post_rank 预审只提出理论值与差异记录。最终 `post_rank` 的更改由研究者人工确认。
- 出现字段缺失、codebook 冲突、地域不匹配、有效期重叠、等级未知或导出保护校验失败时，应停止使用该批输出并查看具体报告。

## 12. 项目结构

```text
run_qt.py                              PySide6 界面入口
cli.py                                 命令行入口
config.yaml                            程序配置
src/engine/codebook.py                 codebook 读取、校验与适用范围
src/engine/candidates.py               候选召回
src/engine/ranker.py                   候选排序
src/engine/policy.py                   自动/人工/多候选分流
src/engine/llm_judge.py                可选的 DeepSeek 候选裁决
src/engine/pipeline.py                 匹配、导出和内容保护校验
src/engine/qt_app.py                   人工复核、变动审核和 post_rank 页面
src/engine/department_expansion.py     有效期匹配、任职区间拆分和 wrank20 填充
src/engine/post_rank_mapping.py        4post_rank.R 职位词表与理论级别推导
resources/4post_rank.R                 原始职位词表来源
src/engine/rank_adjuster.py            等级规范化及部门等级联动计算
docs/workflow.svg                      高清完整流程图
docs/CHANGE_AUDIT_AND_POST_RANK.md     变动审核与 post_rank 的实现说明
tests/                                 自动化测试
```

## 13. 验证和版本维护

修改核心规则后，在项目虚拟环境中运行：

```bash
python -m pytest -q
python -m py_compile src/engine/*.py
```

涉及职位词表、级别顺序、有效期解释、范围匹配或 Excel 输出的修改，应同时更新本 README、流程图和相应实现说明，并用包含已知正确答案的小样本验证正例、冲突、未知职位和边界年月。版本交付时记录代码提交、codebook 文件版本、输入文件版本和运行参数，确保之后能够复现本次审核。

## 14. 近期流程性能优化

人工复核完成后，变动审核使用人工确认后的阶段 Excel 与当前最新 codebook。初始审核队列会预先保存 `scope_match`、`review_group_key`、`review_group_count`、`codebook_version` 和 `model_version`，重新打开人工审核时直接复用这些结果，避免再次扫描整个 codebook 或重新构造名称组。若人工审核过程中修改了 codebook，应在变动审核前重新选择最新文件；页面顶部会提示这一点。

结果预览页面现在作为流程检查清单使用。选择人工复核阶段文件后，程序先让界面显示文件名和处理状态，再开始变动审核；进度条显示读取文件、范围筛选、识别变动部门、有效期匹配与报告生成等阶段。`_load_preview()` 已移除，不再为了显示无业务作用的预览表而读取和渲染完整结果。

这些优化保持了范围过滤、按名称组合并、记录级追溯、有效期拆分日志和最终安全校验。旧审核队列缺少新增字段时仍走兼容计算；新的审核队列则直接使用预计算字段。变动审核报告仍然只提供建议，不会自动覆盖人工阶段文件。
