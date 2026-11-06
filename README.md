# 政府部门名称编码器

这是一个面向政府组织史数据的部门名称匹配与审核引擎。它把输入数据中的脏部门名称 w_name_o 与结构化部门 codebook 进行匹配，并在满足安全条件时，从 codebook 回填三个正式字段：

- w_name：codebook 中的正式部门名称；
- wrank_t：部门级别，值来自 codebook 的 wrank；
- s1_sector：部门类别，值来自 codebook 的 s1_sector。

核心原则是：模型只能提供匹配证据，不能生成正式字段；正式字段必须通过唯一 entity_id 回查 codebook 得到。无法安全判断时，程序转为人工复核，而不是强行编码。

## 主要能力

1. 读取并审计 Excel/CSV codebook。
2. 进行 NFKC、括号内容、地名前缀、标点、后缀和修饰词处理。
3. 通过精确、别名、核心词、RapidFuzz 和字符 TF-IDF 召回候选。
4. 检查任职时间段与 codebook 有效时间段的交集。
5. 检测共享别名、别名/正式名称冲突、多实体、城市缺失、地区不一致和非法时间。
6. 使用 LightGBM LambdaRank 加分类器排序；没有 LightGBM 时退回标准化逻辑回归。
7. 进行概率校准，并按保守政策分流为 AUTO、HUMAN、MULTIPLE 或 NONE。
8. 对 LLM 仲裁提供候选白名单、Pydantic JSON 校验、候选顺序双调用和预算控制。默认使用不联网的 FakeJudge。
9. 导出 reviewed_content.xlsx 和 modification_log.xlsx。
10. 提供 Streamlit 本地人工复核界面和反馈回流功能。

## 安装

当前开发环境使用 Python 3.13；如果部署环境有 Python 3.11，也可以使用 Python 3.11 创建虚拟环境。

~~~bash
cd /path/to/name-matcher
python3.13 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
python -m pip check
python -m pytest -q
~~~

完整功能需要 pandas、numpy、openpyxl、rapidfuzz、scikit-learn、joblib、LightGBM、jieba、PyYAML、Pydantic、pyarrow、Typer、pytest 和 Streamlit。

## 快速开始

### 审计 codebook

~~~bash
python cli.py audit \
  --codebook /path/to/codebook.xlsx \
  --reports-dir reports/codebook_audit
~~~

输出：

- codebook_audit.md：总体审计结果；
- period_conflicts.csv：时间区间重叠、空时间与非空时间混用；
- notes_need_review.csv：notes 中含年份、撤销、合并、更名、挂牌等关键词的行。

### 训练冷启动模型

~~~bash
python cli.py train \
  --codebook /path/to/codebook.xlsx \
  --output models/synthetic_ranker.joblib
~~~

训练资产记录模型版本、特征哈希、训练规模、模型类型、基线指标和校准方式。合成数据只用于冷启动，不应替代人工标注的 calibration 集和冻结测试集。

### 批处理

~~~bash
python cli.py run \
  --codebook /path/to/codebook.xlsx \
  --input /path/to/input.xlsx \
  --output-dir outputs \
  --config config.yaml \
  --model models/synthetic_ranker.joblib
~~~

每次运行建立 outputs/run_<run_id>/，包含：

~~~text
reviewed_content.xlsx
modification_log.xlsx
eval.md
debug/queries.jsonl
run_meta.json
~~~

### 导出人工复核表

~~~bash
python cli.py review-export \
  --input /path/to/input.xlsx \
  --debug outputs/run_<run_id>/debug/queries.jsonl \
  --output outputs/run_<run_id>/review_queue.xlsx
~~~

复核表包含系统决策、候选 entity_id、top-3 候选、LLM 理由以及 human_decision、human_entity_id、human_note 等人工字段。人工只能选择 codebook 中的 entity_id，不能自由书写正式字段。

### 启动 Streamlit

~~~bash
python cli.py review-ui
~~~

界面支持上传 Excel/CSV、运行匹配、分页查看候选、选择 entity_id、confirm/pick/none/hard 以及下载两个主交付文件。

### 回流反馈

~~~bash
python cli.py ingest-feedback \
  --input outputs/run_<run_id>/review_queue.xlsx \
  --output data/aliases_learned.csv \
  --operator analyst_1 \
  --labels data/labels.csv \
  --cases data/cases.csv
~~~

新别名默认进入 pending_second_review，不会立即改变未来的 exact/alias 规则。只有完成第二审核并标记 approved 后，才适合进入正式生效池。

## 输入数据契约

常用输入字段：

~~~text
record_id, year, styr, stmo, fiyr, fimo,
cityname, distname, dtype, w_name_o, context
~~~

- year 只用于上下文和输出，不替代任职时间段判断；
- styr/stmo/fiyr/fimo 共同构成任职时间段；
- fiyr 为空但 fimo 非空会标记为非法；
- cityname 为空时，程序尝试从 w_name_o 前缀识别城市；
- 没有 record_id 时，程序根据完整源行和重复出现次序生成内部 ID，生产数据建议提供稳定 record_id；
- context 只作为辅助证据，不用于生成正式字段。

广州组织史数据中存在 dtype_t 这样的替代字段名，当前 config.yaml 已设置 fallback。

## Codebook 契约

codebook 必须包含：

~~~text
cityname, distname, dtype, drank, s1_sector, wrank,
w_name, alias, notes,
year_valid_from, month_valid_from,
year_valid_to, month_valid_to
~~~

alias 可以使用 ；、;、 、逗号、| 或换行分隔多个别名。

时间语义：

- 四个有效期字段全为空：该行适用于所有时间；
- 起始年份有值、起始月份为空：按当年 1 月处理；
- 结束年份有值、结束月份为空：按当年 12 月处理；
- 起始年份为空视为负无穷，结束年份为空视为正无穷；
- notes 的自然语言不会在运行时解析，只进入离线审计报告。

每行生成：

- entity_id = sha256(cityname + norm_w_name + start_year + start_month)[:10]；
- dept_key = sha256(cityname + norm_w_name)[:10]。

entity_id 是最终回填字段的唯一查表键，dept_key 用于把同一部门的不同时间版本归组。

## 配置文件

config.yaml 是运行行为的主要入口：

~~~yaml
audited_fields:
  - w_name
  - wrank_t
  - s1_sector

field_mapping:
  w_name: w_name
  wrank_t: wrank
  s1_sector: s1_sector
~~~

重要配置：

- matching.candidate_k：候选数量；
- matching.month_tolerance、year_tolerance：时间容差；
- matching.distname_hard_filter：是否对地区不一致硬过滤；
- policy.tau_auto、margin_auto、tau_nil：自动、LLM 和人工阈值；
- models.llm_model、models.llm_budget：LLM 仲裁配置；
- paths.resource_dir、paths.reports_dir：词表和报告目录。

## 判定流程

~~~text
输入记录
  ↓
规范化、城市识别、多实体识别
  ↓
精确、别名、核心词、RapidFuzz、TF-IDF 候选召回
  ↓
时间交集、城市、地区、歧义和冲突标记
  ↓
特征计算与 LightGBM/逻辑回归排序
  ↓
概率校准与 policy 分流
  ├─ AUTO   → 通过 entity_id 回查 codebook 并回填
  ├─ HUMAN  → 等待人工复核
  ├─ LLM    → 候选白名单仲裁，失败则 HUMAN
  └─ MULTIPLE/NONE → 不强行生成正式字段
~~~

共享别名、时间未知、时间冲突、超出有效期、非法输入时间、地区不一致、城市无法识别、多实体输入和 codebook 时间非法时，不会自动通过。

## 输出安全规则

1. 原始输入文件不会被覆盖。
2. reviewed_content.xlsx 保留原列、列顺序、行顺序和非目标单元格。
3. 只有 config.yaml 中 audited_fields 可以被修改。
4. 正式字段只能通过 entity_id 从 codebook 回查。
5. 每个目标字段记录 before/after、方法、概率、margin、模型版本和 codebook 版本。
6. 导出后逐单元格检查非目标列，发现任何变化就拒绝导出。

## 代码结构

~~~text
cli.py                    命令行入口
config.yaml               运行配置
src/engine/codebook.py    codebook 加载、ID 和审计
src/engine/normalize.py   名称规范化
src/engine/index.py       精确、别名、TF-IDF 索引
src/engine/candidates.py  候选生成和时间逻辑
src/engine/features.py    字符、分词、时间和组内特征
src/engine/ranker.py      LightGBM/逻辑回归排序
src/engine/calibration.py 概率校准和阈值
src/engine/policy.py      自动/LLM/人工分流
src/engine/llm_judge.py   LLM 接口、校验、缓存和预算
src/engine/pipeline.py    批处理和安全导出
src/engine/feedback.py    复核表和反馈回流
src/engine/review_app.py  Streamlit 界面
src/engine/evaluate.py    评估和安全性报告
tests/                    合成 fixture 和回归测试
resources/                城市、后缀、冲突词表
~~~

## 测试与审计

~~~bash
python -m pytest -q
python cli.py audit --codebook /path/to/codebook.xlsx --reports-dir reports/audit
python cli.py eval \
  --debug outputs/run_<run_id>/debug/queries.jsonl \
  --input /path/to/input.xlsx \
  --reviewed outputs/run_<run_id>/reviewed_content.xlsx \
  --log outputs/run_<run_id>/modification_log.xlsx
~~~

测试覆盖 codebook 审计、稳定 ID、共享别名、城市缺失、城市前缀、多实体、核心词召回、非法时间、LLM schema、校准器、非目标列安全校验和人工反馈二审状态。

## 已知限制

- 默认 LLM 是本地 FakeJudge，不访问外部 API。生产环境需要注入真实 judge_fn。
- 当前模型是合成数据冷启动模型，概率校准应在人工标注的独立 calibration 集上重新训练。
- 当前版本提供批处理和本地 Streamlit 复核；大规模断点续跑、主动学习选样和新旧模型冻结测试替换仍需进一步工程化。
- 代码仓库不应提交真实输入数据、个人信息、.venv、输出工作簿和模型缓存；这些文件由 .gitignore 排除。

