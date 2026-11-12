# 变动审核与 post_rank 审核说明

本文档对应当前仓库中的 PySide6 版本程序，记录两个后续阶段的**实际代码流程**、输入范围、输出结果和复现限制。

## 1. 总体流程

```text
输入 Excel + codebook
  -> 人工名称复核
  -> 保存 *_人工复核完成.xlsx
  -> 变动审核（只读建议）
  -> 人工在 Excel 中处理变动问题并另存
  -> 上传处理后的 Excel
  -> 生成 *_00待处理完成.xlsx
  -> post_rank 审核（只读建议）
  -> 人工决定是否写回 post_rank
```

变动审核和 post_rank 审核都不会自动把建议写回原始输入文件。上传文件先复制为阶段文件。

## 2. 代码入口

| 功能 | 实现 |
|---|---|
| 变动审核入口 | `src/engine/qt_app.py::_load_change_audit` |
| 逐行变动检查 | `src/engine/qt_app.py::_build_change_preview` |
| 人工复核保存 | `src/engine/qt_app.py::save_pending_batch` |
| post_rank 上传 | `src/engine/qt_app.py::_load_post_rank_review` |
| wrank20 预处理 | `src/engine/qt_app.py::_prepare_wrank20_stage` |
| post_rank 审计 | `src/engine/qt_app.py::_audit_post_rank_file` |
| 级别调整 | `src/engine/rank_adjuster.py::adjust_post_rank` |
| codebook 读取 | `src/engine/codebook.py::Codebook` |

变动审核卡片现在调用 `src/engine/department_expansion.py` 中的
`process_codebook_data`，该模块来自独立的“部门级别变动填充与扩行”脚本，
并由界面上的“导出扩行结果”按钮输出四个工作表：`处理结果`、`问题清单`、
`拆分日志`、`卡片摘要`。界面仍然只在用户明确导出时写入新文件。

## 3. 变动审核

### 3.1 输入字段

主要读取：

```text
year, cityname, distname, w_name, w_name_o,
wrank20, wrank_t, post_rank, p_name,
styr, stmo, fiyr, fimo, record_id
```

通常输入是人工复核后的 `*_人工复核完成.xlsx`，也可用“加载人工复核完成文件”选择。

### 3.2 当前审核范围

当前版本根据输入 codebook 的 `cityname` 和 `distname` 自动筛选适用输入行。只有同时满足某条 codebook 记录的城市/区县范围的输入行才进入变动审核；codebook 中为空的范围字段视为通配。

如果 codebook 只包含广州城市级记录，实际效果仍然是只审 `cityname=广州市 且 distname=广州市`；换成省级或多城市 codebook 后，范围会随 codebook 改变。

### 3.3 codebook 记录和索引

程序取 `cb.valid_rows()` 中 `not row.all_time_empty` 的记录。也就是 `year_valid_from`、`month_valid_from`、`year_valid_to`、`month_valid_to` 至少一项非空的记录。

每条记录收集标准名称和所有 alias，建立：

```text
(normalize(cityname), normalize(distname), normalize(w_name 或 alias))
```

输入优先使用 `w_name`，为空才使用 `w_name_o`。城市、区县、名称必须同时命中。

界面按 `(cityname, distname, w_name)` 汇总，展示 s1_sector、wrank、有效期和 notes，并比较：

- s1_sector 是否有多个值；
- wrank 是否有多个值；
- notes 是否提示更名；
- 有效期段数。

改名前后的不同标准名称会分别出现，程序不会自动合并历史实体。

### 3.4 无有效期变化

如果命中的 codebook 记录全部是全时段记录，程序只比较：

```text
wrank20 与 wrank_t
```

两者都非空且不一致时生成：

```text
issue_type = 级别没变但是wrank不统一
```

只生成报告，不自动修改任一字段。

### 3.5 有效期变化

如果存在非全时段记录：

1. 无 `year`：生成 `w_name有效期无法判断`；
2. 有 `year`：用 `stmo` 定位，缺月份默认 6 月；
3. 转换为月序号；
4. 检查 `start_month <= 输入年月 <= end_month`。

无命中时生成 `w_name超出有效期`，并列出 codebook 的有效期。多个命中且 `(w_name,s1_sector,wrank)` 签名不同，生成 `多种可能`。唯一命中但输入字段不同，生成 `变动建议`，并提供：

```text
proposed_w_name
proposed_s1_sector
proposed_wrank_t
action = 人工审阅后决定是否修改
```

以上建议不会自动写回。

### 3.6 跨月份示例

程序按 `p_name、cityname、distname、w_name、w_name_o` 分组，仅生成预览示例。

如果任职区间跨过明确月份节点：

```text
styr/stmo <= 节点 <= fiyr/fimo
```

预览生成“变化前”和“变化后”复制行：

- 变化前行 `fiyr/fimo = 节点年月`；
- 变化后行 `styr/stmo = 节点年月`。

例如节点为 1997-03、任职为 1995-03 至 1998-09：

```text
变化前：1995-03 至 1997-03
变化后：1997-03 至 1998-09
```

只有年份边界时按年份填充，不复制年度行。当前实现只做演示，不真正拆分原 Excel。

### 3.7 变动审核输出

可能生成：

```text
temporal_review_queue.xlsx
review_queue.xlsx
```

去重字段按实际存在情况使用：`record_id, source_index, issue_type, w_name, year`。人工保存阶段还可能生成 `post_rank_adjustment_log.xlsx`，它仍是建议日志。

### 3.8 有效期变化后的 `wrank20` 补填

这是变动扩充模块的最后一步，只作用于存在有效期记录的部门。程序寻找覆盖目标年份（当前默认 2020 年）的 codebook 段，并把该段的 `wrank` 写入输出结果的 `wrank20`。

例如住房公积金管理中心在 2009 年 2 月前为副厅级、2009 年 2 月后为处级；第二段覆盖 2020 年，因此扩充结果中的 `wrank20` 为处级。无有效期变化的部门不会在此步骤改动 `wrank20`。

该操作只写入导出的扩充结果或 post-review 阶段副本，不覆盖原始输入文件。

## 4. post_rank 审核

### 4.1 输入

上传的应是已经人工复核、并已人工处理变动审核的 Excel。最低字段：

```text
cityname, distname, w_name, wrank20, post_rank
```

完整审计建议有：

```text
year, wrank_t, s1_sector, w_name_o, record_id,
p_name, styr, stmo, fiyr, fimo
```

当前界面同样按 codebook 的 cityname/distname 自动确定显示范围；空范围字段视为通配。

### 4.2 wrank20 阶段

`_prepare_wrank20_stage` 复制上传文件为：

```text
原文件名_00待处理完成.xlsx
```

原上传文件不覆盖。

codebook 中 `wrank` 规范化后为 `00待出`、`00待处理` 或包含 `00待出` 的记录会触发处理。按标准名称、alias、cityname、distname 匹配，命中行设：

```text
wrank20 = 00待处理
```

已是该值的行不计入变更数量。

### 4.3 钢铁厂兼容

当标准名称为 `钢铁企业集团`，或 alias 含 `广州钢铁厂`，额外兼容：

```text
钢铁厂
广州钢铁集团
广州钢铁集团有限公司
```

这是特定历史改名规则，不是通用模糊匹配。

### 4.4 级别数字化

内部从低到高：

```text
副股级、股级、副科级、正科级、副处级、
正处级、副厅级、正厅级、副部级、正部级
```

别名归一化：

| 输入 | 内部 |
|---|---|
| 科级 | 正科级 |
| 处级 | 正处级 |
| 厅级、地厅级 | 正厅级 |
| 部级、省部级 | 正部级 |

显示时正科级、正处级、正厅级、正部级优先显示为科级、处级、地厅级、省部级。

### 4.5 调整公式

```text
delta = score(new_wrank) - score(old_wrank)
adjusted_post_score = score(post_rank) + delta
```

例如部门级别由副厅级变为正处级，`delta=-1`；原职位为副厅级，则建议职位为正处级。

无法识别部门级别、无法识别原 post_rank 或调整后越界时返回 `status=unresolved`，保持原值并要求人工判断。

### 4.6 输出

逐行日志字段：

```text
source_row, record_id, p_name, w_name,
old_wrank_t, new_wrank_t,
post_rank_original, post_rank_adjusted,
delta, status, reason
```

界面按 `当前 w_name + 原始 w_name_o + 条数` 汇总。汇总表不是逐行编辑器。

## 5. 当前实现限制

### 5.1 codebook 范围决定审核范围

两个界面都根据当前 codebook 的 `cityname/distname` 自动筛选。空范围字段视为通配；如果需要严格限制到指定城市，应在 codebook 中明确填写范围。

### 5.2 显示范围和内部审计范围要分别核对

`_load_post_rank_review` 先过滤显示用 DataFrame，`_audit_post_rank_file` 调用 `_apply_temporal_postprocess` 时也应用同一 codebook scope。扩展范围时必须继续保证界面显示、逐行审计和写入校验共享该规则。

### 5.3 不自动拆行、不自动写回

变动示例不会复制原 Excel 行；post_rank 建议不会自动写入原文件。人工确认后应由独立写回步骤完成。

### 5.4 空月份是近似定位

输入缺 `stmo` 时按 6 月定位；示例中开始月份缺失通常按 1 月，结束月份缺失通常按 12 月。该规则只用于审阅定位。

## 6. 交给另一个 agent 的复现步骤

### 6.1 变动审核

1. 读取人工复核阶段 Excel；
2. 按目标 city/dist scope 过滤；
3. 加载 codebook；
4. 建立标准名称和 alias 索引；
5. 判断每行是否有有效期记录；
6. 检查名称有效期；
7. 检查无变化部门的 wrank20 与 wrank_t；
8. 生成唯一记录的变动建议；
9. 生成跨月份人员示例；
10. 输出逐行队列和汇总。

### 6.2 post_rank 审核

1. 复制输入为 `*_00待处理完成.xlsx`；
2. 读取 codebook 的 `00待出/00待处理` 记录；
3. 按名称、alias、city、dist 匹配；
4. 更新阶段副本的 `wrank20`；
5. 用内部级别顺序计算 delta；
6. 调整 post_rank；
7. 标记 unresolved；
8. 输出逐行日志；
9. 按当前名称和原始名称汇总；
10. 不覆盖原始上传文件。

### 6.3 必须验证

- 原始输入文件未覆盖；
- 阶段文件行数和列顺序保持；
- wrank20 只在命中规则的行变化；
- 变动审核没有自动改 w_name、s1_sector、wrank_t；
- post_rank 没有自动改原上传文件；
- unresolved 全部进入人工清单；
- scope 同时作用于界面显示和逐行审计。

## 7. 省级 codebook 使用建议

省级 codebook 不需要修改程序中的广州常量；应在 codebook 中明确填写省级机构和各地级市/区县的 `cityname/distname` 范围。空范围表示通配，必须谨慎使用。

建议每次更换 codebook 后先运行小样本，核对范围命中数量，再运行完整数据。

