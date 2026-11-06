# M0 assumptions

1. The repository's original `codebook.xlsx` is a legacy four-column table
   (`dept_id`, `level`, `standard_name`, `aliases`). The current reference
   codebook follows the 13-column contract and is supplied through the CLI.
2. `entity_id` is the first 10 hexadecimal characters of SHA-256 over
   `cityname|normalized_w_name|year_valid_from|month_valid_from`, with each
   empty time component represented by `NULL`. The same normalized name and
   start boundary therefore produce a stable ID independent of row order.
3. `dept_key` is the first 10 hexadecimal characters of SHA-256 over
   `cityname|normalized_w_name` and groups time-versioned rows for audits.
4. A time interval with all four codebook time fields empty applies to all
   dates. Missing start/end years are treated as negative/positive infinity;
   missing start/end months become January/December when the corresponding
   year is present.
5. A duplicate normalized `cityname + w_name` is reported even when the rows
   are non-overlapping time versions. This is an audit signal, not an error by
   itself.
6. Alias splitting uses semicolon, Chinese semicolon, comma, Chinese comma,
   ideographic comma, vertical bar, or line breaks. Alias ambiguity is
   reported and no alias is deleted.
7. `notes` are scanned only for audit keywords. M0 does not interpret notes to
   infer dates or alter runtime matching behavior.
8. The M0 command is intentionally limited to codebook loading, stable IDs,
   and audit reports. Matching, normalization, candidates, models, UI, and
   export are later milestones.
9. The new reference codebook contains source cells such as `广州市 | 深圳市`
   and 44 rows with missing department fields. M0 preserves multi-city cells
   as source values and reports them; it also reports missing required fields.
   Expanding one source row into several city-specific entities is deferred
   until the candidate-generation milestone because it changes row identity
   semantics.
10. The project uses an isolated `.venv` created with the available Python
    3.13 interpreter (the machine did not expose Python 3.11). The declared
    `pydantic`, `pyyaml`, `pytest`, `lightgbm`, `jieba`, `streamlit`, and
    `pyarrow` dependencies are installed there; a small schema fallback keeps
    imports readable in a minimal system Python.
11. The supplied Guangzhou input workbook maps `dtype` to `dtype_t`, has no
    `record_id` or `context`, and already contains the three audited target
    columns. The pipeline derives a deterministic row identifier from the
    complete source row plus its duplicate occurrence and builds a compact
    context from `p_name` and `post1` when present. Original columns remain
    unchanged in `reviewed_content.xlsx`.
12. City and district codebook cells containing `|` are treated as a set of
    applicable values during retrieval while their source text remains intact.
    This avoids rewriting the source codebook during matching.
13. A FakeJudge remains the default LLM provider: no external API is called by
    the local pipeline. With the installed dependencies, LightGBM LambdaRank,
    jieba token features, Pydantic boundary validation, PyYAML configuration,
    and Streamlit review are available. Real LLM integration still requires a
    caller-supplied `judge_fn`.
