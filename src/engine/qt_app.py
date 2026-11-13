"""PySide6 desktop review application with a non-blocking worker thread."""

from __future__ import annotations

import sys
import tempfile
import json
import ast
import shutil
import re
from types import SimpleNamespace
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar,
    QPushButton, QSpinBox, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
    QHeaderView, QAbstractItemView, QStackedWidget, QListWidget, QListWidgetItem,
    QSizePolicy, QDialog, QScrollArea,
)

from .llm_judge import DeepSeekJudge, LLMJudge
from .pipeline import Pipeline
from .codebook import Codebook, normalize_name, split_aliases
from .rank_adjuster import adjust_post_rank
from .post_rank_mapping import audit_post_rank_frame
from .department_expansion import CFG as EXPANSION_DEFAULT_CONFIG, build_card_summary, process_codebook_data, split_multi_value


def _parse_structured(value):
    """Read list/dict columns from JSONL, Excel, or pandas string repr safely."""
    if isinstance(value, (list, dict)):
        return value
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return []
    for parser in (json.loads, ast.literal_eval):
        try:
            parsed = parser(text)
            if isinstance(parsed, (list, dict)):
                return parsed
        except (ValueError, SyntaxError, TypeError, json.JSONDecodeError):
            continue
    return []


class RunWorker(QObject):
    progress = Signal(str, int, int)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, input_path: Path, codebook_path: Path, config_path: Path, model_path: str, provider: str, api_key: str, model: str, base_url: str, timeout: float, budget: int):
        super().__init__()
        self.args = locals().copy()
        self.args.pop("self")

    @Slot()
    def run(self) -> None:
        try:
            a = self.args
            self.progress.emit("初始化匹配引擎", 0, 1)
            temp_dir = Path(tempfile.gettempdir()) / "department_matcher_qt"
            temp_dir.mkdir(parents=True, exist_ok=True)
            if a["provider"] == "DeepSeek API":
                judge = LLMJudge(model=a["model"], judge_fn=DeepSeekJudge(a["api_key"], a["model"], a["base_url"], a["timeout"]), cache_path=temp_dir / "deepseek_llm_cache.jsonl", budget=a["budget"])
            else:
                judge = LLMJudge(model="fake", cache_path=temp_dir / "fake_llm_cache.jsonl", budget=0)
            self.progress.emit("加载 codebook 和模型", 0, 1)
            pipeline = Pipeline(a["codebook_path"], config_path=a["config_path"], model_path=a["model_path"] or None, llm_judge=judge)
            result = pipeline.run(a["input_path"], temp_dir / "runs", progress_callback=lambda s, d, t: self.progress.emit(s, d, t))
            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(f"{type(exc).__name__}: {exc}")


class HumanReviewDialog(QDialog):
    """Card-based adjudication dialog.

    Decisions are deliberately staged in memory.  Excel/codebook/feedback
    writes happen only after the user clicks the batch-save action in the
    parent window, so a slow workbook write cannot interrupt every click.
    """

    def __init__(self, debug: pd.DataFrame, source_path: Path, codebook_path: Path, parent=None, source_positions: list[list[int]] | None = None, existing_decisions: list[dict] | None = None):
        super().__init__(parent)
        self.debug = debug.reset_index(drop=True); self.source_positions = source_positions or [[index] for index in range(len(self.debug))]; self.source_path = source_path; self.codebook_path = codebook_path; self.position = 0
        # Do not open a large source workbook while merely browsing cards.
        # The workbook is opened once at batch commit time instead.
        self.source_book = None
        self.codebook_model = Codebook(self.codebook_path)
        self._added_aliases: dict[str, str] = {}
        self.pending_decisions: list[dict] = list(existing_decisions or [])
        self.setWindowTitle("人工复核 · 候选裁决"); self.resize(1080, 720)
        self.summary = QLabel(); self.summary.setWordWrap(True); self.summary.setObjectName("reviewSummary")
        self.candidates_box = QVBoxLayout(); self.candidate_group = QGroupBox("候选部门（点击候选后确认）"); self.candidate_group.setLayout(self.candidates_box)
        self.note = QLineEdit(); self.note.setPlaceholderText("人工备注（可选）")
        self.standard = QLineEdit(); self.standard.setPlaceholderText("手动标准 w_name（仅用于特殊情况）")
        self.rank = QLineEdit(); self.rank.setPlaceholderText("wrank_t")
        self.sector = QLineEdit(); self.sector.setPlaceholderText("s1_sector")
        fields = QHBoxLayout(); fields.addWidget(self.standard, 3); fields.addWidget(self.rank, 1); fields.addWidget(self.sector, 2)
        self.selected_entity = ""
        self.selected_candidate_values: dict[str, str] = {}
        nav_buttons = QHBoxLayout()
        self.previous_button = QPushButton("← 上一条"); self.previous_button.clicked.connect(self.previous); nav_buttons.addWidget(self.previous_button)
        self.next_button = QPushButton("下一条 →"); self.next_button.clicked.connect(self.next); nav_buttons.addWidget(self.next_button); nav_buttons.addStretch()
        buttons = QHBoxLayout(); self.action_buttons = []
        for text, action, name in [("✓ 确认系统建议（暂存）", "confirm", "primary"), ("保留当前编码（暂存）", "keep_current", ""), ("选择候选并暂存", "pick", ""), ("暂存手动标准", "manual", ""), ("无法匹配（暂存）", "none", ""), ("跳过", "skip", "")]:
            b = QPushButton(text); b.setObjectName(name); b.clicked.connect(lambda checked=False, a=action: self.commit(a)); buttons.addWidget(b); self.action_buttons.append(b)
        self.finish_button = QPushButton("完成暂存并返回"); self.finish_button.clicked.connect(lambda: self.accept()); buttons.addWidget(self.finish_button)
        buttons.addStretch()
        root = QVBoxLayout(self); root.setContentsMargins(24, 22, 24, 22); root.setSpacing(14); root.addWidget(self.summary); root.addWidget(self.candidate_group); root.addWidget(self.note); root.addLayout(fields); root.addLayout(nav_buttons); root.addLayout(buttons)
        self._render()

    def _render(self):
        if self.position >= len(self.debug):
            self.summary.setText(f"本轮审核已暂存完成，共 {len(self.pending_decisions)} 条。你可以点击‘上一条’修改之前的审核，或点击‘完成暂存并返回’查看批次清单。")
            while self.candidates_box.count():
                item = self.candidates_box.takeAt(0); widget = item.widget()
                if widget: widget.deleteLater()
            self.previous_button.setEnabled(bool(self.debug)); self.next_button.setEnabled(False)
            for button in self.action_buttons: button.setEnabled(False)
            self.finish_button.setEnabled(True)
            return
        row = self.debug.iloc[self.position]
        pending = self._pending_for_row(row)
        self.selected_entity = str((pending or {}).get("entity_id") or row.get("entity_id", ""))
        pending_values = (pending or {}).get("values", {})
        self.selected_candidate_values = {
            "w_name": str(pending_values.get("w_name", "") or ""),
            "wrank_t": str(pending_values.get("wrank_t", "") or ""),
            "s1_sector": str(pending_values.get("s1_sector", "") or ""),
        }
        group_count = int(row.get("review_group_count", 1) or 1)
        self.summary.setText(f"第 {self.position + 1} / {len(self.debug)} 个名称组　｜　原始记录 {group_count} 条\n问题类型：{row.get('review_reason', row.get('issue_type','多种可能'))}\n原始输入：{row.get('raw_name','')}\n已有编码：{row.get('existing_w_name','')} ｜ {row.get('existing_wrank_t','')} ｜ {row.get('existing_s1_sector','')}\n系统建议：{self.selected_entity} ｜ {row.get('reason','')}\n确认后将一次性写入该名称组的全部原始行；原始称呼会作为 alias 写入 codebook（若发生冲突则拒绝保存）。")
        while self.candidates_box.count():
            item = self.candidates_box.takeAt(0); widget = item.widget()
            if widget: widget.deleteLater()
        candidates = _parse_structured(row.get("top3_candidates", []))
        # review_queue.xlsx stores candidate dictionaries as a Python repr in
        # some pandas/openpyxl versions.  If that column is absent or empty,
        # reconstruct cards from candidate_ids and the current codebook.
        if not candidates:
            candidate_ids = _parse_structured(row.get("candidate_ids", []))
            candidates = []
            for cid in candidate_ids:
                selected = self._row_for_entity(str(cid), row)
                if selected is not None:
                    candidates.append({"entity_id": selected.entity_id, "w_name": selected.w_name, "wrank": selected.wrank, "s1_sector": selected.s1_sector})
            if not candidates and self.selected_entity:
                selected = self._row_for_entity(self.selected_entity, row)
                if selected is not None:
                    candidates = [{"entity_id": selected.entity_id, "w_name": selected.w_name, "wrank": selected.wrank, "s1_sector": selected.s1_sector}]
        for candidate in candidates[:5] if isinstance(candidates, list) else []:
            if not isinstance(candidate, dict): continue
            cid = str(candidate.get("entity_id", "")); label = f"{candidate.get('w_name','')}   {candidate.get('wrank','')}   {candidate.get('s1_sector','')}\nentity_id: {cid}"
            button = QPushButton(label); button.setObjectName("candidate"); button.setCheckable(True); button.setChecked(cid == self.selected_entity and (not self.selected_candidate_values or all(str(candidate.get(k, "") or "") == self.selected_candidate_values.get(k, "") for k in ("w_name", "wrank", "s1_sector")))); button.clicked.connect(lambda checked=False, value=cid, values=candidate: self._choose(value, values)); self.candidates_box.addWidget(button)
        self.note.setText(str((pending or {}).get("note", "")))
        values = (pending or {}).get("values", {})
        self.standard.setText(str(values.get("w_name", "")))
        self.rank.setText(str(values.get("wrank_t", "")))
        self.sector.setText(str(values.get("s1_sector", "")))
        self.previous_button.setEnabled(self.position > 0)
        self.next_button.setEnabled(self.position < len(self.debug) - 1 or self.position == len(self.debug) - 1)
        for button in self.action_buttons: button.setEnabled(True)
        self.finish_button.setEnabled(True)

    def _pending_for_row(self, row):
        ids = {str(value) for value in _parse_structured(row.get("_group_record_ids", []))}
        if not ids and row.get("record_id", ""):
            ids = {str(row.get("record_id"))}
        for decision in reversed(self.pending_decisions):
            if ids and ids == {str(value) for value in decision.get("record_ids", [])}:
                return decision
        return None

    def previous(self):
        if not len(self.debug): return
        if self.position >= len(self.debug): self.position = len(self.debug) - 1
        elif self.position > 0: self.position -= 1
        self._render()

    def next(self):
        if not len(self.debug): return
        if self.position < len(self.debug): self.position += 1
        self._render()

    def _choose(self, entity_id: str, candidate: dict | None = None):
        self.selected_entity = entity_id
        if candidate:
            self.selected_candidate_values = {
                "w_name": str(candidate.get("w_name", "") or ""),
                "wrank_t": str(candidate.get("wrank", "") or ""),
                "s1_sector": str(candidate.get("s1_sector", "") or ""),
            }
        for i in range(self.candidates_box.count()):
            w = self.candidates_box.itemAt(i).widget()
            if w: w.setChecked(str(w.text()).split("entity_id: ")[-1] == entity_id and (candidate is None or all(str(candidate.get(k, "") or "") in w.text() for k in ("w_name", "wrank", "s1_sector"))))

    def _source_rows(self, record_ids: list[str]) -> list[int]:
        if self.source_book is None:
            return []
        sheet = self.source_book[self.source_book.sheetnames[0]]
        headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
        if "record_id" in headers:
            wanted = set(record_ids); found = []
            for r in range(2, sheet.max_row + 1):
                if str(sheet.cell(r, headers["record_id"]).value or "") in wanted: found.append(r)
            if len(found) == len(record_ids): return found
        return [position + 2 for position in self.source_positions[self.position]]

    def _row_for_entity(self, entity_id: str, context=None):
        """Resolve duplicate entity IDs using the input year/month.

        The stable entity ID intentionally identifies a department across
        validity segments, so a name such as 办公室 can map to several rows.
        The previous implementation required one physical codebook row and
        therefore rejected otherwise valid candidate selections.
        """
        if not entity_id:
            return None
        rows = self.codebook_model.rows_for_entity(entity_id)
        if not rows:
            return None
        if context is not None:
            try:
                year = None if pd.isna(context.get("year")) or str(context.get("year")).strip() == "" else int(float(context.get("year")))
            except (TypeError, ValueError):
                year = None
            try:
                month = None if pd.isna(context.get("stmo")) or str(context.get("stmo")).strip() == "" else int(float(context.get("stmo")))
            except (TypeError, ValueError):
                month = None
            if year is not None:
                key = year * 12 + ((month or 6) - 1)
                in_period = [row for row in rows if row.start_month <= key <= row.end_month]
                if in_period:
                    rows = in_period
        # Duplicate identical rows are common in the codebook; retain the
        # first row after period filtering because their coding values agree.
        return rows[0]

    def _row_for_existing_values(self, row):
        current_name = str(row.get("existing_w_name", "") or "").strip()
        current_rank = str(row.get("existing_wrank_t", "") or "").strip()
        current_sector = str(row.get("existing_s1_sector", "") or "").strip()
        if not current_name or not current_rank or not current_sector:
            raise ValueError("当前输入的 w_name、wrank_t、s1_sector 没有全部填写，无法确认")
        cb = self.codebook_model; city = normalize_name(str(row.get("cityname", "") or ""))
        matches = [item for item in cb.valid_rows() if normalize_name(item.w_name) == normalize_name(current_name) and item.wrank == current_rank and item.s1_sector == current_sector and (not city or any(normalize_name(value) == city for value in cb.split_values(item.cityname)))]
        if len(matches) != 1:
            raise ValueError(f"当前编码在 codebook 中没有唯一对应标准：{current_name} / {current_rank} / {current_sector}")
        return matches[0]

    def _write_alias(self, raw_name: str, entity_id: str):
        selected = self._row_for_entity(entity_id)
        if selected is None: raise ValueError("候选部门不存在，无法写入 alias")
        raw = str(raw_name or "").strip(); raw_norm = normalize_name(raw)
        if not raw_norm or raw_norm == selected.norm_w_name: return
        cb = self.codebook_model
        selected_cities = {normalize_name(x) for x in cb.split_values(selected.cityname)}
        for row in cb.rows:
            if row.entity_id == entity_id: continue
            cities = {normalize_name(x) for x in cb.split_values(row.cityname)}
            if selected_cities & cities and (raw_norm == row.norm_w_name or raw_norm in {normalize_name(a) for a in row.aliases}):
                raise ValueError(f"alias 冲突：‘{raw}’ 已属于 {row.w_name}，本次保存被阻止")
        if raw_norm in self._added_aliases and self._added_aliases[raw_norm] != entity_id:
            raise ValueError(f"alias 冲突：‘{raw}’ 已在本次复核中归属于其他部门")
        if self._added_aliases.get(raw_norm) == entity_id:
            return
        if raw_norm in {normalize_name(a) for a in selected.aliases}: return
        if not self.codebook_path.with_suffix(self.codebook_path.suffix + ".bak").exists(): shutil.copy2(self.codebook_path, self.codebook_path.with_suffix(self.codebook_path.suffix + ".bak"))
        book = load_workbook(self.codebook_path); sheet = book[book.sheetnames[0]]; alias_col = next((c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value == "alias"), None)
        if alias_col is None: raise ValueError("codebook 缺少 alias 列")
        cell = sheet.cell(selected.source_row, alias_col); existing = split_aliases(cell.value); cell.value = "；".join(existing + [raw]); book.save(self.codebook_path); self._added_aliases[raw_norm] = entity_id

    def _write_source(self, row, values: dict[str, str], action: str):
        backup = self.source_path.with_suffix(self.source_path.suffix + ".bak")
        if not backup.exists(): shutil.copy2(self.source_path, backup)
        sheet = self.source_book[self.source_book.sheetnames[0]]; headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
        record_ids = list(row.get("_group_record_ids", []))
        excel_rows = self._source_rows(record_ids)
        for key, value in values.items():
            if key not in headers:
                headers[key] = sheet.max_column + 1; sheet.cell(1, headers[key]).value = key
            for excel_row in excel_rows: sheet.cell(excel_row, headers[key]).value = value
        self.source_book.save(self.source_path)

    def _write_feedback(self, row, action: str, entity_id: str, values: dict[str, str]):
        path = Path(__file__).resolve().parents[2] / "data" / "human_feedback.csv"; path.parent.mkdir(parents=True, exist_ok=True)
        rows = []
        for record_id in row.get("_group_record_ids", [str(row.get("record_id", ""))]):
            rows.append({"record_id": str(record_id), "raw_name": str(row.get("raw_name", "")), "cityname": str(row.get("cityname", "")), "action": action, "entity_id": entity_id, "w_name": values.get("w_name", ""), "wrank_t": values.get("wrank_t", ""), "s1_sector": values.get("s1_sector", ""), "human_note": self.note.text().strip()})
        pd.DataFrame(rows).to_csv(path, mode="a", header=not path.exists(), index=False)

    def _add_manual_codebook_row(self, row, standard: str, rank: str, sector: str):
        """Add a genuinely new standard row, preserving the codebook schema."""
        cb = self.codebook_model; norm = normalize_name(standard)
        if any(item.norm_w_name == norm and normalize_name(str(row.get("cityname", ""))) in {normalize_name(x) for x in cb.split_values(item.cityname)} for item in cb.rows):
            raise ValueError(f"标准名称冲突：‘{standard}’ 已存在于当前城市 codebook")
        if not self.codebook_path.with_suffix(self.codebook_path.suffix + ".bak").exists(): shutil.copy2(self.codebook_path, self.codebook_path.with_suffix(self.codebook_path.suffix + ".bak"))
        book = load_workbook(self.codebook_path); sheet = book[book.sheetnames[0]]; headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
        required = ["cityname", "distname", "dtype", "drank", "s1_sector", "wrank", "w_name", "alias", "notes", "year_valid_from", "month_valid_from", "year_valid_to", "month_valid_to"]
        if any(key not in headers for key in required): raise ValueError("codebook 字段不完整，无法新增标准部门")
        values = {"cityname": str(row.get("cityname", "")), "distname": str(row.get("distname", "")), "dtype": str(row.get("dtype", "") or "待核实"), "drank": str(row.get("drank", "") or "待核实"), "s1_sector": sector, "wrank": rank, "w_name": standard, "alias": str(row.get("raw_name", "")), "notes": "人工复核新增；请后续审计", "year_valid_from": "", "month_valid_from": "", "year_valid_to": "", "month_valid_to": ""}
        target = sheet.max_row + 1
        for key, value in values.items(): sheet.cell(target, headers[key]).value = value
        book.save(self.codebook_path); self.codebook_model = Codebook(self.codebook_path)

    def _stage(self, row, action: str, entity_id: str, values: dict[str, str], manual_new: bool = False):
        record_ids = [str(value) for value in _parse_structured(row.get("_group_record_ids", []))]
        if not record_ids:
            record_ids = [str(row.get("record_id", ""))]
        # debug rows are emitted in original input order, so grouped_positions
        # maps directly to Excel data rows (+1 for the header row).
        source_rows = [int(position) + 2 for position in self.source_positions[self.position]]
        decision = {
            "record_ids": record_ids,
            "source_rows": source_rows,
            "raw_name": str(row.get("raw_name", "") or "").strip(),
            "cityname": str(row.get("cityname", "") or "").strip(),
            "distname": str(row.get("distname", "") or "").strip(),
            "issue_type": str(row.get("issue_type", "") or ""),
            "action": action,
            "entity_id": entity_id,
            "values": values,
            "note": self.note.text().strip(),
            "manual_new": manual_new,
            "dtype": str(row.get("dtype", "") or ""),
            "drank": str(row.get("drank", "") or ""),
        }
        key = tuple(decision["record_ids"])
        self.pending_decisions = [old for old in self.pending_decisions if tuple(old.get("record_ids", [])) != key]
        self.pending_decisions.append(decision)
        owner = self.parent()
        if owner is not None and hasattr(owner, "_merge_pending_decisions"):
            owner._merge_pending_decisions([decision])
            owner._persist_pending_batch()

    def commit(self, action: str):
        if action == "skip":
            self.position += 1; self._render(); return
        row = self.debug.iloc[self.position]
        entity_id = self.selected_entity if action in {"confirm", "pick"} else ""
        try:
            manual_new = False
            if action == "keep_current":
                selected = self._row_for_existing_values(row); entity_id = selected.entity_id
                values = {"w_name": selected.w_name, "wrank_t": selected.wrank, "s1_sector": selected.s1_sector}
            elif action == "manual":
                if not all([self.standard.text().strip(), self.rank.text().strip(), self.sector.text().strip()]):
                    raise ValueError("手动标准必须同时填写 w_name、wrank_t、s1_sector")
                typed_name, typed_rank, typed_sector = self.standard.text().strip(), self.rank.text().strip(), self.sector.text().strip()
                city = normalize_name(str(row.get("cityname", "")))
                existing = [item for item in self.codebook_model.valid_rows() if normalize_name(item.w_name) == normalize_name(typed_name) and (not city or any(normalize_name(value) == city for value in self.codebook_model.split_values(item.cityname)))]
                if len(existing) > 1:
                    raise ValueError(f"标准名称冲突：‘{typed_name}’ 在当前城市存在多个 codebook 版本")
                if len(existing) == 1:
                    selected = existing[0]; entity_id = selected.entity_id
                    values = {"w_name": selected.w_name, "wrank_t": selected.wrank, "s1_sector": selected.s1_sector}
                else:
                    values = {"w_name": typed_name, "wrank_t": typed_rank, "s1_sector": typed_sector}; manual_new = True
            elif action == "none":
                values = {"w_name": "", "wrank_t": "", "s1_sector": ""}
            else:
                values = dict(self.selected_candidate_values) if self.selected_candidate_values.get("w_name") and entity_id else {}
                if not values:
                    selected = self._row_for_entity(entity_id, row)
                    if selected is None: raise ValueError("请选择一个有效候选")
                    values = {"w_name": selected.w_name, "wrank_t": selected.wrank, "s1_sector": selected.s1_sector}
            self._stage(row, action, entity_id, values, manual_new=manual_new)
            self.position += 1; self._render()
        except Exception as exc:
            QMessageBox.warning(self, "暂存被阻止", str(exc))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("部门名称编码工作台")
        self.resize(1180, 760)
        self.worker_thread: QThread | None = None
        self.result: dict | None = None
        self.source_path: Path | None = None
        self.pending_path = Path(__file__).resolve().parents[2] / "data" / "pending_review.json"
        self.pending_decisions: list[dict] = self._load_pending_batch()
        self._build()

    def _load_pending_batch(self) -> list[dict]:
        if not self.pending_path.exists():
            return []
        try:
            value = json.loads(self.pending_path.read_text(encoding="utf-8"))
            return value if isinstance(value, list) else []
        except (OSError, ValueError, TypeError):
            return []

    def _persist_pending_batch(self) -> None:
        self.pending_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.pending_path.with_suffix(".json.tmp")
        temporary.write_text(json.dumps(self.pending_decisions, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        temporary.replace(self.pending_path)

    def _build(self) -> None:
        root = QWidget(); self.setCentralWidget(root); shell = QHBoxLayout(root); shell.setContentsMargins(0, 0, 0, 0); shell.setSpacing(0)
        nav = QWidget(); nav.setObjectName("sidebar"); nav.setFixedWidth(250); nav_layout = QVBoxLayout(nav); nav_layout.setContentsMargins(20, 25, 20, 18); nav_layout.setSpacing(8)
        brand = QLabel("WNAME\n人工审核工作台"); brand.setObjectName("brand"); brand.setTextFormat(Qt.PlainText); brand.setWordWrap(True); brand.setMinimumHeight(78); brand.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed); nav_layout.addWidget(brand)
        nav_hint = QLabel("部门名称标准化与人工裁决"); nav_hint.setObjectName("navhint"); nav_layout.addWidget(nav_hint); nav_layout.addSpacing(22)
        self.nav_list = QListWidget(); self.nav_list.setObjectName("navlist"); self.nav_list.setSpacing(5)
        for text in ["开始运行", "结果预览", "人工复核", "变动审阅", "post_rank 审阅", "使用说明"]:
            item = QListWidgetItem(text); self.nav_list.addItem(item)
        self.nav_list.setCurrentRow(0); nav_layout.addWidget(self.nav_list); nav_layout.addStretch()
        version = QLabel("本地运行 · 数据不上传\nPySide6 Desktop"); version.setObjectName("version"); nav_layout.addWidget(version)
        shell.addWidget(nav)
        content = QWidget(); content_layout = QVBoxLayout(content); content_layout.setContentsMargins(30, 24, 30, 24); content_layout.setSpacing(16); shell.addWidget(content, 1)
        header = QHBoxLayout(); title = QLabel("部门名称编码工作台\n<span class='subtitle'>标准化匹配　·　候选裁决　·　人工复核</span>"); title.setObjectName("title"); title.setTextFormat(Qt.RichText); header.addWidget(title); header.addStretch(); self.header_status = QLabel("● 就绪"); self.header_status.setObjectName("ready"); header.addWidget(self.header_status, 0, Qt.AlignTop); content_layout.addLayout(header)
        pages = QStackedWidget(); content_layout.addWidget(pages, 1); self.pages = pages; self.nav_list.currentRowChanged.connect(pages.setCurrentIndex)
        setup = QWidget(); setup_layout = QVBoxLayout(setup); setup_layout.setContentsMargins(2, 4, 2, 2); setup_layout.setSpacing(14); pages.addWidget(setup)
        intro = QLabel("选择数据文件和标准 codebook，然后开始审核。已有运行结果也可以直接加载。")
        intro.setObjectName("hint"); setup_layout.addWidget(intro)
        files_box = QGroupBox("文件与模型"); files_layout = QFormLayout(files_box); files_layout.setHorizontalSpacing(18); files_layout.setVerticalSpacing(10)
        self.codebook = QLineEdit(str(Path(__file__).resolve().parents[2] / "codebook.xlsx")); files_layout.addRow("标准 codebook", self._path_row(self.codebook, "选择 codebook"))
        self.config = QLineEdit(str(Path(__file__).resolve().parents[2] / "config.yaml")); files_layout.addRow("运行配置", self._path_row(self.config, "选择 config"))
        self.model_path = QLineEdit("models/synthetic_ranker.joblib"); files_layout.addRow("排序模型", self.model_path)
        self.input = QLineEdit(); files_layout.addRow("待审核数据", self._path_row(self.input, "选择输入文件", save=False))
        setup_layout.addWidget(files_box)
        yearbook_box = QGroupBox("年鉴时间范围（用于补充缺失任职年月）")
        yearbook_layout = QFormLayout(yearbook_box); yearbook_layout.setHorizontalSpacing(18); yearbook_layout.setVerticalSpacing(10)
        yearbook_row = QWidget(); yearbook_row_layout = QHBoxLayout(yearbook_row); yearbook_row_layout.setContentsMargins(0, 0, 0, 0)
        self.yearbook_start_year = QSpinBox(); self.yearbook_start_year.setRange(0, 3000); self.yearbook_start_year.setSpecialValueText("自动"); self.yearbook_start_year.setValue(0)
        self.yearbook_start_month = QSpinBox(); self.yearbook_start_month.setRange(0, 12); self.yearbook_start_month.setSpecialValueText("自动"); self.yearbook_start_month.setValue(1)
        self.yearbook_end_year = QSpinBox(); self.yearbook_end_year.setRange(0, 3000); self.yearbook_end_year.setSpecialValueText("自动"); self.yearbook_end_year.setValue(0)
        self.yearbook_end_month = QSpinBox(); self.yearbook_end_month.setRange(0, 12); self.yearbook_end_month.setSpecialValueText("自动"); self.yearbook_end_month.setValue(12)
        yearbook_row_layout.addWidget(QLabel("起点年")); yearbook_row_layout.addWidget(self.yearbook_start_year); yearbook_row_layout.addWidget(QLabel("月")); yearbook_row_layout.addWidget(self.yearbook_start_month)
        yearbook_row_layout.addSpacing(12); yearbook_row_layout.addWidget(QLabel("终点年")); yearbook_row_layout.addWidget(self.yearbook_end_year); yearbook_row_layout.addWidget(QLabel("月")); yearbook_row_layout.addWidget(self.yearbook_end_month); yearbook_row_layout.addStretch()
        yearbook_layout.addRow("年鉴范围", yearbook_row)
        setup_layout.addWidget(yearbook_box)
        llm_box = QGroupBox("可选：二次裁决"); llm_layout = QFormLayout(llm_box); llm_layout.setHorizontalSpacing(18); llm_layout.setVerticalSpacing(10)
        self.provider = QComboBox(); self.provider.addItems(["不使用外部 LLM", "DeepSeek API"]); llm_layout.addRow("裁决方式", self.provider)
        self.api_key = QLineEdit(); self.api_key.setEchoMode(QLineEdit.Password); self.api_key.setPlaceholderText("仅本次运行使用，不会写入配置文件"); llm_layout.addRow("DeepSeek API Key", self.api_key)
        self.llm_model = QLineEdit("deepseek-chat"); llm_layout.addRow("模型名称", self.llm_model)
        self.base_url = QLineEdit("https://api.deepseek.com"); llm_layout.addRow("API 地址", self.base_url)
        self.budget = QSpinBox(); self.budget.setRange(1, 10000); self.budget.setValue(100); llm_layout.addRow("调用上限", self.budget)
        setup_layout.addWidget(llm_box)
        buttons = QHBoxLayout(); self.run_button = QPushButton("▶  开始运行"); self.run_button.setObjectName("primary"); self.run_button.clicked.connect(self.start_run); buttons.addWidget(self.run_button)
        load_button = QPushButton("打开已有人工复核结果"); load_button.clicked.connect(self.load_existing_review); buttons.addWidget(load_button); buttons.addStretch(); setup_layout.addLayout(buttons)
        self.progress = QProgressBar(); setup_layout.addWidget(self.progress)
        self.status = QLabel("等待上传数据"); setup_layout.addWidget(self.status)
        self.log = QPlainTextEdit(); self.log.setReadOnly(True); self.log.setPlaceholderText("运行日志会显示在这里…"); self.log.setMaximumHeight(140); setup_layout.addWidget(self.log)

        help_tab = QWidget(); help_layout = QVBoxLayout(help_tab); help_layout.setContentsMargins(22, 18, 22, 18); pages.addWidget(help_tab)
        help = QLabel("<h2>怎么开始</h2><p>1. 在“运行配置”选择待审核 Excel 和 codebook。<br>2. 填写年鉴起点和终点年月；缺失任职年月会按这里临时补充，导出前恢复。<br>3. 如果要使用 DeepSeek，选择“DeepSeek API”并填写 Key；否则选择“不使用外部 LLM”。<br>4. 点击“开始运行”，程序完成后会自动切换到“人工复核”。<br>5. 点击“打开候选卡片复核（按名称组）”，同名记录会合并处理。</p><h2>完整工作流程</h2><p>人工复核完成后，点击“保存本批复核结果”。程序会生成独立的“_人工复核完成.xlsx”文件，原始输入文件不变，并自动把这个阶段文件送入“变动审阅”。变动审阅会根据当前 codebook 的 cityname/distname 自动确定范围，结果是建议报告，不会自动改数据。你根据报告在 Excel 中人工修改并保存一个新文件，最后进入“post_rank 审阅”上传这个修改后的文件。上传后，程序会先把 codebook 中标为 00待处理 的部门对应行的 wrank20 统一为 00待处理，再生成 post_rank 建议。</p><h2>人工复核怎么看</h2><p>候选卡片中的每次点击都会先进入“本批暂存”，不会立刻写 Excel。你可以在页面下方检查暂存的标准名称、级别和部门类型；确认无误后再保存。alias 冲突会阻止对应记录写入，并保留在暂存区供修改。</p><h2>结果文件</h2><p>运行目录会生成 reviewed_content.xlsx、review_queue.xlsx、modification_log.xlsx；变动审阅会生成名称有效期和变动阶段问题报告；post_rank 审阅只生成建议，不会自动写回。</p>")
        help.setWordWrap(True); help_layout.addWidget(help); help_layout.addStretch()
        preview_page = QWidget(); preview_layout = QVBoxLayout(preview_page); preview_layout.setContentsMargins(2, 4, 2, 2)
        checklist_box = QGroupBox("流程检查清单")
        checklist_layout = QVBoxLayout(checklist_box)
        self.checklist_label = QLabel(); self.checklist_label.setWordWrap(True); self.checklist_label.setTextFormat(Qt.RichText)
        checklist_layout.addWidget(self.checklist_label); preview_layout.addWidget(checklist_box)
        pages.insertWidget(1, preview_page)
        self._refresh_checklist()
        review_tab = QWidget(); review_layout = QVBoxLayout(review_tab); review_layout.setContentsMargins(2, 4, 2, 2); pages.insertWidget(2, review_tab)
        review_bar = QHBoxLayout(); review_layout.addLayout(review_bar)
        review_bar.addWidget(QLabel("问题类型")); self.review_filter = QComboBox(); self.review_filter.addItems(["全部问题", "错误输入", "有冲突", "多种可能", "缺失输入"]); self.review_filter.currentTextChanged.connect(self._refresh_review); review_bar.addWidget(self.review_filter); review_bar.addStretch()
        self.review_count = QLabel("运行后显示需要人工复核的记录"); review_bar.addWidget(self.review_count)
        self.review_open_button = QPushButton("打开候选卡片复核（按名称组）"); self.review_open_button.setObjectName("primary"); self.review_open_button.clicked.connect(self.open_human_review); review_bar.addWidget(self.review_open_button)
        self.review_table = QTableWidget(); self._configure_table(self.review_table); review_layout.addWidget(self.review_table)
        batch_box = QGroupBox("本批暂存的审核与修改（尚未写入文件）")
        batch_layout = QVBoxLayout(batch_box)
        self.batch_count = QLabel("当前没有暂存修改")
        batch_layout.addWidget(self.batch_count)
        self.batch_table = QTableWidget(); self._configure_table(self.batch_table); self.batch_table.setMaximumHeight(190); batch_layout.addWidget(self.batch_table)
        batch_buttons = QHBoxLayout(); batch_buttons.addStretch()
        self.edit_batch_button = QPushButton("编辑暂存审核"); self.edit_batch_button.clicked.connect(self.edit_pending_batch); batch_buttons.addWidget(self.edit_batch_button)
        self.delete_batch_button = QPushButton("删除选中的暂存"); self.delete_batch_button.clicked.connect(self.delete_selected_pending); batch_buttons.addWidget(self.delete_batch_button)
        self.clear_batch_button = QPushButton("清空本批暂存"); self.clear_batch_button.clicked.connect(self.clear_pending_batch); batch_buttons.addWidget(self.clear_batch_button)
        self.save_batch_button = QPushButton("保存本批复核结果"); self.save_batch_button.setObjectName("primary"); self.save_batch_button.clicked.connect(self.save_pending_batch); batch_buttons.addWidget(self.save_batch_button)
        batch_layout.addLayout(batch_buttons)
        self.save_progress = QProgressBar(); self.save_progress.setRange(0, 100); self.save_progress.setValue(0); self.save_progress.setFormat("尚未开始保存"); self.save_progress.setVisible(False); batch_layout.addWidget(self.save_progress)
        review_layout.addWidget(batch_box)

        change_tab = QWidget(); change_layout = QVBoxLayout(change_tab); change_layout.setContentsMargins(2, 4, 2, 2); change_layout.setSpacing(10); pages.insertWidget(3, change_tab)
        change_notice = QLabel("⚠ 进入变动审核前，请确认当前‘标准 codebook’已经是人工复核后的最新版本。若人工审核过程中新增、修改或合并了 codebook 记录，请先在‘开始运行’页面重新选择并确认新的 codebook，再上传人工复核完成文件。变动审核会以这两个文件的当前版本为准。")
        change_notice.setWordWrap(True); change_notice.setObjectName("stageNotice"); change_notice.setMinimumHeight(54); change_layout.addWidget(change_notice)
        change_intro = QLabel("本卡接收人工复核保存后生成的‘_人工复核完成.xlsx’；根据当前 codebook 的 cityname/distname 自动确定审阅范围，检查 w_name、s1_sector、wrank_t，并在扩充完成后补填覆盖 2020 年的 wrank20。结果只提供人工审核建议，不自动修改原始数据。")
        change_intro.setWordWrap(True); change_intro.setObjectName("hint"); change_layout.addWidget(change_intro)
        change_bar = QHBoxLayout(); self.change_count = QLabel("请先运行，或加载‘_人工复核完成.xlsx’"); change_bar.addWidget(self.change_count); change_bar.addStretch(); self.change_stage_button = QPushButton("加载人工复核完成文件"); self.change_stage_button.clicked.connect(self._load_human_reviewed_stage); change_bar.addWidget(self.change_stage_button); self.change_refresh_button = QPushButton("刷新变动审阅"); self.change_refresh_button.setObjectName("primary"); self.change_refresh_button.clicked.connect(self._load_change_audit); change_bar.addWidget(self.change_refresh_button); self.change_export_button = QPushButton("导出扩行结果"); self.change_export_button.clicked.connect(self._export_change_expansion); change_bar.addWidget(self.change_export_button); change_layout.addLayout(change_bar)
        # Keep the change-review card focused on the two views that are
        # actually needed for the user's manual decision: the detailed issue
        # list and the compact summary.  The codebook/department/demo tables
        # remain available to the backend for report generation, but are no
        # longer placed in the crowded main view.
        self.change_codebook_box = QGroupBox("codebook 中包含有效期的记录"); cb_layout = QVBoxLayout(self.change_codebook_box); self.change_codebook_table = QTableWidget(); self._configure_table(self.change_codebook_table); cb_layout.addWidget(self.change_codebook_table)
        self.change_department_box = QGroupBox("经历过级别、类别或名称变动的部门清单"); dept_layout = QVBoxLayout(self.change_department_box); self.change_department_table = QTableWidget(); self._configure_table(self.change_department_table); dept_layout.addWidget(self.change_department_table)
        self.change_issue_box = QGroupBox("名称有效期和变动阶段问题"); issue_layout = QVBoxLayout(self.change_issue_box); self.change_issue_table = QTableWidget(); self._configure_table(self.change_issue_table); self.change_issue_table.setMinimumHeight(360); issue_layout.addWidget(self.change_issue_table); change_layout.addWidget(self.change_issue_box, 1)
        self.change_summary_box = QGroupBox("变动审核汇总"); sum_layout = QVBoxLayout(self.change_summary_box); self.change_summary_label = QLabel("运行后显示汇总"); self.change_summary_label.setWordWrap(True); self.change_summary_label.setMinimumHeight(105); sum_layout.addWidget(self.change_summary_label); change_layout.addWidget(self.change_summary_box)

        post_tab = QWidget(); post_layout = QVBoxLayout(post_tab); post_layout.setContentsMargins(2,4,2,2); pages.insertWidget(4, post_tab)
        post_layout.addWidget(QLabel("请先在变动审核中人工修改并保存 Excel，再在这里上传修改后的文件。上传后程序会先按 codebook 自动把匹配到 00待处理 部门的 wrank20 统一为 00待处理，并生成独立阶段文件；随后读取当前 wrank_t，使用 4post_rank.R 的四组职位词表识别 post1，推导理论 post_rank 并列出不一致与未知职位。post_rank 本身不会自动写回。"))
        post_bar=QHBoxLayout(); self.post_file_label=QLabel("尚未上传变动审核完成文件"); post_bar.addWidget(self.post_file_label); self.post_upload_button=QPushButton("上传并生成 post_rank 审阅"); self.post_upload_button.setObjectName("primary"); self.post_upload_button.clicked.connect(self._load_post_rank_review); post_bar.addWidget(self.post_upload_button); self.post_finalize_button=QPushButton("安全校验并生成最终成稿"); self.post_finalize_button.setEnabled(False); self.post_finalize_button.clicked.connect(self._finalize_post_rank); post_bar.addWidget(self.post_finalize_button); post_layout.addLayout(post_bar)
        self.post_count=QLabel("等待上传"); post_layout.addWidget(self.post_count); self.post_table=QTableWidget(); self._configure_table(self.post_table); post_layout.addWidget(self.post_table)

    def _configure_table(self, table: QTableWidget) -> None:
        table.setAlternatingRowColors(True); table.setSelectionBehavior(QAbstractItemView.SelectRows); table.setEditTriggers(QAbstractItemView.NoEditTriggers); table.setWordWrap(False)
        table.verticalHeader().setVisible(False); table.horizontalHeader().setStretchLastSection(True); table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive); table.setSortingEnabled(True)

    def _refresh_checklist(self) -> None:
        result = getattr(self, "result", None) or {}
        human_stage = bool(self.source_path and self.source_path.exists() and "人工复核完成" in self.source_path.stem)
        change_ready = bool(getattr(self, "_change_audit", None))
        post_ready = bool(getattr(self, "_post_uploaded_path", None))
        final_path = getattr(self, "_final_output_path", None)
        final_ready = bool(final_path and Path(final_path).exists())
        rows = [(bool(result), "初始匹配"), (human_stage, "人工复核阶段文件"), (change_ready, "变动审阅报告"), (post_ready, "post_rank 审阅文件"), (final_ready, "最终成稿")]
        self.checklist_label.setText("<br>".join(("✅ " if done else "○ ") + label for done, label in rows))

    def _path_row(self, edit: QLineEdit, title: str, save: bool = True) -> QWidget:
        widget = QWidget(); row = QHBoxLayout(widget); row.setContentsMargins(0, 0, 0, 0); row.addWidget(edit)
        button = QPushButton("浏览…"); button.clicked.connect(lambda: self._browse(edit, title, save)); row.addWidget(button); return widget

    def _browse(self, edit: QLineEdit, title: str, save: bool) -> None:
        path, _ = QFileDialog.getOpenFileName(self, title, "", "Excel/CSV (*.xlsx *.xls *.csv);;All files (*)")
        if path: edit.setText(path)

    def start_run(self) -> None:
        paths = [Path(self.input.text()), Path(self.codebook.text()), Path(self.config.text())]
        if not paths[0].exists() or not paths[1].exists() or not paths[2].exists():
            QMessageBox.warning(self, "文件不存在", "请检查输入文件、codebook 和 config 路径。"); return
        if self.provider.currentText() == "DeepSeek API" and not self.api_key.text().strip():
            QMessageBox.warning(self, "缺少 API Key", "选择 DeepSeek 时必须填写 API Key。"); return
        self.run_button.setEnabled(False); self.log.clear(); self.progress.setValue(0); self.status.setText("准备运行…")
        self.header_status.setText("● 运行中"); self.header_status.setStyleSheet("color:#9b641c;background:#fff0d8;border-radius:12px;padding:6px 12px;font-weight:700;")
        self.worker_thread = QThread(self); self.worker = RunWorker(paths[0], paths[1], paths[2], self.model_path.text(), self.provider.currentText(), self.api_key.text(), self.llm_model.text(), self.base_url.text(), 30.0, self.budget.value())
        self.worker.moveToThread(self.worker_thread); self.worker_thread.started.connect(self.worker.run); self.worker.progress.connect(self.on_progress); self.worker.finished.connect(self.on_finished); self.worker.failed.connect(self.on_failed); self.worker.finished.connect(self.worker_thread.quit); self.worker.failed.connect(self.worker_thread.quit); self.worker_thread.finished.connect(self.worker.deleteLater); self.worker_thread.finished.connect(self.thread_done); self.worker_thread.start()

    @Slot(str, int, int)
    def on_progress(self, stage: str, done: int, total: int) -> None:
        self.status.setText(f"{stage}：{done}/{total}"); self.log.appendPlainText(self.status.text()); self.progress.setValue(int(done * 100 / total) if total else 0)

    @Slot(object)
    def on_finished(self, result: dict) -> None:
        self.result = result; self.source_path = Path(result.get("input_path", "")) if result.get("input_path") else self.source_path; self.debug_frame = pd.read_json(Path(result["output_dir"]) / "debug" / "queries.jsonl", lines=True); self._scoped_review_frame = None; self._scoped_review_source = None; self.progress.setValue(100); self.status.setText(f"完成：{result['input_rows']:,} 行，{result['unique_queries']:,} 条去重查询"); self.log.appendPlainText(self.status.text()); self.header_status.setText("● 已完成"); self.header_status.setStyleSheet("color:#2a8b63;background:#e2f4eb;border-radius:12px;padding:6px 12px;font-weight:700;"); self._refresh_checklist(); self.nav_list.setCurrentRow(2)

    @staticmethod
    def _as_int(value):
        try:
            if value is None or pd.isna(value) or str(value).strip() == "": return None
            return int(float(value))
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _as_text(value):
        try:
            if value is None or pd.isna(value): return ""
        except (TypeError, ValueError):
            pass
        return str(value).strip()

    @classmethod
    def _month_key(cls, year, month=None, default_month=1):
        y = cls._as_int(year)
        if y is None: return None
        m = cls._as_int(month) or default_month
        return y * 12 + max(1, min(12, m)) - 1

    @classmethod
    def _display_period(cls, row):
        def text(value):
            return "空" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)
        start = f"{text(row.get('year_valid_from'))}-{text(row.get('month_valid_from'))}"
        end = f"{text(row.get('year_valid_to'))}-{text(row.get('month_valid_to'))}"
        return f"{start} 至 {end}"

    @staticmethod
    def _row_matches_codebook_scope(data_row: pd.Series, codebook: Codebook) -> bool:
        """Return whether a data row belongs to any codebook city/dist scope.

        Empty codebook scope cells are treated as wildcards.  This replaces the
        former Guangzhou-only filter and is shared by change and post-rank
        stages.
        """
        if not {"cityname", "distname"}.issubset(codebook.frame.columns):
            return True
        city = str(data_row.get("cityname", "") or "").strip()
        dist = str(data_row.get("distname", "") or "").strip()
        for _, code_row in codebook.frame.iterrows():
            code_city = split_multi_value(code_row.get("cityname", ""))
            code_dist = split_multi_value(code_row.get("distname", ""))
            city_ok = not code_city or city in code_city
            dist_ok = not code_dist or dist in code_dist
            if city_ok and dist_ok:
                return True
        return False

    def _filter_codebook_scope(self, frame: pd.DataFrame, codebook: Codebook) -> pd.DataFrame:
        if frame.empty:
            return frame.copy()
        # Build the allowed geographic pairs once.  The former row-wise
        # implementation scanned every codebook row for every data row.
        exact_pairs: set[tuple[str, str]] = set()
        wildcard_scopes: list[tuple[set[str], set[str]]] = []
        for row in codebook.valid_rows():
            cities = {normalize_name(value) for value in codebook.split_values(row.cityname) if normalize_name(value)}
            dists = {normalize_name(value) for value in codebook.split_values(row.distname) if normalize_name(value)}
            if cities and dists:
                exact_pairs.update((city, dist) for city in cities for dist in dists)
            else:
                wildcard_scopes.append((cities, dists))
        city_values = frame.get("cityname", pd.Series("", index=frame.index)).map(lambda value: normalize_name(self._as_text(value)))
        dist_values = frame.get("distname", pd.Series("", index=frame.index)).map(lambda value: normalize_name(self._as_text(value)))
        mask = pd.Series([(city, dist) in exact_pairs for city, dist in zip(city_values, dist_values)], index=frame.index)
        for cities, dists in wildcard_scopes:
            mask |= ((~city_values.isin(cities)) if not cities else city_values.isin(cities)) & ((~dist_values.isin(dists)) if not dists else dist_values.isin(dists))
        return frame.loc[mask].copy()

    def _load_change_audit(self) -> None:
        if self.source_path is None or not self.source_path.exists():
            self.change_count.setText("请先选择或运行输入数据"); return
        self.change_refresh_button.setEnabled(False)
        self.change_count.setText("正在读取阶段文件并生成变动审核，请稍候…")
        self.status.setText("正在生成变动审核：读取文件、匹配范围并处理有效期…")
        self.header_status.setText("● 变动审核处理中")
        self.header_status.setStyleSheet("color:#9b641c;background:#fff0d8;border-radius:12px;padding:6px 12px;font-weight:700;")
        self.progress.setRange(0, 100); self.progress.setValue(0); self.progress.setFormat("变动审核：准备开始")
        QApplication.processEvents()
        try:
            if self.source_path.suffix.lower() != ".csv":
                # The change audit only needs these fields.  Avoid loading
                # wide workbooks and formula-heavy columns into the UI thread.
                wanted = {"year", "cityname", "distname", "w_name", "w_name_o", "wrank20", "wrank_t", "post_rank", "p_name", "styr", "stmo", "fiyr", "fimo", "record_id"}
                header = pd.read_excel(self.source_path, dtype=object, nrows=0)
                usecols = [column for column in header.columns if column in wanted]
                input_frame = pd.read_excel(self.source_path, dtype=object, usecols=usecols)
            else:
                input_frame = pd.read_csv(self.source_path, dtype=object, usecols=lambda column: column in {"year", "cityname", "distname", "w_name", "w_name_o", "wrank20", "wrank_t", "post_rank", "p_name", "styr", "stmo", "fiyr", "fimo", "record_id"})
            # Duplicate Excel headers make row["field"] return a DataFrame
            # instead of a scalar and later boolean checks become ambiguous.
            duplicate_headers = [str(name) for name in input_frame.columns[input_frame.columns.duplicated()].tolist()]
            conflicting_headers = []
            for name in dict.fromkeys(duplicate_headers):
                repeated = input_frame.loc[:, input_frame.columns == name]
                if len(repeated.columns) > 1 and not repeated.nunique(dropna=False).le(1).all():
                    conflicting_headers.append(name)
            if conflicting_headers:
                raise ValueError("变动审核无法继续：输入文件存在内容不一致的重复列名：" + "、".join(conflicting_headers) + "。请先在 Excel 中保留唯一列名。")
            if duplicate_headers:
                input_frame = input_frame.loc[:, ~input_frame.columns.duplicated()].copy()
            self.progress.setValue(15); self.progress.setFormat("变动审核：已读取阶段文件　15%"); QApplication.processEvents()
            cb = Codebook(Path(self.codebook.text()))
            self.progress.setValue(25); self.progress.setFormat("变动审核：已读取 codebook　25%"); QApplication.processEvents()
            # Scope is determined by the current codebook, not by a hardcoded
            # city.  Blank city/dist cells in codebook act as wildcards.
            input_frame = self._filter_codebook_scope(input_frame, cb)
            self.progress.setValue(40); self.progress.setFormat("变动审核：已完成 cityname/distname 范围筛选　40%"); QApplication.processEvents()
            change_rows = [row for row in cb.valid_rows() if not row.all_time_empty]
            cb_records = []
            for row in change_rows:
                cb_records.append({"cityname": row.cityname, "distname": row.distname, "s1_sector": row.s1_sector, "wrank": row.wrank, "w_name": row.w_name, "有效期": self._display_period({"year_valid_from": row.year_valid_from, "month_valid_from": row.month_valid_from, "year_valid_to": row.year_valid_to, "month_valid_to": row.month_valid_to}), "notes": row.notes})
            self._fill_table(self.change_codebook_table, pd.DataFrame(cb_records), ["cityname", "distname", "s1_sector", "wrank", "w_name", "有效期", "notes"])
            department_records = []
            grouped_codebook = {}
            for row in change_rows:
                key = (row.cityname, row.distname, row.w_name)
                grouped_codebook.setdefault(key, []).append(row)
            for (cityname, distname, w_name), items in grouped_codebook.items():
                sectors = sorted({str(item.s1_sector).strip() for item in items if str(item.s1_sector).strip()})
                ranks = sorted({str(item.wrank).strip() for item in items if str(item.wrank).strip()})
                names = sorted({str(item.w_name).strip() for item in items if str(item.w_name).strip()})
                fields = []
                if len(names) > 1: fields.append("w_name")
                if len(sectors) > 1: fields.append("s1_sector")
                if len(ranks) > 1: fields.append("wrank_t")
                notes_text = "；".join(str(item.notes).strip() for item in items if str(item.notes).strip())
                if "更名" in notes_text or "名称" in notes_text:
                    if "w_name" not in fields: fields.append("w_name（备注提示）")
                if not fields: fields.append("有效期分段，待核实")
                department_records.append({"cityname": cityname, "distname": distname, "w_name": w_name, "变化字段": "、".join(fields), "s1_sector范围": " → ".join(sectors), "wrank_t范围": " → ".join(ranks), "有效期段数": len(items), "备注": notes_text})
            self.change_departments_frame = pd.DataFrame(department_records)
            self._fill_table(self.change_department_table, self.change_departments_frame, ["cityname", "distname", "w_name", "变化字段", "s1_sector范围", "wrank_t范围", "有效期段数", "备注"])
            self.progress.setValue(55); self.progress.setFormat("变动审核：已识别 codebook 变动部门　55%"); QApplication.processEvents()
            audit, transformed, _ = self._build_change_preview_expanded(input_frame, cb)
            self.progress.setValue(85); self.progress.setFormat("变动审核：已完成有效期匹配与任职区间拆分　85%"); QApplication.processEvents()
            # 页面按部门汇总“级别未变但 wrank 不统一”，避免同一部门的
            # 年度记录重复占满审核表；逐行明细仍保留在扩充结果文件中。
            if not audit.empty and "issue_type" in audit.columns:
                mismatch = audit["issue_type"].astype(str).eq("级别没变但是wrank不统一")
                if mismatch.any():
                    detail = audit.loc[mismatch].copy()
                    group_cols = [c for c in ["cityname", "distname", "w_name"] if c in detail.columns]
                    if group_cols:
                        summaries = []
                        for _, group in detail.groupby(group_cols, dropna=False, sort=False):
                            first = group.iloc[0].copy()
                            wrt = "、".join(sorted({str(v).strip() for v in group.get("wrank_t", []) if str(v).strip() and str(v).strip().lower() != "nan"}))
                            wr20 = "、".join(sorted({str(v).strip() for v in group.get("wrank20", []) if str(v).strip() and str(v).strip().lower() != "nan"}))
                            first["reason"] = f"该部门级别没变但 wrank_t 是{wrt or '空'}，wrank20 是{wr20 or '空'}"
                            first["条数"] = len(group)
                            summaries.append(first)
                        summary_frame = pd.DataFrame(summaries)
                        audit = pd.concat([audit.loc[~mismatch], summary_frame], ignore_index=True)
            self.change_audit_frame = audit; self.change_transformed_frame = transformed; self._change_audit = audit; self._refresh_checklist()
            if self.result and self.result.get("output_dir") and not audit.empty:
                temporal_rows = audit[audit["issue_type"].astype(str).isin({"w_name超出有效期", "w_name有效期无法判断"})].copy()
                if not temporal_rows.empty:
                    temporal_path = Path(self.result["output_dir"]) / "temporal_review_queue.xlsx"
                    temporal_rows.to_excel(temporal_path, index=False)
                    review_queue_value = self.result.get("review_queue")
                    review_queue = Path(review_queue_value) if review_queue_value else None
                    if review_queue is not None and review_queue.exists():
                        existing_queue = pd.read_excel(review_queue, dtype=object)
                        merged_queue = pd.concat([existing_queue, temporal_rows], ignore_index=True)
                        dedupe = [column for column in ["record_id", "source_index", "issue_type", "w_name", "year"] if column in merged_queue.columns]
                        if dedupe: merged_queue = merged_queue.drop_duplicates(subset=dedupe, keep="first")
                        merged_queue.to_excel(review_queue, index=False)
            issue_counts = audit["issue_type"].value_counts().to_dict() if not audit.empty else {}
            self._fill_table(self.change_issue_table, audit, ["year", "p_name", "cityname", "distname", "w_name", "issue_type", "reason"])
            self.change_count.setText(f"codebook 变动记录 {len(change_rows):,} 条；检测到变动相关输入 {getattr(self, '_change_input_count', 0):,} 条；名称超出有效期 {issue_counts.get('w_name超出有效期', 0):,} 条；级别未变但 wrank 不统一 {issue_counts.get('级别没变但是wrank不统一', 0):,} 条")
            self.change_summary_label.setText(self._format_change_summary(audit))
            self.progress.setValue(100); self.progress.setFormat("变动审核：报告生成完成　100%")
        except Exception as exc:
            self.change_count.setText(f"变动审阅生成失败：{exc}")
            self.status.setText(f"变动审阅失败：{exc}")
        finally:
            self.change_refresh_button.setEnabled(True)
            self.progress.setRange(0, 100)
            if not self.change_count.text().startswith("变动审阅生成失败"):
                self.status.setText("变动审核完成，可以查看报告并人工处理。")
                self.header_status.setText("● 变动审核完成")
                self.header_status.setStyleSheet("color:#2a8b63;background:#e2f4eb;border-radius:12px;padding:6px 12px;font-weight:700;")
            QApplication.processEvents()

    def _load_human_reviewed_stage(self) -> None:
        """Load a saved human-review stage without rerunning the matcher."""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择人工复核完成文件",
            "",
            "人工复核完成文件 (*.xlsx *.xlsm);;Excel (*.xlsx *.xlsm)",
        )
        if not path:
            return
        selected = Path(path)
        self.source_path = selected
        self.result = None
        self.status.setText(f"已选择文件：{selected.name}，即将开始变动审核…")
        self.log.appendPlainText(self.status.text())
        self.nav_list.setCurrentRow(3)
        # Let Qt paint the selected-file message before starting the heavy
        # audit.  Calling the audit synchronously here made the whole window
        # appear frozen immediately after the native file dialog closed.
        QTimer.singleShot(0, self._load_change_audit)

    @staticmethod
    def _format_change_summary(audit: pd.DataFrame) -> str:
        if audit is None or audit.empty: return "共发现 0 条变动审核问题。"
        counts=audit.get("issue_type",pd.Series(dtype=object)).value_counts().to_dict()
        lines=[f"共发现 {len(audit):,} 条：", f"名称超出有效期：{counts.get('w_name超出有效期',0):,} 条", f"唯一变动建议：{counts.get('变动建议',0):,} 条", f"级别未变但 wrank 不统一：{counts.get('级别没变但是wrank不统一',0):,} 条", "", "处理顺序：先确认 codebook 有效期，再核对历史名称，最后由人工决定修改 codebook 或原始数据。", "以上均为系统建议，没有自动写回。"]
        return "\n".join(lines)

    def _load_post_rank_review(self) -> None:
        path,_=QFileDialog.getOpenFileName(self,"选择已完成变动审核的 Excel","","Excel (*.xlsx *.xlsm)")
        if not path: return
        self.post_upload_button.setEnabled(False); self.post_count.setText("正在读取文件并生成 post_rank 审阅，请稍候…"); QApplication.processEvents()
        try:
            # Before post_rank review, apply the deterministic 2020-level
            # cleanup requested by the user.  The uploaded file is preserved;
            # the next-stage copy receives the wrank20 updates.
            prepared_path, wrank20_changes = self._prepare_wrank20_stage(Path(path))
            self._post_uploaded_path = Path(path)
            baseline = Path(path).with_name(Path(path).stem + "_post_rank基线" + Path(path).suffix)
            if not baseline.exists():
                shutil.copy2(path, baseline)
            self._post_baseline_path = baseline
            full_frame = pd.read_excel(prepared_path,dtype=object)
            if not {"cityname","distname"}.issubset(full_frame.columns): raise ValueError("文件缺少 cityname 或 distname")
            post_codebook = Codebook(Path(self.codebook.text()))
            frame = self._filter_codebook_scope(full_frame, post_codebook)
            temp=prepared_path; old=self.source_path; self.source_path=temp
            logs,_=self._audit_post_rank_file(temp)
            self.source_path=old
            post_audit = audit_post_rank_frame(frame)
            self.post_file_label.setText(f"原文件：{path}\n阶段文件：{prepared_path}")
            self.post_count.setText(f"已自动统一 {wrank20_changes:,} 行 wrank20=00待处理；发现 {len(post_audit):,} 条职位级别需人工核查；另有 {len(logs):,} 条部门级别联动建议。")
            if not post_audit.empty:
                self._fill_table(self.post_table, post_audit, ["year", "cityname", "distname", "w_name", "post1", "wrank_t", "职位类别", "理论post_rank", "原post_rank", "issue_type"])
            else:
                self._fill_table(self.post_table,pd.DataFrame(),["year", "w_name", "post1", "wrank_t", "理论post_rank", "原post_rank", "issue_type"])
            self.post_finalize_button.setEnabled(True); self._refresh_checklist()
        except Exception as exc: self.post_count.setText(f"post_rank 审阅失败：{exc}")
        finally:
            self.post_upload_button.setEnabled(True); QApplication.processEvents()

    def _finalize_post_rank(self) -> None:
        """Validate the manually edited post_rank file and create final output."""
        uploaded = getattr(self, "_post_uploaded_path", None)
        baseline = getattr(self, "_post_baseline_path", None)
        if uploaded is None or baseline is None or not Path(uploaded).exists() or not Path(baseline).exists():
            QMessageBox.information(self, "尚未准备完成", "请先上传修改后的 post_rank 文件。")
            return
        try:
            source_book = load_workbook(baseline, data_only=False)
            edited_book = load_workbook(uploaded, data_only=False)
            if source_book.sheetnames != edited_book.sheetnames:
                raise ValueError("安全校验失败：工作表集合发生变化")
            allowed = {"post_rank"}
            changes = []
            for sheet_name in source_book.sheetnames:
                old_sheet, new_sheet = source_book[sheet_name], edited_book[sheet_name]
                if old_sheet.max_row != new_sheet.max_row or old_sheet.max_column != new_sheet.max_column:
                    raise ValueError(f"安全校验失败：{sheet_name} 行列数量发生变化")
                headers = {str(old_sheet.cell(1, c).value).strip(): c for c in range(1, old_sheet.max_column + 1) if old_sheet.cell(1, c).value is not None}
                if "post_rank" not in headers:
                    continue
                post_col = headers["post_rank"]
                for row in range(2, old_sheet.max_row + 1):
                    old_value, new_value = old_sheet.cell(row, post_col).value, new_sheet.cell(row, post_col).value
                    if old_value != new_value:
                        changes.append({"sheet": sheet_name, "source_row": row, "field": "post_rank", "before": old_value, "after": new_value})
                for col in range(1, old_sheet.max_column + 1):
                    header = str(old_sheet.cell(1, col).value or "").strip()
                    if header in allowed or not header:
                        continue
                    for row in range(1, old_sheet.max_row + 1):
                        if old_sheet.cell(row, col).value != new_sheet.cell(row, col).value:
                            raise ValueError(f"安全校验失败：{header}!{row} 被修改；post_rank 阶段只允许修改 post_rank")
            final_path = Path(uploaded).with_name(Path(uploaded).stem.replace("_post_rank基线", "") + "_最终成稿" + Path(uploaded).suffix)
            shutil.copy2(uploaded, final_path)
            log_path = final_path.with_name(final_path.stem + "_post_rank修改日志.xlsx")
            pd.DataFrame(changes, columns=["sheet", "source_row", "field", "before", "after"]).to_excel(log_path, index=False)
            self.post_count.setText(f"安全校验通过：{len(changes):,} 条 post_rank 修改；最终成稿已生成：{final_path.name}")
            self._final_output_path = final_path; self._refresh_checklist()
            QMessageBox.information(self, "最终成稿已生成", f"{final_path}\n\n修改日志：{log_path}")
        except Exception as exc:
            QMessageBox.critical(self, "最终成稿未生成", str(exc))

    def _prepare_wrank20_stage(self, path: Path) -> tuple[Path, int]:
        """Create a post-review stage file with codebook 00-rank cleanup.

        This is deliberately independent of validity intervals: ``wrank20``
        represents the 2020 department level.  Historical names are matched
        through codebook aliases, with the steel-factory name transition kept
        as an explicit compatibility rule.
        """
        codebook_path = Path(self.codebook.text())
        cb = Codebook(codebook_path)
        zero_rows = []
        for row in cb.valid_rows():
            rank = normalize_name(row.wrank)
            if rank in {"00待出", "00待处理"} or "00待出" in rank:
                variants = {normalize_name(row.w_name), *(normalize_name(alias) for alias in row.aliases)} - {""}
                if normalize_name(row.w_name) == normalize_name("钢铁企业集团") or normalize_name("广州钢铁厂") in variants:
                    variants.update({normalize_name("钢铁厂"), normalize_name("广州钢铁集团"), normalize_name("广州钢铁集团有限公司")})
                zero_rows.append((row, variants))
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("wrank20 自动处理目前只支持 xlsx/xlsm")
        output = path.with_name(f"{path.stem}_00待处理完成{path.suffix}")
        shutil.copy2(path, output)
        book = load_workbook(output, keep_vba=path.suffix.lower() == ".xlsm")
        sheet = book[book.sheetnames[0]]
        headers = {str(sheet.cell(1, col).value).strip(): col for col in range(1, sheet.max_column + 1) if sheet.cell(1, col).value is not None}
        required = ["wrank20", "w_name", "cityname", "distname"]
        missing = [field for field in required if field not in headers]
        if missing:
            raise ValueError(f"文件缺少字段：{missing}")
        w20_col, name_col = headers["wrank20"], headers["w_name"]
        original_col = headers.get("w_name_o")
        city_col, dist_col = headers["cityname"], headers["distname"]

        def scope_values(value):
            return {normalize_name(part) for part in re.split(r"\s*[|｜、,，;；]\s*", str(value or "")) if normalize_name(part)}

        def applies(row, city, dist):
            cities, dists = scope_values(row.cityname), scope_values(row.distname)
            return (not cities or normalize_name(city) in cities) and (not dists or normalize_name(dist) in dists)

        changed = 0
        for excel_row in range(2, sheet.max_row + 1):
            city, dist = sheet.cell(excel_row, city_col).value, sheet.cell(excel_row, dist_col).value
            names = {normalize_name(sheet.cell(excel_row, name_col).value)}
            if original_col:
                names.add(normalize_name(sheet.cell(excel_row, original_col).value))
            if any(applies(row, city, dist) and names & variants for row, variants in zero_rows):
                if normalize_name(sheet.cell(excel_row, w20_col).value) != normalize_name("00待处理"):
                    sheet.cell(excel_row, w20_col).value = "00待处理"
                    changed += 1
        book.save(output)
        return output, changed

    def _audit_post_rank_file(self, path: Path):
        obj=MainWindow.__new__(MainWindow)
        return MainWindow._apply_temporal_postprocess(obj,path,Path(self.codebook.text()))

    def _fill_table(self, table: QTableWidget, frame: pd.DataFrame, columns: list[str]) -> None:
        frame = frame.copy() if frame is not None else pd.DataFrame()
        columns = [column for column in columns if column in frame.columns]
        table.setSortingEnabled(False); table.setRowCount(min(len(frame), 500)); table.setColumnCount(len(columns)); table.setHorizontalHeaderLabels(columns)
        for i, (_, row) in enumerate(frame.head(500).iterrows()):
            for j, column in enumerate(columns):
                value = row.get(column, "")
                if value is None or (isinstance(value, float) and pd.isna(value)): value = ""
                table.setItem(i, j, QTableWidgetItem(str(value)))
        # Measuring every cell is expensive for 500-row audit tables. The
        # header remains interactive so the user can resize a column when
        # needed; small tables still get a one-time automatic fit.
        if len(frame) <= 80:
            table.resizeColumnsToContents()
        table.setSortingEnabled(True)

    def _build_change_preview_expanded(self, frame: pd.DataFrame, cb: Codebook):
        """Run the standalone department expansion engine for the change card.

        The engine returns full-column processed rows, a problem queue, and a
        split log.  The UI remains read-only: only the explicit export button
        writes these results to a new workbook.
        """
        settings = {}
        try:
            import yaml
            config_path = Path(self.config.text())
            if config_path.exists():
                raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
                settings = dict(raw.get("department_expansion", {}) or {})
        except Exception:
            settings = {}
        ui_start_year = self.yearbook_start_year.value() if hasattr(self, "yearbook_start_year") else 0
        ui_start_month = self.yearbook_start_month.value() if hasattr(self, "yearbook_start_month") else 1
        ui_end_year = self.yearbook_end_year.value() if hasattr(self, "yearbook_end_year") else 0
        ui_end_month = self.yearbook_end_month.value() if hasattr(self, "yearbook_end_month") else 12
        if ui_start_year > 0:
            settings["yearbook_year_start"] = ui_start_year
        if ui_end_year > 0:
            settings["yearbook_year_end"] = ui_end_year
        if ui_start_month > 0:
            settings["yearbook_month_start"] = ui_start_month
        if ui_end_month > 0:
            settings["yearbook_month_end"] = ui_end_month
        if settings.get("yearbook_year_start") and settings.get("yearbook_year_end"):
            start_key = int(settings["yearbook_year_start"]) * 12 + int(settings.get("yearbook_month_start", 1)) - 1
            end_key = int(settings["yearbook_year_end"]) * 12 + int(settings.get("yearbook_month_end", 12)) - 1
            if start_key > end_key:
                raise ValueError("年鉴起止时间无效：起点必须早于或等于终点")
        years = pd.to_numeric(frame.get("year", pd.Series(dtype=object)), errors="coerce").dropna()
        if len(years):
            if settings.get("yearbook_year_start") is None:
                settings["yearbook_year_start"] = int(years.min())
            if settings.get("yearbook_year_end") is None:
                settings["yearbook_year_end"] = int(years.max())
        if settings.get("yearbook_month_start") is None:
            settings["yearbook_month_start"] = 1
        if settings.get("yearbook_month_end") is None:
            settings["yearbook_month_end"] = 12
        if settings.get("enable_backfill") is None:
            settings["enable_backfill"] = True
        expansion_cfg = EXPANSION_DEFAULT_CONFIG.copy()
        expansion_cfg.update({
            "yearbook_year_start": settings.get("yearbook_year_start"),
            "yearbook_month_start": settings.get("yearbook_month_start", 1),
            "yearbook_year_end": settings.get("yearbook_year_end"),
            "yearbook_month_end": settings.get("yearbook_month_end", 12),
            "enable_backfill": settings.get("enable_backfill", True),
        })
        for key in expansion_cfg:
            if key in settings:
                expansion_cfg[key] = settings[key]
        result_processed, result_issues, result_split = process_codebook_data(cb.frame.copy(), frame.copy(), expansion_cfg)
        result = SimpleNamespace(
            processed=result_processed,
            issues=result_issues,
            split_log=result_split,
            summary=build_card_summary(result_processed, result_issues, result_split, len(frame)),
        )
        self._last_expansion_result = result
        audit = result.issues.copy()
        if "audit_result" in audit.columns:
            audit["issue_type"] = audit["audit_result"]
            audit["reason"] = audit["audit_result"]
        self._change_input_count = len(audit)
        transformed = result.processed.copy()
        # The UI only needs the audit table and full expansion export.  Do not
        # build random demonstration samples from the expanded frame: that
        # caused another groupby and copied up to 20 rows per split group.
        return audit, transformed, []

    def _export_change_expansion(self) -> None:
        result = getattr(self, "_last_expansion_result", None)
        if result is None:
            QMessageBox.information(self, "尚未生成结果", "请先运行或刷新变动审核。")
            return
        default = "变化已经扩充.xlsx"
        if self.source_path is not None:
            default = str(self.source_path.with_name(self.source_path.stem + "_变化已经扩充.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, "导出部门扩充结果", default, "Excel (*.xlsx)")
        if not path:
            return
        target = Path(path)
        if self.source_path is not None and target.resolve() == self.source_path.resolve():
            QMessageBox.warning(self, "禁止覆盖输入文件", "扩行结果必须导出为新文件，不能覆盖当前输入文件。")
            return
        try:
            with pd.ExcelWriter(target, engine="openpyxl") as writer:
                result.processed.to_excel(writer, sheet_name="处理结果", index=False)
                result.issues.to_excel(writer, sheet_name="问题清单", index=False)
                result.split_log.to_excel(writer, sheet_name="拆分日志", index=False)
                result.summary.to_excel(writer, sheet_name="卡片摘要", index=False)
            QMessageBox.information(self, "导出完成", f"部门扩充结果已保存：\n{target}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    # Legacy preview retained for backward compatibility with external callers;
    # the UI now uses _build_change_preview_expanded above.
    def _build_change_preview(self, frame: pd.DataFrame, cb: Codebook):
        if "w_name" not in frame.columns and "w_name_o" not in frame.columns:
            return pd.DataFrame(), frame.copy(), []
        rows = list(cb.valid_rows())
        def variants(item):
            return {normalize_name(item.w_name), *(normalize_name(alias) for alias in item.aliases)} - {""}
        scoped_index: dict[tuple[str, str, str], list] = {}
        for item in rows:
            cities = {normalize_name(x) for x in cb.split_values(item.cityname)} - {""}
            dists = {normalize_name(x) for x in cb.split_values(item.distname)} - {""}
            for city in cities:
                for dist in dists:
                    for name in variants(item):
                        scoped_index.setdefault((city, dist, name), []).append(item)
        def candidates(data_row):
            name = normalize_name(self._as_text(data_row.get("w_name")) or self._as_text(data_row.get("w_name_o")))
            city = normalize_name(self._as_text(data_row.get("cityname"))); dist = normalize_name(self._as_text(data_row.get("distname")))
            return scoped_index.get((city, dist, name), [])
        audit_records = []
        changed_input_count = 0
        for idx, data_row in frame.iterrows():
            matches = candidates(data_row)
            if not matches: continue
            changed = [item for item in matches if not item.all_time_empty]
            if not changed:
                expected = self._as_text(data_row.get("wrank_t")); original = self._as_text(data_row.get("wrank20"))
                if original and expected and original != expected:
                    audit_records.append({"source_index": idx, "record_id": data_row.get("record_id", ""), "year": data_row.get("year", ""), "p_name": data_row.get("p_name", ""), "cityname": data_row.get("cityname", ""), "distname": data_row.get("distname", ""), "w_name": data_row.get("w_name", "") or data_row.get("w_name_o", ""), "issue_type": "级别没变但是wrank不统一", "wrank20": original, "wrank_t": expected})
            else:
                changed_input_count += 1
                year = self._as_int(data_row.get("year"))
                if year is None:
                    audit_records.append({"source_index": idx, "record_id": data_row.get("record_id", ""), "year": "", "p_name": data_row.get("p_name", ""), "cityname": data_row.get("cityname", ""), "distname": data_row.get("distname", ""), "w_name": data_row.get("w_name", "") or data_row.get("w_name_o", ""), "issue_type": "w_name有效期无法判断", "reason": "缺少 year，无法判断名称有效期"})
                else:
                    key = self._month_key(year, self._as_int(data_row.get("stmo")) or 6)
                    active = [item for item in changed if item.start_month <= key <= item.end_month]
                    if not active:
                        periods = "；".join(self._display_period({"year_valid_from": item.year_valid_from, "month_valid_from": item.month_valid_from, "year_valid_to": item.year_valid_to, "month_valid_to": item.month_valid_to}) for item in changed)
                        audit_records.append({"source_index": idx, "record_id": data_row.get("record_id", ""), "year": year, "p_name": data_row.get("p_name", ""), "cityname": data_row.get("cityname", ""), "distname": data_row.get("distname", ""), "w_name": data_row.get("w_name", "") or data_row.get("w_name_o", ""), "issue_type": "w_name超出有效期", "reason": f"该名称仅在以下有效期内使用：{periods}"})
        self._change_input_count = changed_input_count
        changed_frame = frame.copy()
        changed_frame["_change_action"] = ""
        changed_frame["_source_index"] = changed_frame.index
        # Legacy callers receive the audit and transformed frame only.  The
        # former random demonstration construction duplicated grouping and
        # copied rows without affecting any audit decision.
        return pd.DataFrame(audit_records), changed_frame, []
        demos = []
        group_cols = [column for column in ["p_name", "cityname", "distname", "w_name", "w_name_o"] if column in frame.columns]
        if not group_cols: group_cols = ["_source_index"]
        for _, group in frame.groupby(group_cols, sort=False, dropna=False):
            first = group.iloc[0]; matches = candidates(first); changed = sorted([item for item in matches if not item.all_time_empty], key=lambda item: item.start_month)
            if not changed: continue
            # A month node is a genuine transition only when a preceding
            # codebook segment exists.  A department whose first known row
            # starts in (say) 1996-08 is not itself a before/after split.
            nodes = sorted({item.start_month for item in changed if item.month_valid_from is not None and any(previous.start_month < item.start_month for previous in changed)})
            transitions = []
            for node in nodes:
                before = [item for item in changed if item.start_month < node]
                after = [item for item in changed if item.start_month >= node]
                if before and after:
                    transitions.append((node, max(before, key=lambda item: item.start_month), min(after, key=lambda item: item.start_month)))
            sample_records = []
            for source_idx, data_row in group.iterrows():
                year = self._as_int(data_row.get("year")); start = self._month_key(data_row.get("styr"), data_row.get("stmo"), 1); end = self._month_key(data_row.get("fiyr"), data_row.get("fimo"), 12)
                crossing = [(node, before_row, after_row) for node, before_row, after_row in transitions if year is not None and start is not None and end is not None and start <= node <= end and year == node // 12]
                def make_copy(selected, action):
                    item = data_row.copy(); item["w_name"] = selected.w_name; item["s1_sector"] = selected.s1_sector; item["wrank_t"] = selected.wrank; item["_change_action"] = action
                    if "post_rank" in data_row.index:
                        adjustment = adjust_post_rank(data_row.get("post_rank"), data_row.get("wrank_t"), selected.wrank)
                        item["post_rank_original"] = data_row.get("post_rank", "")
                        item["post_rank_adjusted"] = adjustment.adjusted
                        item["post_rank_adjustment_reason"] = adjustment.reason
                    return item
                if crossing:
                    # A single annual record can theoretically contain more
                    # than one month node.  Emit one segment between each
                    # adjacent node so the preview never loses a transition.
                    first_node, first_before, _ = crossing[0]
                    pre = make_copy(first_before, "复制行：变化前")
                    pre["fiyr"], pre["fimo"] = first_node // 12, first_node % 12 + 1
                    sample_records.append(pre)
                    for transition_index, (node, _, after_row) in enumerate(crossing):
                        segment = make_copy(after_row, "复制行：变化后")
                        segment["styr"], segment["stmo"] = node // 12, node % 12 + 1
                        if transition_index + 1 < len(crossing):
                            next_node = crossing[transition_index + 1][0]
                            segment["fiyr"], segment["fimo"] = next_node // 12, next_node % 12 + 1
                        else:
                            segment["fiyr"], segment["fimo"] = data_row.get("fiyr"), data_row.get("fimo")
                        sample_records.append(segment)
                    continue
                if not transitions and year is not None:
                    year_key = self._month_key(year, 6)
                    eligible = [item for item in changed if item.start_month <= year_key <= item.end_month]
                    selected = min(eligible, key=lambda item: abs(item.start_month - year_key)) if eligible else changed[-1]
                else:
                    year_key = self._month_key(year, 6) if year is not None else None
                    eligible = [item for item in changed if year_key is not None and item.start_month <= year_key]
                    selected = max(eligible, key=lambda item: item.start_month) if eligible else changed[0]
                item = make_copy(selected, "按年份/月自动填充")
                if transitions and year_key is not None:
                    previous_nodes = [node for node, _, _ in transitions if node <= year_key]
                    next_nodes = [node for node, _, _ in transitions if node > year_key]
                    if next_nodes:
                        next_node = min(next_nodes)
                        if end is None or end >= next_node:
                            item["fiyr"], item["fimo"] = next_node // 12, next_node % 12 + 1
                            item["_change_action"] = "变化前：结束边界设为变动月份"
                    if previous_nodes:
                        previous_node = max(previous_nodes)
                        if start is None or start <= previous_node:
                            item["styr"], item["stmo"] = previous_node // 12, previous_node % 12 + 1
                            item["_change_action"] = "变化后：开始边界设为变动月份"
                sample_records.append(item)
            if sample_records:
                result = pd.DataFrame(sample_records)
                demo_cols = [column for column in ["year", "p_name", "w_name_o", "w_name", "s1_sector", "wrank20", "wrank_t", "post_rank", "post_rank_original", "post_rank_adjusted", "post_rank_adjustment_reason", "styr", "stmo", "fiyr", "fimo", "_change_action"] if column in result.columns]
                description = "有明确月份节点，跨越节点的年度行已复制，并将变化前/后的任职边界都设为该月份。" if transitions else "只有年份边界，按 year 直接填充，不复制年度行。"
                demos.append({"w_name": str(first.get("w_name", "") or first.get("w_name_o", "")), "p_name": str(first.get("p_name", "") or "（未记录姓名）"), "cityname": first.get("cityname", ""), "distname": first.get("distname", ""), "description": description, "frame": result[demo_cols]})
        return pd.DataFrame(audit_records), changed_frame, demos

    def _scope_pairs(self) -> set[tuple[str, str]]:
        """Return exact city/distname pairs represented in the codebook.

        Human review is intentionally limited to these pairs.  In particular,
        a Guangzhou city-level codebook row must not silently be used for a
        Yuexiu/Tianhe/etc. county-level input row.
        """
        path = Path(self.codebook.text())
        key = (str(path.resolve()), path.stat().st_mtime_ns if path.exists() else 0)
        if getattr(self, "_scope_cache_key", None) == key:
            return self._scope_cache_pairs
        pairs: set[tuple[str, str]] = set()
        try:
            cb = Codebook(path)
            for row in cb.valid_rows():
                cities = cb.split_values(row.cityname)
                dists = cb.split_values(row.distname)
                for city in cities:
                    for dist in dists:
                        city_norm, dist_norm = normalize_name(city), normalize_name(dist)
                        if city_norm and dist_norm: pairs.add((city_norm, dist_norm))
        except Exception:
            pairs = set()
        self._scope_cache_key = key; self._scope_cache_pairs = pairs
        return pairs

    def _apply_review_scope(self, frame: pd.DataFrame) -> pd.DataFrame:
        if "scope_match" in frame.columns:
            cached = frame.copy()
            cached["_scope_match"] = cached["scope_match"].fillna(False).astype(bool)
            return cached
        cached = getattr(self, "_scoped_review_frame", None)
        if cached is not None and getattr(self, "_scoped_review_source", None) is frame:
            return cached
        frame = frame.copy()
        try:
            cb = Codebook(Path(self.codebook.text()))
        except Exception:
            cb = None
        if cb is None or "cityname" not in frame.columns or "distname" not in frame.columns:
            frame["_scope_match"] = False
        else:
            frame["_scope_match"] = frame.apply(lambda row: self._row_matches_codebook_scope(row, cb), axis=1)
        self._scoped_review_source = self.debug_frame if hasattr(self, "debug_frame") else None
        self._scoped_review_frame = frame
        return frame

    def open_human_review(self) -> None:
        if not hasattr(self, "debug_frame"):
            QMessageBox.information(self, "尚未加载", "请先运行匹配，或打开已有审核队列。"); return
        if self.source_path is None or not self.source_path.exists():
            source, _ = QFileDialog.getOpenFileName(self, "选择原始输入 Excel", "", "Excel (*.xlsx *.xlsm)")
            if not source: return
            self.source_path = Path(source)
        scoped = self._apply_review_scope(self.debug_frame)
        mask = scoped["issue_type"].astype(str).str.strip().ne("") if "issue_type" in scoped.columns else scoped["decision"].isin({"HUMAN", "MULTIPLE"})
        excluded = int((mask & ~scoped["_scope_match"]).sum())
        review = scoped.loc[mask & scoped["_scope_match"]].copy()
        if excluded:
            self.status.setText(f"已排除 {excluded:,} 条 cityname/distname 不在 codebook 范围内的记录；这些记录不会进入人工审核。")
            self.log.appendPlainText(self.status.text())
        if review.empty:
            QMessageBox.information(self, "没有可审核记录", "当前审核队列中没有同时满足 cityname 和 distname 与 codebook 精确匹配的记录。")
            return
        # Group exact raw names within the same geographic context. This keeps
        # one decision per repeated name while avoiding cross-city spillover.
        if "review_group_key" in review.columns:
            group_columns = ["review_group_key"]
        else:
            group_columns = [column for column in ["raw_name", "cityname", "distname", "existing_w_name", "existing_wrank_t", "existing_s1_sector", "issue_type", "issue_fields"] if column in review.columns]
        if "candidate_ids" in review.columns:
            review["_review_candidate_key"] = review["candidate_ids"].map(lambda value: json.dumps(value, ensure_ascii=False, sort_keys=True, default=str))
            group_columns.append("_review_candidate_key")
        grouped_rows = []; grouped_positions = []
        for _, group in review.groupby(group_columns, sort=False, dropna=False):
            first = group.iloc[0].copy(); first["review_group_count"] = len(group); first["_group_record_ids"] = [str(value) for value in group["record_id"].tolist()] if "record_id" in group.columns else [""] * len(group); grouped_rows.append(first); grouped_positions.append([int(index) for index in group.index])
        grouped = pd.DataFrame(grouped_rows).reset_index(drop=True)
        dialog = HumanReviewDialog(grouped, self.source_path, Path(self.codebook.text()), self, source_positions=grouped_positions, existing_decisions=self.pending_decisions)
        dialog.exec()
        if dialog.pending_decisions:
            self._merge_pending_decisions(dialog.pending_decisions)
            self._persist_pending_batch()
            self._refresh_batch_view()
        self._refresh_review()

    @staticmethod
    def _pending_key(decision: dict) -> tuple[str, ...]:
        return tuple(str(value) for value in decision.get("record_ids", []))

    def _merge_pending_decisions(self, decisions: list[dict]) -> None:
        merged = {self._pending_key(item): item for item in self.pending_decisions}
        for item in decisions: merged[self._pending_key(item)] = item
        self.pending_decisions = list(merged.values())

    def edit_pending_batch(self) -> None:
        if not self.pending_decisions:
            QMessageBox.information(self, "没有暂存审核", "当前没有可编辑的暂存记录。"); return
        self.open_human_review()

    def delete_selected_pending(self) -> None:
        selected = sorted({index.row() for index in self.batch_table.selectionModel().selectedRows()}, reverse=True)
        if not selected:
            QMessageBox.information(self, "请选择暂存记录", "请先在下方暂存表中点击要删除的记录。")
            return
        for index in selected:
            if 0 <= index < len(self.pending_decisions): self.pending_decisions.pop(index)
        self._persist_pending_batch()
        self._refresh_batch_view()

    def _refresh_batch_view(self) -> None:
        rows = self.pending_decisions
        self.batch_count.setText(f"已暂存 {len(rows):,} 个名称组；点击‘保存本批复核结果’后才会写入 codebook 和原始数据。")
        columns = ["raw_name", "action", "w_name", "wrank_t", "s1_sector", "entity_id", "note"]
        self.batch_table.setRowCount(len(rows)); self.batch_table.setColumnCount(len(columns)); self.batch_table.setHorizontalHeaderLabels(columns)
        for i, row in enumerate(rows):
            values = row.get("values", {})
            display = {"raw_name": row.get("raw_name", ""), "action": row.get("action", ""), "w_name": values.get("w_name", ""), "wrank_t": values.get("wrank_t", ""), "s1_sector": values.get("s1_sector", ""), "entity_id": row.get("entity_id", ""), "note": row.get("note", "")}
            for j, col in enumerate(columns): self.batch_table.setItem(i, j, QTableWidgetItem(str(display.get(col, ""))))
        if len(rows) <= 80:
            self.batch_table.resizeColumnsToContents()

    def _set_save_progress(self, value: int, message: str) -> None:
        if not hasattr(self, "save_progress"):
            return
        self.save_progress.setVisible(True)
        self.save_progress.setValue(max(0, min(100, int(value))))
        self.save_progress.setFormat(f"{message}　{int(value)}%")
        QApplication.processEvents()

    def clear_pending_batch(self) -> None:
        if not self.pending_decisions:
            return
        answer = QMessageBox.question(self, "清空暂存", "清空后本批人工选择无法恢复，确定继续吗？")
        if answer == QMessageBox.StandardButton.Yes:
            self.pending_decisions.clear(); self._persist_pending_batch(); self._refresh_batch_view()

    @staticmethod
    def _backup_once(path: Path) -> None:
        backup = path.with_suffix(path.suffix + ".before_human_batch.bak")
        if not backup.exists(): shutil.copy2(path, backup)

    def _apply_source_updates(self, path: Path, decisions: list[dict], apply_post_rank: bool = True) -> list[dict]:
        book = load_workbook(path)
        sheet = book[book.sheetnames[0]]
        headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
        adjustment_logs: list[dict] = []
        record_rows: dict[str, int] = {}
        if "record_id" in headers:
            for excel_row in range(2, sheet.max_row + 1):
                value = sheet.cell(excel_row, headers["record_id"]).value
                if value is not None and str(value).strip(): record_rows[str(value)] = excel_row
        for field in ("w_name", "wrank_t", "s1_sector"):
            if field not in headers:
                headers[field] = sheet.max_column + 1; sheet.cell(1, headers[field]).value = field
        # Keep an audit trail beside the adjusted value.  The original
        # post_rank is copied only once, so rerunning a batch never compounds
        # the same adjustment.
        if apply_post_rank and "post_rank" in headers:
            for field in ("post_rank_original", "post_rank_adjustment_reason", "post_rank_adjustment_status"):
                if field not in headers:
                    headers[field] = sheet.max_column + 1; sheet.cell(1, headers[field]).value = field
        for decision in decisions:
            rows = list(decision.get("source_rows", []))
            if not rows and record_rows:
                rows = [record_rows[str(value)] for value in decision.get("record_ids", []) if str(value) in record_rows]
            for excel_row in rows:
                if not isinstance(excel_row, int) or excel_row < 2 or excel_row > sheet.max_row:
                    raise ValueError(f"原始数据行号无效：{excel_row}")
                old_wrank = sheet.cell(excel_row, headers["wrank_t"]).value if "wrank_t" in headers else ""
                old_post = sheet.cell(excel_row, headers["post_rank"]).value if "post_rank" in headers else ""
                new_wrank = decision.get("values", {}).get("wrank_t", old_wrank)
                adjustment = adjust_post_rank(old_post, old_wrank, new_wrank) if apply_post_rank and "post_rank" in headers else None
                if adjustment is not None and adjustment.delta != 0:
                    if not sheet.cell(excel_row, headers["post_rank_original"]).value:
                        sheet.cell(excel_row, headers["post_rank_original"]).value = old_post
                    sheet.cell(excel_row, headers["post_rank_adjustment_reason"]).value = adjustment.reason
                    sheet.cell(excel_row, headers["post_rank_adjustment_status"]).value = adjustment.status
                    if adjustment.status == "adjusted":
                        sheet.cell(excel_row, headers["post_rank"]).value = adjustment.adjusted
                    adjustment_logs.append({
                        "source_row": excel_row,
                        "record_id": decision.get("record_ids", [""])[0] if decision.get("record_ids") else "",
                        "p_name": decision.get("p_name", ""),
                        "w_name": decision.get("raw_name", ""),
                        "old_wrank_t": old_wrank,
                        "new_wrank_t": new_wrank,
                        "post_rank_original": old_post,
                        "post_rank_adjusted": adjustment.adjusted,
                        "delta": adjustment.delta,
                        "status": adjustment.status,
                        "reason": adjustment.reason,
                    })
                for field, value in decision.get("values", {}).items():
                    if field in headers: sheet.cell(excel_row, headers[field]).value = value
        book.save(path)
        return adjustment_logs

    def _apply_temporal_postprocess(self, path: Path, codebook_path: Path) -> tuple[list[dict], list[dict]]:
        """Audit validity and rank changes without mutating the workbook.

        Temporal changes and post_rank changes are proposals only.  They are
        written to the review logs/queues and must be approved by the human
        reviewer before a separate, explicit write operation is introduced.
        """
        book = load_workbook(path)
        sheet = book[book.sheetnames[0]]
        headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
        if "w_name" not in headers:
            return [], []
        cb = Codebook(codebook_path)
        scoped: dict[tuple[str, str, str], list] = {}
        for item in cb.valid_rows():
            cities = {normalize_name(value) for value in cb.split_values(item.cityname)} - {""}
            dists = {normalize_name(value) for value in cb.split_values(item.distname)} - {""}
            names = {normalize_name(item.w_name), *(normalize_name(alias) for alias in item.aliases)} - {""}
            for city in cities:
                for dist in dists:
                    for name in names:
                        scoped.setdefault((city, dist, name), []).append(item)

        adjustment_logs: list[dict] = []
        issues: list[dict] = []

        for excel_row in range(2, sheet.max_row + 1):
            city = normalize_name(sheet.cell(excel_row, headers.get("cityname", 0)).value if headers.get("cityname") else "")
            dist = normalize_name(sheet.cell(excel_row, headers.get("distname", 0)).value if headers.get("distname") else "")
            scope_row = pd.Series({"cityname": city, "distname": dist})
            if not self._row_matches_codebook_scope(scope_row, cb):
                continue
            name_value = sheet.cell(excel_row, headers["w_name"]).value
            name = normalize_name(name_value)
            old_wrank = sheet.cell(excel_row, headers["wrank_t"]).value if "wrank_t" in headers else ""
            proposed_w_name = name_value
            proposed_sector = sheet.cell(excel_row, headers["s1_sector"]).value if "s1_sector" in headers else ""
            proposed_wrank = old_wrank
            matches = scoped.get((city, dist, name), []) if city and dist and name else []
            changed = [item for item in matches if not item.all_time_empty]
            if changed:
                year = self._as_int(sheet.cell(excel_row, headers["year"]).value) if "year" in headers else None
                month = self._as_int(sheet.cell(excel_row, headers["stmo"]).value) if "stmo" in headers else 6
                if year is None:
                    issues.append({"source_row": excel_row, "record_id": sheet.cell(excel_row, headers["record_id"]).value if "record_id" in headers else "", "cityname": sheet.cell(excel_row, headers.get("cityname", 0)).value if headers.get("cityname") else "", "distname": sheet.cell(excel_row, headers.get("distname", 0)).value if headers.get("distname") else "", "w_name": name_value, "year": "", "issue_type": "w_name有效期无法判断", "reason": "缺少 year，无法判断部门名称是否处于有效期内", "action": "人工审阅后决定是否修改"})
                else:
                    key = self._month_key(year, month or 6)
                    active = [item for item in changed if item.start_month <= key <= item.end_month]
                    if not active:
                        periods = "；".join(self._display_period({"year_valid_from": item.year_valid_from, "month_valid_from": item.month_valid_from, "year_valid_to": item.year_valid_to, "month_valid_to": item.month_valid_to}) for item in changed)
                        issues.append({"source_row": excel_row, "record_id": sheet.cell(excel_row, headers["record_id"]).value if "record_id" in headers else "", "cityname": sheet.cell(excel_row, headers.get("cityname", 0)).value if headers.get("cityname") else "", "distname": sheet.cell(excel_row, headers.get("distname", 0)).value if headers.get("distname") else "", "w_name": name_value, "year": year, "issue_type": "w_name超出有效期", "reason": f"该名称仅在以下有效期内使用：{periods}", "action": "人工审阅后决定是否修改"})
                    else:
                        signatures = {(item.w_name, item.s1_sector, item.wrank) for item in active}
                        if len(signatures) > 1:
                            issues.append({"source_row": excel_row, "record_id": sheet.cell(excel_row, headers["record_id"]).value if "record_id" in headers else "", "cityname": sheet.cell(excel_row, headers.get("cityname", 0)).value if headers.get("cityname") else "", "distname": sheet.cell(excel_row, headers.get("distname", 0)).value if headers.get("distname") else "", "w_name": name_value, "year": year, "issue_type": "多种可能", "reason": "同一名称和年份对应多个不同的 codebook 编码", "action": "人工审阅后决定是否修改"})
                        else:
                            selected = active[0]
                            proposed_w_name = selected.w_name
                            proposed_sector = selected.s1_sector
                            proposed_wrank = selected.wrank
                            if (normalize_name(str(name_value or "")) != normalize_name(selected.w_name)
                                    or str(proposed_sector or "") != str(sheet.cell(excel_row, headers["s1_sector"]).value if "s1_sector" in headers else "")
                                    or str(proposed_wrank or "") != str(old_wrank or "")):
                                issues.append({"source_row": excel_row, "record_id": sheet.cell(excel_row, headers["record_id"]).value if "record_id" in headers else "", "cityname": sheet.cell(excel_row, headers.get("cityname", 0)).value if headers.get("cityname") else "", "distname": sheet.cell(excel_row, headers.get("distname", 0)).value if headers.get("distname") else "", "w_name": name_value, "year": year, "issue_type": "变动建议", "reason": "codebook 有唯一有效记录，但不会自动写回", "proposed_w_name": proposed_w_name, "proposed_s1_sector": proposed_sector, "proposed_wrank_t": proposed_wrank, "action": "人工审阅后决定是否修改"})

            if "post_rank" not in headers:
                continue
            final_wrank = proposed_wrank
            baseline_wrank = sheet.cell(excel_row, headers["wrank20"]).value if "wrank20" in headers and sheet.cell(excel_row, headers["wrank20"]).value not in (None, "") else old_wrank
            previous_post = sheet.cell(excel_row, headers["post_rank"]).value
            adjustment = adjust_post_rank(previous_post, baseline_wrank, final_wrank)
            if adjustment.delta:
                adjustment_logs.append({"source_row": excel_row, "record_id": sheet.cell(excel_row, headers["record_id"]).value if "record_id" in headers else "", "p_name": sheet.cell(excel_row, headers["p_name"]).value if "p_name" in headers else "", "w_name": sheet.cell(excel_row, headers["w_name"]).value, "old_wrank_t": baseline_wrank, "new_wrank_t": final_wrank, "post_rank_original": previous_post, "post_rank_adjusted": adjustment.adjusted, "delta": adjustment.delta, "status": adjustment.status, "reason": adjustment.reason})
        # Audit-only stage: deliberately do not save the workbook.
        return adjustment_logs, issues

    def save_pending_batch(self) -> None:
        if not self.pending_decisions:
            QMessageBox.information(self, "没有待保存修改", "请先打开候选卡片并完成至少一条审核。")
            return
        if self.source_path is None or not self.source_path.exists():
            QMessageBox.warning(self, "缺少原始数据", "当前批次没有可写入的原始 Excel 文件。")
            return
        codebook_path = Path(self.codebook.text())
        if not codebook_path.exists():
            QMessageBox.warning(self, "缺少 codebook", "当前 codebook 路径不存在。")
            return
        decisions = list(self.pending_decisions)
        all_decisions = list(decisions)
        self.save_batch_button.setEnabled(False)
        self._set_save_progress(0, "准备保存")
        try:
            # Validate all aliases and manual standards before touching either file.
            self._set_save_progress(5, "读取 codebook")
            cb = Codebook(codebook_path)
            by_entity = {row.entity_id: row for row in cb.rows}
            # Alias ownership is department-level, not validity-segment-level.
            # The same logical department may have several codebook rows for
            # different periods, and those rows are allowed to share aliases.
            aliases_in_batch: dict[str, str] = {}
            manual_norms: set[tuple[str, str]] = set()
            blocked_aliases: dict[tuple[str, ...], str] = {}
            total_decisions = max(1, len(decisions))
            for decision_index, decision in enumerate(decisions, start=1):
                decision_key = self._pending_key(decision)
                raw = str(decision.get("raw_name", "")).strip(); raw_norm = normalize_name(raw)
                entity_id = str(decision.get("entity_id", "") or "")
                values = decision.get("values", {})
                decision_scope = pd.Series({"cityname": decision.get("cityname", ""), "distname": decision.get("distname", "")})
                if not self._row_matches_codebook_scope(decision_scope, cb):
                    raise ValueError(
                        f"暂存记录超出审核范围：cityname={decision.get('cityname', '') or '空'}、"
                        f"distname={decision.get('distname', '') or '空'} 不在当前 codebook 的精确范围内。"
                        "请点击‘编辑选中的暂存审核’删除或跳过这条记录。"
                    )
                if decision.get("manual_new"):
                    key = (normalize_name(str(decision.get("cityname", ""))), normalize_name(str(values.get("w_name", ""))))
                    if key in manual_norms or any(normalize_name(row.w_name) == key[1] and (not key[0] or key[0] in {normalize_name(x) for x in cb.split_values(row.cityname)}) for row in cb.valid_rows()):
                        raise ValueError(f"标准名称冲突：‘{values.get('w_name', '')}’ 已存在于当前城市 codebook")
                    manual_norms.add(key)
                    for other in cb.rows:
                        other_cities = {normalize_name(x) for x in cb.split_values(other.cityname)}
                        if (not key[0] or not other_cities or key[0] in other_cities) and (raw_norm == other.norm_w_name or raw_norm in {normalize_name(a) for a in other.aliases}):
                            blocked_aliases[decision_key] = f"alias 冲突：‘{raw}’ 已属于 ‘{other.w_name}’，不能作为新标准的 alias"
                            break
                elif entity_id and entity_id not in by_entity:
                    raise ValueError(f"候选部门不存在：{entity_id}")
                if entity_id and raw_norm:
                    selected = by_entity.get(entity_id)
                    if selected and raw_norm != selected.norm_w_name:
                        if raw_norm in aliases_in_batch and aliases_in_batch[raw_norm] != selected.dept_key:
                            blocked_aliases[decision_key] = f"alias 冲突：‘{raw}’ 在本批被归入多个部门"
                            continue
                        aliases_in_batch[raw_norm] = selected.dept_key
                        selected_cities = {normalize_name(x) for x in cb.split_values(selected.cityname)}
                        for other in cb.rows:
                            if other.dept_key == selected.dept_key: continue
                            other_cities = {normalize_name(x) for x in cb.split_values(other.cityname)}
                            if (not selected_cities or not other_cities or selected_cities & other_cities) and (raw_norm == other.norm_w_name or raw_norm in {normalize_name(a) for a in other.aliases}):
                                blocked_aliases[decision_key] = (
                                    f"alias 冲突：‘{raw}’ 已属于 ‘{other.w_name}’；"
                                    f"冲突 codebook 范围为 cityname={other.cityname or '空'}、distname={other.distname or '空'}，"
                                    f"本次选择为 ‘{selected.w_name}’（cityname={selected.cityname or '空'}、distname={selected.distname or '空'}）。"
                                    "请点击‘编辑选中的暂存审核’修改这一条，或清空本批后重新审核。"
                                )
                                break
                        if decision_key in blocked_aliases:
                            continue
                if decision_index == total_decisions or decision_index % max(1, total_decisions // 10) == 0:
                    self._set_save_progress(5 + int(decision_index / total_decisions * 20), "检查 alias 冲突")
            if blocked_aliases:
                decisions = [decision for decision in decisions if self._pending_key(decision) not in blocked_aliases]
                self.pending_decisions = all_decisions
                self._persist_pending_batch()
                messages = "\n".join(f"• {message}" for message in blocked_aliases.values())
                if not decisions:
                    QMessageBox.warning(self, "本批暂存均存在 alias 冲突", f"没有安全可保存的记录。冲突记录仍保留在暂存文件中：\n{messages}")
                    return
                QMessageBox.warning(self, "部分记录存在 alias 冲突", f"将跳过 {len(blocked_aliases)} 条冲突记录，先保存其余 {len(decisions)} 条。冲突记录仍保留在暂存文件中，可稍后编辑：\n{messages}")
            answer = QMessageBox.question(self, "保存本批复核结果", f"将一次性写入 {len(decisions)} 个名称组，并更新原始 Excel 与 codebook。已自动创建备份，确定保存吗？")
            if answer != QMessageBox.StandardButton.Yes:
                self._set_save_progress(0, "已取消保存")
                return
            self._set_save_progress(30, "创建阶段文件")
            reviewed_output = self.source_path.with_name(self.source_path.stem + "_人工复核完成" + self.source_path.suffix)
            shutil.copy2(self.source_path, reviewed_output)
            self._backup_once(reviewed_output); self._backup_once(codebook_path)
            book = load_workbook(codebook_path); sheet = book[book.sheetnames[0]]
            headers = {str(sheet.cell(1, c).value): c for c in range(1, sheet.max_column + 1) if sheet.cell(1, c).value}
            alias_col = headers.get("alias")
            if alias_col is None: raise ValueError("codebook 缺少 alias 列")
            department_rows: dict[str, list[int]] = {}
            for row in cb.rows:
                department_rows.setdefault(row.dept_key, []).append(row.source_row)
            for decision in decisions:
                values = decision.get("values", {}); entity_id = str(decision.get("entity_id", "") or "")
                if decision.get("manual_new"):
                    required = ["cityname", "distname", "dtype", "drank", "s1_sector", "wrank", "w_name", "alias", "notes", "year_valid_from", "month_valid_from", "year_valid_to", "month_valid_to"]
                    missing = [field for field in required if field not in headers]
                    if missing: raise ValueError(f"codebook 字段不完整，缺少：{missing}")
                    target = sheet.max_row + 1
                    new_values = {"cityname": decision.get("cityname", ""), "distname": decision.get("distname", ""), "dtype": decision.get("dtype", "待核实"), "drank": decision.get("drank", "待核实"), "s1_sector": values.get("s1_sector", ""), "wrank": values.get("wrank_t", ""), "w_name": values.get("w_name", ""), "alias": decision.get("raw_name", ""), "notes": "人工复核新增；请后续审计"}
                    for field in ["year_valid_from", "month_valid_from", "year_valid_to", "month_valid_to"]: new_values[field] = ""
                    for field, value in new_values.items(): sheet.cell(target, headers[field]).value = value
                elif entity_id and str(decision.get("raw_name", "")).strip() and normalize_name(str(decision.get("raw_name", ""))) != normalize_name(str(values.get("w_name", ""))):
                    selected = by_entity.get(entity_id)
                    targets = department_rows.get(selected.dept_key, []) if selected else []
                    raw = str(decision.get("raw_name", "")).strip()
                    for target in targets:
                        existing = split_aliases(sheet.cell(target, alias_col).value)
                        if normalize_name(raw) not in {normalize_name(a) for a in existing}:
                            sheet.cell(target, alias_col).value = "；".join(existing + [raw])
            book.save(codebook_path)
            self._set_save_progress(45, "已更新 codebook")
            # Stage 1: write human decisions only.  Temporal validity and
            # post_rank are intentionally applied afterwards as separate
            # deterministic stages.
            self._apply_source_updates(reviewed_output, decisions, apply_post_rank=False)
            self._set_save_progress(60, "已生成人工复核完成文件")
            # The expensive temporal expansion belongs to the explicit
            # change-review stage.  Running it here blocked the GUI while
            # saving and also caused the same workbook to be scanned twice.
            post_rank_logs, temporal_issues = [], []
            self._set_save_progress(80, "阶段文件已保存，等待变动审阅")
            if self.result and self.result.get("output_dir") and post_rank_logs:
                log_path = Path(self.result["output_dir"]) / "post_rank_adjustment_log.xlsx"
                pd.DataFrame(post_rank_logs).to_excel(log_path, index=False)
            if self.result and self.result.get("output_dir") and temporal_issues:
                temporal_path = Path(self.result["output_dir"]) / "temporal_review_queue.xlsx"
                pd.DataFrame(temporal_issues).to_excel(temporal_path, index=False)
                review_queue_value = self.result.get("review_queue")
                review_queue = Path(review_queue_value) if review_queue_value else None
                if review_queue is not None and review_queue.exists():
                    existing_queue = pd.read_excel(review_queue, dtype=object)
                    merged_queue = pd.concat([existing_queue, pd.DataFrame(temporal_issues)], ignore_index=True)
                    dedupe = [column for column in ["record_id", "source_row", "issue_type", "w_name", "year"] if column in merged_queue.columns]
                    if dedupe: merged_queue = merged_queue.drop_duplicates(subset=dedupe, keep="first")
                    merged_queue.to_excel(review_queue, index=False)
            feedback_path = Path(__file__).resolve().parents[2] / "data" / "human_feedback.csv"; feedback_path.parent.mkdir(parents=True, exist_ok=True)
            feedback_rows = []
            for decision in decisions:
                for record_id in decision.get("record_ids", [""]):
                    feedback_rows.append({"record_id": record_id, "raw_name": decision.get("raw_name", ""), "cityname": decision.get("cityname", ""), "action": decision.get("action", ""), "entity_id": decision.get("entity_id", ""), "w_name": decision.get("values", {}).get("w_name", ""), "wrank_t": decision.get("values", {}).get("wrank_t", ""), "s1_sector": decision.get("values", {}).get("s1_sector", ""), "human_note": decision.get("note", "")})
            new_feedback = pd.DataFrame(feedback_rows)
            if feedback_path.exists():
                old_feedback = pd.read_csv(feedback_path, dtype=object)
                pd.concat([old_feedback, new_feedback], ignore_index=True).to_csv(feedback_path, index=False)
            else:
                new_feedback.to_csv(feedback_path, index=False)
            if self.result and self.result.get("output_dir"):
                batch_log = Path(self.result["output_dir"]) / "human_review_batch.json"
                batch_log.write_text(json.dumps({"decisions": decisions}, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
                modification_path = Path(self.result["output_dir"]) / "modification_log.xlsx"
                if modification_path.exists():
                    old_log = pd.read_excel(modification_path, dtype=object)
                    human_log = pd.DataFrame([{
                        "record_id": record_id,
                        "raw_name": decision.get("raw_name", ""),
                        "method": "human_batch",
                        "action": decision.get("action", ""),
                        "entity_id": decision.get("entity_id", ""),
                        "new_w_name": decision.get("values", {}).get("w_name", ""),
                        "new_wrank_t": decision.get("values", {}).get("wrank_t", ""),
                        "new_s1_sector": decision.get("values", {}).get("s1_sector", ""),
                        "human_note": decision.get("note", ""),
                    } for decision in decisions for record_id in decision.get("record_ids", [""])])
                    for column in old_log.columns:
                        if column not in human_log.columns: human_log[column] = ""
                    human_log = human_log.reindex(columns=old_log.columns, fill_value="")
                    pd.concat([old_log, human_log], ignore_index=True).to_excel(modification_path, index=False)
            self.pending_decisions = [decision for decision in all_decisions if self._pending_key(decision) in blocked_aliases]
            self._persist_pending_batch(); self._refresh_batch_view(); self._set_save_progress(100, "保存完成（变动待审）"); QMessageBox.information(self, "保存完成", f"人工复核完成文件已生成：\n{reviewed_output}\n\n原始输入文件保持不变。级别、类别、名称有效期及 post_rank 仅生成审核报告，需人工确认后再处理。")
            # The next stage must read the human-reviewed stage file, not the
            # untouched original input.  This makes the workflow continuous:
            # human review -> saved stage file -> change audit.
            self.source_path = reviewed_output
            self.status.setText(f"人工复核已保存：{reviewed_output.name}；现在可进入‘变动审阅’。")
            self.log.appendPlainText(self.status.text())
            self._refresh_checklist()
            self.nav_list.setCurrentRow(3)
        except Exception as exc:
            self._set_save_progress(0, "保存失败")
            QMessageBox.critical(self, "批量保存失败", str(exc))
        finally:
            self.save_batch_button.setEnabled(True)

    def load_existing_review(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "打开人工复核队列", "", "审核队列 (*.xlsx *.jsonl);;All files (*)")
        if not path:
            return
        try:
            p = Path(path)
            sibling_debug = p.parent / "debug" / "queries.jsonl"
            if sibling_debug.exists():
                self.debug_frame = pd.read_json(sibling_debug, lines=True)
            elif p.suffix.lower() == ".xlsx":
                self.debug_frame = pd.read_excel(p)
            else:
                self.debug_frame = pd.read_json(p, lines=True)
            self._scoped_review_frame = None; self._scoped_review_source = None
            self.source_path = None
            self._show_loaded_frame()
            self.status.setText(f"已加载人工复核队列：{int(self.debug_frame['decision'].isin({'HUMAN', 'MULTIPLE'}).sum()) if 'decision' in self.debug_frame.columns else len(self.debug_frame):,} 条")
        except Exception as exc:
            QMessageBox.critical(self, "加载失败", f"无法读取审核队列：{exc}")

    def _show_loaded_frame(self) -> None:
        self._refresh_review()

    def _refresh_review(self) -> None:
        if not hasattr(self, "debug_frame"):
            return
        selected = self.review_filter.currentText()
        scoped = self._apply_review_scope(self.debug_frame)
        excluded = int((scoped.get("issue_type", pd.Series("", index=scoped.index)).astype(str).str.strip().ne("") & ~scoped["_scope_match"]).sum())
        if "issue_type" in scoped.columns:
            base = scoped[scoped["issue_type"].astype(str).str.strip().ne("") & scoped["_scope_match"]]
            frame = base if selected == "全部问题" else base[base["issue_type"] == selected]
            total = len(frame)
        else:
            allowed = {"HUMAN", "MULTIPLE"} if selected == "全部问题" else {selected}
            frame = scoped[scoped["decision"].isin(allowed) & scoped["_scope_match"]]
            total = len(frame)
        frame = frame.head(500)
        columns = ["record_id", "raw_name", "cityname", "year", "review_reason", "issue_type", "issue_fields", "existing_w_name", "existing_wrank_t", "existing_s1_sector", "system_w_name", "system_wrank_t", "system_s1_sector", "decision", "entity_id", "reason"]
        columns = [c for c in columns if c in frame.columns]
        scope_note = f"；另有 {excluded:,} 条因 cityname/distname 不匹配已排除" if excluded else ""
        self.review_count.setText(f"共 {total:,} 条，当前显示前 {len(frame):,} 条{scope_note}")
        self.review_table.setRowCount(len(frame)); self.review_table.setColumnCount(len(columns)); self.review_table.setHorizontalHeaderLabels(columns)
        for i, (_, row) in enumerate(frame.iterrows()):
            for j, col in enumerate(columns): self.review_table.setItem(i, j, QTableWidgetItem(str(row.get(col, ""))))
        if len(frame) <= 80:
            self.review_table.resizeColumnsToContents()

    @Slot(str)
    def on_failed(self, message: str) -> None:
        self.status.setText("运行失败"); self.header_status.setText("● 运行失败"); self.header_status.setStyleSheet("color:#a94442;background:#fbe4e4;border-radius:12px;padding:6px 12px;font-weight:700;"); self.log.appendPlainText(message); QMessageBox.critical(self, "运行失败", message)

    def thread_done(self) -> None:
        self.run_button.setEnabled(True); self.worker = None; self.worker_thread = None


def main() -> int:
    app = QApplication(sys.argv)
    preferred = ["PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Noto Sans CJK SC", "Arial"]
    available = set(QFontDatabase.families())
    family = next((name for name in preferred if name in available), QFontDatabase.systemFont(QFontDatabase.GeneralFont).family())
    app.setFont(QFont(family, 12))
    app.setStyleSheet("""
        QWidget { color: #243447; background: #f6f8fb; }
        QMainWindow { background: #f6f8fb; }
        QWidget#sidebar { background: #173b5b; }
        QLabel#brand { color: #ffffff; font-size: 22px; font-weight: 800; padding: 0 4px; background: transparent; }
        QLabel#navhint, QLabel#version { color: #a8c3d5; font-size: 11px; padding: 0 4px; background: transparent; }
        QListWidget#navlist { background: transparent; border: none; outline: none; color: #d9e9f3; font-size: 14px; }
        QListWidget#navlist::item { padding: 12px 13px; border-radius: 6px; margin: 1px 0; }
        QListWidget#navlist::item:hover { background: #285877; }
        QListWidget#navlist::item:selected { background: #4a9ac3; color: #ffffff; font-weight: 700; }
        QLabel#title { font-size: 25px; font-weight: 700; color: #153b5b; padding: 4px 2px 8px; background: transparent; }
        QLabel#title .subtitle { font-size: 13px; font-weight: 400; color: #6b7d8f; }
        QLabel#hint { color: #617487; padding: 4px 2px; background: transparent; }
        QLabel#stageNotice { color: #7a4b00; background: #fff4d6; border: 1px solid #e5bd67; border-radius: 7px; padding: 10px 12px; font-weight: 700; }
        QLabel#ready { color: #2a8b63; background: #e2f4eb; border-radius: 12px; padding: 6px 12px; font-weight: 700; }
        QTabWidget::pane { border: 1px solid #d9e2ec; border-radius: 8px; background: #ffffff; top: -1px; }
        QTabBar::tab { background: #e9eff5; color: #516579; padding: 10px 18px; margin-right: 3px; border-top-left-radius: 7px; border-top-right-radius: 7px; }
        QTabBar::tab:selected { background: #ffffff; color: #145a86; font-weight: 700; }
        QGroupBox { background: #ffffff; border: 1px solid #d9e2ec; border-radius: 8px; margin-top: 12px; padding: 15px 14px 12px; font-weight: 700; color: #254b69; }
        QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 6px; background: #ffffff; }
        QLineEdit, QComboBox, QSpinBox { background: #ffffff; border: 1px solid #c8d5e1; border-radius: 5px; padding: 7px 9px; min-height: 18px; }
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1px solid #4b9ac6; }
        QPushButton { background: #edf3f8; border: 1px solid #c5d4e1; border-radius: 5px; padding: 8px 15px; min-height: 18px; }
        QPushButton:hover { background: #deedf7; }
        QPushButton#primary { background: #176b9b; color: white; border: 1px solid #176b9b; font-weight: 700; }
        QPushButton#primary:hover { background: #12577f; }
        QPushButton#candidate { text-align: left; background: #ffffff; border: 1px solid #c8d5e1; padding: 10px 12px; min-height: 42px; }
        QPushButton#candidate:hover { background: #eaf5fb; border: 1px solid #4b9ac6; }
        QPushButton#candidate:checked { background: #d6efff; color: #0b4e78; border: 2px solid #1688c2; font-weight: 700; }
        QProgressBar { height: 18px; border: 1px solid #c8d5e1; border-radius: 5px; background: #ffffff; text-align: center; }
        QProgressBar::chunk { background: #42a5d5; border-radius: 4px; }
        QPlainTextEdit { background: #ffffff; border: 1px solid #d9e2ec; border-radius: 5px; padding: 5px; }
        QTableWidget { background: #ffffff; alternate-background-color: #f2f7fb; border: 1px solid #d9e2ec; gridline-color: #e5edf3; selection-background-color: #d7ecfa; selection-color: #163a52; }
        QHeaderView::section { background: #e8f0f6; color: #31536d; border: none; border-right: 1px solid #d5e1ea; border-bottom: 1px solid #c9d8e3; padding: 7px 6px; font-weight: 700; }
    """)
    window = MainWindow(); window.show(); return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
