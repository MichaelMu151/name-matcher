"""Command-line entry point for the staged department matching engine."""

from __future__ import annotations

import argparse
from pathlib import Path

from src.engine.codebook import Codebook
from src.engine.evaluate import evaluate_predictions
from src.engine.feedback import export_review, ingest_feedback
from src.engine.pipeline import Pipeline, _version
from src.engine.ranker import Ranker


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="department-matcher")
    commands = parser.add_subparsers(dest="command", required=True)
    audit = commands.add_parser("audit", help="load a codebook and write M0 audit reports")
    audit.add_argument("--codebook", required=True, help="CSV/XLSX codebook path")
    audit.add_argument("--reports-dir", default="reports", help="audit report directory")
    build = commands.add_parser("build-index", help="build the persisted M2 codebook index snapshot")
    build.add_argument("--codebook", required=True)
    build.add_argument("--output", default="models/codebook_index.csv")
    run = commands.add_parser("run", help="match an input workbook and export two delivery files")
    run.add_argument("--codebook", required=True)
    run.add_argument("--input", required=True)
    run.add_argument("--output-dir", default="outputs")
    run.add_argument("--config", default="config.yaml")
    run.add_argument("--model", default="")
    train = commands.add_parser("train", help="fit the synthetic cold-start ranker")
    train.add_argument("--codebook", required=True)
    train.add_argument("--output", default="models/synthetic_ranker.joblib")
    train.add_argument("--reports-dir", default="reports")
    review = commands.add_parser("review-export", help="make a human review workbook")
    review.add_argument("--input", required=True)
    review.add_argument("--debug", required=True)
    review.add_argument("--output", required=True)
    feedback = commands.add_parser("ingest-feedback", help="append human feedback pending second review")
    feedback.add_argument("--input", required=True)
    feedback.add_argument("--output", default="data/aliases_learned.csv")
    feedback.add_argument("--operator", default="")
    feedback.add_argument("--second-reviewer", default="")
    feedback.add_argument("--batch-id", default="")
    feedback.add_argument("--labels", default="")
    feedback.add_argument("--cases", default="")
    evaluation = commands.add_parser("eval", help="write a small evaluation report")
    evaluation.add_argument("--debug", required=True)
    evaluation.add_argument("--gold", default="")
    evaluation.add_argument("--output", default="reports/eval.md")
    evaluation.add_argument("--final", action="store_true", help="允许读取冻结 test gold 集")
    evaluation.add_argument("--input", default="")
    evaluation.add_argument("--reviewed", default="")
    evaluation.add_argument("--log", default="")
    commands.add_parser("review-ui", help="launch Streamlit review app")
    bench = commands.add_parser("bench", help="run a lightweight pipeline timing benchmark")
    bench.add_argument("--codebook", required=True)
    bench.add_argument("--input", required=True)
    bench.add_argument("--output-dir", default="outputs")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "audit":
        codebook = Codebook(args.codebook, reports_dir=Path(args.reports_dir))
        result = codebook.write_reports()
        print(f"loaded {len(codebook.rows)} rows")
        for key, value in result.summary.items():
            print(f"{key}: {value}")
        return 0
    if args.command == "build-index":
        codebook = Codebook(args.codebook)
        output = Path(args.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        codebook.enriched_frame().to_csv(output, index=False)
        print(output)
        return 0
    if args.command == "run":
        result = Pipeline(args.codebook, args.config, args.model or None).run(args.input, args.output_dir)
        print(result)
        return 0
    if args.command == "train":
        codebook = Codebook(args.codebook, reports_dir=args.reports_dir)
        ranker = Ranker()
        print(ranker.fit_synthetic(codebook))
        ranker.save(args.output, data_hash=_version(args.codebook))
        return 0
    if args.command == "review-export":
        print(export_review(args.input, args.debug, args.output))
        return 0
    if args.command == "ingest-feedback":
        print(ingest_feedback(args.input, args.output, operator=args.operator, second_reviewer=args.second_reviewer, batch_id=args.batch_id or None, labels_path=args.labels or None, cases_path=args.cases or None))
        return 0
    if args.command == "eval":
        print(evaluate_predictions(args.debug, args.gold or None, args.output, final=args.final, input_path=args.input or None, reviewed_path=args.reviewed or None, log_path=args.log or None))
        return 0
    if args.command == "review-ui":
        from src.engine.review_app import run_app
        run_app()
        return 0
    if args.command == "bench":
        import time
        started = time.perf_counter()
        result = Pipeline(args.codebook).run(args.input, args.output_dir)
        result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        print(result)
        return 0
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
