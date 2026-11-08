"""Small durable SQLite cache for deterministic query decisions."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any


class QueryCache:
    def __init__(self, path: str | Path, namespace: str):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(self.path, timeout=30)
        self.connection.execute("PRAGMA journal_mode=WAL")
        self.connection.execute("PRAGMA synchronous=NORMAL")
        self.connection.execute(
            "CREATE TABLE IF NOT EXISTS query_cache (namespace TEXT NOT NULL, query_key TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(namespace, query_key))"
        )
        self.connection.commit()
        self.namespace = namespace

    def get(self, query_key: tuple[Any, ...]) -> tuple[Any, list[dict[str, Any]], list[str]] | None:
        key = json.dumps(query_key, ensure_ascii=False, sort_keys=True, default=str)
        row = self.connection.execute("SELECT payload FROM query_cache WHERE namespace=? AND query_key=?", (self.namespace, key)).fetchone()
        if row is None:
            return None
        from .policy import Decision
        payload = json.loads(row[0])
        return Decision(**payload["decision"]), payload["debug"], payload["flags"]

    def put(self, query_key: tuple[Any, ...], result: tuple[Any, list[dict[str, Any]], list[str]]) -> None:
        decision, debug, flags = result
        key = json.dumps(query_key, ensure_ascii=False, sort_keys=True, default=str)
        payload = json.dumps({"decision": decision.__dict__, "debug": debug, "flags": flags}, ensure_ascii=False, default=str)
        self.connection.execute("INSERT OR REPLACE INTO query_cache(namespace, query_key, payload) VALUES(?,?,?)", (self.namespace, key, payload))
        self.connection.commit()

    def close(self) -> None:
        self.connection.close()

    def __enter__(self) -> "QueryCache":
        return self

    def __exit__(self, *_: Any) -> None:
        self.close()
