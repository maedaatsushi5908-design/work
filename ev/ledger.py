"""
追記専用・改変検知付き台帳

自分の予測を後から書き換えられる記録は、記録として無価値である。
「あのときこう思っていた」は必ず結果に都合よく歪む（後知恵バイアス）。

この台帳は2重に改変を防ぐ:
  1. SQLite のトリガーで UPDATE / DELETE を物理的に禁止する
  2. 各行が直前の行のハッシュを含むハッシュチェーンを構成する
     → DBファイルを直接編集しても verify() で検出できる

保存先はローカルの単一ファイル。外部サービスに依存しない。
"""

from __future__ import annotations

import hashlib
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple, Union

GENESIS_HASH = "0" * 64


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def canonical_json(payload: Dict[str, Any]) -> str:
    """ハッシュ計算用の正規化JSON。キー順を固定し、空白を排除する。"""
    return json.dumps(payload, sort_keys=True, ensure_ascii=False, separators=(",", ":"))


@dataclass(frozen=True)
class Entry:
    seq: int
    ts: str
    kind: str
    payload: Dict[str, Any]
    prev_hash: str
    hash: str


def compute_hash(prev_hash: str, ts: str, kind: str, payload: Dict[str, Any]) -> str:
    material = f"{prev_hash}|{ts}|{kind}|{canonical_json(payload)}"
    return hashlib.sha256(material.encode("utf-8")).hexdigest()


_SCHEMA = """
CREATE TABLE IF NOT EXISTS entries (
    seq       INTEGER PRIMARY KEY AUTOINCREMENT,
    ts        TEXT NOT NULL,
    kind      TEXT NOT NULL,
    payload   TEXT NOT NULL,
    prev_hash TEXT NOT NULL,
    hash      TEXT NOT NULL UNIQUE
);
CREATE INDEX IF NOT EXISTS idx_entries_kind ON entries(kind);

CREATE TRIGGER IF NOT EXISTS entries_no_update
BEFORE UPDATE ON entries
BEGIN
    SELECT RAISE(ABORT, '台帳は追記専用です: 既存レコードの更新は禁止されています');
END;

CREATE TRIGGER IF NOT EXISTS entries_no_delete
BEFORE DELETE ON entries
BEGIN
    SELECT RAISE(ABORT, '台帳は追記専用です: 既存レコードの削除は禁止されています');
END;
"""


class LedgerIntegrityError(Exception):
    """ハッシュチェーンの不整合。DBが改変された可能性がある。"""


class Ledger:
    """
    追記専用台帳。

    使用例:
        with Ledger("journal.db") as led:
            led.append("forecast", {"ticker": "1234.T", "p": 0.7})
            ok, msg = led.verify()
    """

    def __init__(self, path: Union[str, Path], clock: Optional[Callable[[], str]] = None):
        self.path = Path(path)
        self._clock = clock or utc_now_iso
        if self.path.parent and str(self.path.parent) != "":
            self.path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(str(self.path))
        self._conn.row_factory = sqlite3.Row
        self._conn.executescript(_SCHEMA)
        self._conn.commit()

    # ── コンテキストマネージャ ──────────────────────────────

    def __enter__(self) -> "Ledger":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def close(self) -> None:
        self._conn.close()

    # ── 書き込み ─────────────────────────────────────────

    def append(self, kind: str, payload: Dict[str, Any]) -> Entry:
        """
        エントリを追記する。既存行には一切触れない。

        Raises:
            ValueError: kind が空、または payload が JSON 化できない場合
        """
        if not kind or not kind.strip():
            raise ValueError("append: kind が空です")
        try:
            canonical_json(payload)
        except (TypeError, ValueError) as e:
            raise ValueError(f"append: payload が JSON 化できません: {e}") from e

        ts = self._clock()
        prev = self.last_hash()
        h = compute_hash(prev, ts, kind, payload)

        cur = self._conn.execute(
            "INSERT INTO entries (ts, kind, payload, prev_hash, hash) VALUES (?,?,?,?,?)",
            (ts, kind, canonical_json(payload), prev, h),
        )
        self._conn.commit()
        return Entry(
            seq=int(cur.lastrowid),
            ts=ts,
            kind=kind,
            payload=payload,
            prev_hash=prev,
            hash=h,
        )

    # ── 読み出し ─────────────────────────────────────────

    def last_hash(self) -> str:
        row = self._conn.execute(
            "SELECT hash FROM entries ORDER BY seq DESC LIMIT 1"
        ).fetchone()
        return row["hash"] if row else GENESIS_HASH

    def __len__(self) -> int:
        row = self._conn.execute("SELECT COUNT(*) AS c FROM entries").fetchone()
        return int(row["c"])

    def entries(self, kind: Optional[str] = None) -> Iterator[Entry]:
        if kind is None:
            rows = self._conn.execute("SELECT * FROM entries ORDER BY seq")
        else:
            rows = self._conn.execute(
                "SELECT * FROM entries WHERE kind = ? ORDER BY seq", (kind,)
            )
        for r in rows:
            yield Entry(
                seq=int(r["seq"]),
                ts=r["ts"],
                kind=r["kind"],
                payload=json.loads(r["payload"]),
                prev_hash=r["prev_hash"],
                hash=r["hash"],
            )

    def find(self, kind: str, **match: Any) -> List[Entry]:
        """payload のキー一致でエントリを抽出する。"""
        out = []
        for e in self.entries(kind):
            if all(e.payload.get(k) == v for k, v in match.items()):
                out.append(e)
        return out

    # ── 検証 ────────────────────────────────────────────

    def verify(self) -> tuple[bool, str]:
        """
        ハッシュチェーン全体を検証する。

        Returns:
            (整合しているか, 説明メッセージ)
        """
        prev = GENESIS_HASH
        count = 0
        for e in self.entries():
            expected = compute_hash(prev, e.ts, e.kind, e.payload)
            if e.prev_hash != prev:
                return (
                    False,
                    f"seq={e.seq}: prev_hash 不一致（記録={e.prev_hash[:12]}…, "
                    f"期待={prev[:12]}…）。レコードが挿入または削除された可能性があります。",
                )
            if e.hash != expected:
                return (
                    False,
                    f"seq={e.seq}: hash 不一致（記録={e.hash[:12]}…, "
                    f"期待={expected[:12]}…）。payload が改変された可能性があります。",
                )
            prev = e.hash
            count += 1
        return (True, f"整合。{count}件のエントリを検証しました。")

    def verify_or_raise(self) -> None:
        ok, msg = self.verify()
        if not ok:
            raise LedgerIntegrityError(msg)
