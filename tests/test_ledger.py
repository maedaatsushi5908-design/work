import sqlite3
import tempfile
import unittest
from pathlib import Path

from ev.ledger import GENESIS_HASH, Ledger, LedgerIntegrityError, canonical_json, compute_hash


class LedgerTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "journal.db"

    def tearDown(self):
        self.tmp.cleanup()


class TestCanonicalJson(LedgerTestCase):
    def test_key_order_does_not_matter(self):
        self.assertEqual(
            canonical_json({"b": 1, "a": 2}), canonical_json({"a": 2, "b": 1})
        )

    def test_japanese_is_not_escaped(self):
        self.assertIn("東京", canonical_json({"x": "東京"}))

    def test_hash_is_deterministic(self):
        a = compute_hash(GENESIS_HASH, "2026-01-01T00:00:00+00:00", "k", {"x": 1})
        b = compute_hash(GENESIS_HASH, "2026-01-01T00:00:00+00:00", "k", {"x": 1})
        self.assertEqual(a, b)
        self.assertEqual(len(a), 64)

    def test_hash_changes_with_payload(self):
        a = compute_hash(GENESIS_HASH, "t", "k", {"x": 1})
        b = compute_hash(GENESIS_HASH, "t", "k", {"x": 2})
        self.assertNotEqual(a, b)


class TestAppendAndRead(LedgerTestCase):
    def test_append_and_length(self):
        with Ledger(self.path) as led:
            self.assertEqual(len(led), 0)
            led.append("forecast", {"ticker": "1234"})
            led.append("forecast", {"ticker": "5678"})
            self.assertEqual(len(led), 2)

    def test_chain_links_correctly(self):
        with Ledger(self.path) as led:
            e1 = led.append("a", {"i": 1})
            e2 = led.append("a", {"i": 2})
            self.assertEqual(e1.prev_hash, GENESIS_HASH)
            self.assertEqual(e2.prev_hash, e1.hash)
            self.assertEqual(led.last_hash(), e2.hash)

    def test_entries_filtered_by_kind(self):
        with Ledger(self.path) as led:
            led.append("forecast", {"i": 1})
            led.append("resolution", {"i": 2})
            led.append("forecast", {"i": 3})
            self.assertEqual(len(list(led.entries("forecast"))), 2)
            self.assertEqual(len(list(led.entries())), 3)

    def test_find_by_payload(self):
        with Ledger(self.path) as led:
            led.append("f", {"ticker": "1234.T", "n": 1})
            led.append("f", {"ticker": "5678.T", "n": 2})
            hits = led.find("f", ticker="5678.T")
            self.assertEqual(len(hits), 1)
            self.assertEqual(hits[0].payload["n"], 2)

    def test_payload_roundtrips(self):
        payload = {"a": 1, "b": [1, 2, 3], "c": {"d": "日本語"}, "e": None, "f": 1.5}
        with Ledger(self.path) as led:
            led.append("k", payload)
        with Ledger(self.path) as led:
            self.assertEqual(list(led.entries())[0].payload, payload)

    def test_persists_across_reopen(self):
        with Ledger(self.path) as led:
            led.append("k", {"x": 1})
            h = led.last_hash()
        with Ledger(self.path) as led:
            self.assertEqual(len(led), 1)
            self.assertEqual(led.last_hash(), h)

    def test_rejects_empty_kind(self):
        with Ledger(self.path) as led:
            with self.assertRaises(ValueError):
                led.append("", {"x": 1})
            with self.assertRaises(ValueError):
                led.append("   ", {"x": 1})

    def test_rejects_unserializable_payload(self):
        with Ledger(self.path) as led:
            with self.assertRaises(ValueError):
                led.append("k", {"bad": object()})

    def test_creates_parent_directory(self):
        nested = Path(self.tmp.name) / "deep" / "sub" / "j.db"
        with Ledger(nested) as led:
            led.append("k", {"x": 1})
        self.assertTrue(nested.exists())


class TestImmutability(LedgerTestCase):
    def test_update_is_blocked_by_trigger(self):
        with Ledger(self.path) as led:
            led.append("k", {"x": 1})
            with self.assertRaises(sqlite3.IntegrityError):
                led._conn.execute("UPDATE entries SET payload = '{}' WHERE seq = 1")

    def test_delete_is_blocked_by_trigger(self):
        with Ledger(self.path) as led:
            led.append("k", {"x": 1})
            with self.assertRaises(sqlite3.IntegrityError):
                led._conn.execute("DELETE FROM entries WHERE seq = 1")

    def test_triggers_survive_reopen(self):
        with Ledger(self.path) as led:
            led.append("k", {"x": 1})
        with Ledger(self.path) as led:
            with self.assertRaises(sqlite3.IntegrityError):
                led._conn.execute("DELETE FROM entries")


class TestVerify(LedgerTestCase):
    def test_empty_ledger_verifies(self):
        with Ledger(self.path) as led:
            ok, msg = led.verify()
            self.assertTrue(ok)
            self.assertIn("0件", msg)

    def test_intact_chain_verifies(self):
        with Ledger(self.path) as led:
            for i in range(10):
                led.append("k", {"i": i})
            ok, msg = led.verify()
            self.assertTrue(ok)
            self.assertIn("10件", msg)

    def test_detects_payload_tampering_after_dropping_triggers(self):
        """
        トリガーを外してDBを直接書き換えても、ハッシュチェーンが検出する。
        これが二重防御の意味。
        """
        with Ledger(self.path) as led:
            led.append("forecast", {"probability": 0.30})
            led.append("forecast", {"probability": 0.40})

        conn = sqlite3.connect(str(self.path))
        conn.execute("DROP TRIGGER entries_no_update")
        conn.execute(
            "UPDATE entries SET payload = ? WHERE seq = 1",
            (canonical_json({"probability": 0.90}),),
        )
        conn.commit()
        conn.close()

        with Ledger(self.path) as led:
            ok, msg = led.verify()
            self.assertFalse(ok)
            self.assertIn("hash 不一致", msg)
            with self.assertRaises(LedgerIntegrityError):
                led.verify_or_raise()

    def test_detects_deletion_after_dropping_triggers(self):
        with Ledger(self.path) as led:
            for i in range(4):
                led.append("k", {"i": i})

        conn = sqlite3.connect(str(self.path))
        conn.execute("DROP TRIGGER entries_no_delete")
        conn.execute("DELETE FROM entries WHERE seq = 2")
        conn.commit()
        conn.close()

        with Ledger(self.path) as led:
            ok, msg = led.verify()
            self.assertFalse(ok)
            self.assertIn("prev_hash 不一致", msg)

    def test_verify_or_raise_passes_when_intact(self):
        with Ledger(self.path) as led:
            led.append("k", {"x": 1})
            led.verify_or_raise()


class TestFixedClock(LedgerTestCase):
    def test_injected_clock_is_used(self):
        with Ledger(self.path, clock=lambda: "2026-01-01T00:00:00+00:00") as led:
            e = led.append("k", {"x": 1})
            self.assertEqual(e.ts, "2026-01-01T00:00:00+00:00")

    def test_same_content_same_hash_with_fixed_clock(self):
        p2 = Path(self.tmp.name) / "b.db"
        clock = lambda: "2026-01-01T00:00:00+00:00"
        with Ledger(self.path, clock=clock) as a, Ledger(p2, clock=clock) as b:
            ea = a.append("k", {"x": 1})
            eb = b.append("k", {"x": 1})
            self.assertEqual(ea.hash, eb.hash)


if __name__ == "__main__":
    unittest.main()
