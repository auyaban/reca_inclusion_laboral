from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app
from formularios import common


class _FakeResponse:
    def __init__(self, payload: bytes = b"") -> None:
        self._payload = payload
        self.status = 200

    def read(self) -> bytes:
        return self._payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class SupabaseUsageOptimizationsTests(unittest.TestCase):
    def test_get_cached_payload_returns_stale_payload_when_refresh_fails(self) -> None:
        cache_key = "test_cached_payload"
        original_memory_cache = dict(common._LOCAL_PAYLOAD_CACHE)
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                with patch.object(common, "_get_cache_dir", return_value=tmpdir):
                    common._LOCAL_PAYLOAD_CACHE.clear()

                    first = common._get_cached_payload(
                        cache_key,
                        lambda: ["9001", "9002"],
                        ttl_seconds=86400,
                    )
                    self.assertEqual(first, ["9001", "9002"])

                    stale_path = Path(tmpdir) / "payload_cache" / f"{cache_key}.json"
                    self.assertTrue(stale_path.exists())

                    fallback = common._get_cached_payload(
                        cache_key,
                        lambda: (_ for _ in ()).throw(RuntimeError("network down")),
                        ttl_seconds=0,
                        force=True,
                        allow_stale_on_error=True,
                    )
                    self.assertEqual(fallback, ["9001", "9002"])
        finally:
            common._LOCAL_PAYLOAD_CACHE.clear()
            common._LOCAL_PAYLOAD_CACHE.update(original_memory_cache)

    def test_supabase_upsert_uses_return_minimal(self) -> None:
        seen_headers = {}

        def _fake_urlopen(request, timeout=0):
            del timeout
            seen_headers.update({key.lower(): value for key, value in request.header_items()})
            return _FakeResponse()

        with patch.object(common, "_load_supabase_credentials", return_value=("https://example.supabase.co", "key")):
            with patch.object(common, "_supabase_get_access_token", return_value="jwt"):
                with patch("urllib.request.urlopen", side_effect=_fake_urlopen):
                    result = common._supabase_upsert("demo", [{"id": 1}], on_conflict="id")

        self.assertEqual(result, [])
        self.assertEqual(seen_headers.get("prefer"), "resolution=merge-duplicates,return=minimal")

    def test_supabase_patch_uses_return_minimal(self) -> None:
        seen_headers = {}

        def _fake_urlopen(request, timeout=0):
            del timeout
            seen_headers.update({key.lower(): value for key, value in request.header_items()})
            return _FakeResponse()

        with patch.object(common, "_load_supabase_credentials", return_value=("https://example.supabase.co", "key")):
            with patch.object(common, "_supabase_get_access_token", return_value="jwt"):
                with patch("urllib.request.urlopen", side_effect=_fake_urlopen):
                    result = common._supabase_patch("demo", {"id": "abc"}, {"estado": "ok"})

        self.assertEqual(result, [])
        self.assertEqual(seen_headers.get("prefer"), "return=minimal")

    def test_merge_company_catalog_rows_overwrites_by_company_id_and_keeps_latest_sync(self) -> None:
        base_rows = [
            {
                "id": "1",
                "nombre_empresa": "Alpha SAS",
                "nit_empresa": "9001",
                "updated_at": "2026-04-07T10:00:00+00:00",
                "estado": "Activo",
            },
            {
                "id": "2",
                "nombre_empresa": "Beta SAS",
                "nit_empresa": "9002",
                "updated_at": "2026-04-07T11:00:00+00:00",
                "estado": "Activo",
            },
        ]
        delta_rows = [
            {
                "id": "2",
                "nombre_empresa": "Beta SAS",
                "nit_empresa": "9002",
                "updated_at": "2026-04-07T12:30:00+00:00",
                "estado": "Suspendido",
            },
            {
                "id": "3",
                "nombre_empresa": "Gamma SAS",
                "nit_empresa": "9003",
                "updated_at": "2026-04-07T12:45:00+00:00",
                "estado": "Activo",
            },
        ]

        merged = app._merge_company_catalog_rows(base_rows, delta_rows)

        self.assertEqual(len(merged), 3)
        by_id = {row["id"]: row for row in merged}
        self.assertEqual(by_id["2"]["estado"], "Suspendido")
        self.assertEqual(app._company_catalog_last_sync(merged), "2026-04-07T12:45:00+00:00")

    def test_company_name_cache_returns_empty_when_catalog_requires_auth(self) -> None:
        with patch.object(app, "_load_company_search_index", side_effect=RuntimeError("HTTP 401")):
            self.assertEqual(app._get_company_name_cache(), [])
            self.assertEqual(app._get_company_name_suggestions_from_index("em"), [])
            self.assertEqual(app._get_company_nit_suggestions_from_index("90"), [])


if __name__ == "__main__":
    unittest.main()
