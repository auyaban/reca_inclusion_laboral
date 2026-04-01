from __future__ import annotations

import unittest
from unittest.mock import patch

import google_api_requests as google_requests


class _FakeResponse(dict):
    def __init__(self, status, headers=None):
        headers = headers or {}
        super().__init__(headers)
        self.status = status
        self.headers = headers


class _FakeGoogleError(Exception):
    def __init__(self, status, message="boom", headers=None):
        super().__init__(message)
        self.resp = _FakeResponse(status, headers=headers)


class _SequenceRequest:
    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = 0

    def execute(self):
        self.calls += 1
        if not self._responses:
            raise AssertionError("No fake responses left")
        next_item = self._responses.pop(0)
        if isinstance(next_item, Exception):
            raise next_item
        return next_item


class GoogleApiRequestTests(unittest.TestCase):
    def test_execute_google_request_with_retry_retries_after_429(self) -> None:
        request = _SequenceRequest([
            _FakeGoogleError(429, headers={"Retry-After": "0"}),
            {"ok": True},
        ])

        with patch.object(google_requests.time, "sleep") as sleep_mock:
            result = google_requests.execute_google_request_with_retry(
                request,
                operation_name="test.retry_429",
            )

        self.assertEqual(result, {"ok": True})
        self.assertEqual(request.calls, 2)
        sleep_mock.assert_called_once_with(0.0)

    def test_execute_google_request_with_retry_retries_after_5xx(self) -> None:
        request = _SequenceRequest([
            _FakeGoogleError(503),
            {"ok": True},
        ])

        with (
            patch.object(google_requests.random, "random", return_value=0.0),
            patch.object(google_requests.time, "sleep") as sleep_mock,
        ):
            result = google_requests.execute_google_request_with_retry(
                request,
                operation_name="test.retry_503",
            )

        self.assertEqual(result, {"ok": True})
        self.assertEqual(request.calls, 2)
        sleep_mock.assert_called_once_with(1.0)

    def test_execute_google_request_with_retry_does_not_retry_403(self) -> None:
        request = _SequenceRequest([_FakeGoogleError(403)])

        with patch.object(google_requests.time, "sleep") as sleep_mock:
            with self.assertRaises(_FakeGoogleError):
                google_requests.execute_google_request_with_retry(
                    request,
                    operation_name="test.no_retry_403",
                )

        self.assertEqual(request.calls, 1)
        sleep_mock.assert_not_called()

    def test_execute_google_request_with_retry_honors_retry_after_header(self) -> None:
        request = _SequenceRequest([
            _FakeGoogleError(429, headers={"Retry-After": "7"}),
            {"ok": True},
        ])

        with patch.object(google_requests.time, "sleep") as sleep_mock:
            result = google_requests.execute_google_request_with_retry(
                request,
                operation_name="test.retry_after",
            )

        self.assertEqual(result, {"ok": True})
        sleep_mock.assert_called_once_with(7.0)

    def test_execute_google_create_with_confirmation_returns_confirmed_resource_after_ambiguous_error(self) -> None:
        request_count = 0

        def request_factory():
            nonlocal request_count
            request_count += 1
            return _SequenceRequest([_FakeGoogleError(503)])

        confirm_created = lambda: {"id": "confirmed-resource"}

        with patch.object(google_requests.time, "sleep") as sleep_mock:
            result = google_requests.execute_google_create_with_confirmation(
                request_factory,
                confirm_created,
                operation_name="test.confirm_create",
            )

        self.assertEqual(result, {"id": "confirmed-resource"})
        self.assertEqual(request_count, 1)
        sleep_mock.assert_not_called()


if __name__ == "__main__":
    unittest.main()
