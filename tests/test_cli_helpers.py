"""Tests for CLI helper functions."""

import pytest
from outlookctl.cli import parse_recipient_args, parse_date


class TestParseRecipientArgs:
    def test_all_recipients(self):
        to_list, cc_list, bcc_list = parse_recipient_args(
            to="alice@example.com,bob@example.com",
            cc="charlie@example.com",
            bcc="secret@example.com",
        )
        assert to_list == ["alice@example.com", "bob@example.com"]
        assert cc_list == ["charlie@example.com"]
        assert bcc_list == ["secret@example.com"]

    def test_to_only(self):
        to_list, cc_list, bcc_list = parse_recipient_args(
            to="alice@example.com",
            cc=None,
            bcc=None,
        )
        assert to_list == ["alice@example.com"]
        assert cc_list == []
        assert bcc_list == []

    def test_empty_all(self):
        to_list, cc_list, bcc_list = parse_recipient_args(
            to=None,
            cc=None,
            bcc=None,
        )
        assert to_list == []
        assert cc_list == []
        assert bcc_list == []

    def test_whitespace_trimming(self):
        to_list, cc_list, bcc_list = parse_recipient_args(
            to="  alice@example.com  ,  bob@example.com  ",
            cc=None,
            bcc=None,
        )
        assert to_list == ["alice@example.com", "bob@example.com"]

    def test_single_recipient_no_comma(self):
        to_list, cc_list, bcc_list = parse_recipient_args(
            to="single@example.com",
            cc=None,
            bcc=None,
        )
        assert to_list == ["single@example.com"]

    def test_trailing_comma(self):
        """Trailing comma should not produce empty string in list."""
        to_list, cc_list, bcc_list = parse_recipient_args(
            to="alice@example.com,",
            cc=None,
            bcc=None,
        )
        assert to_list == ["alice@example.com"]
        assert "" not in to_list


class TestParseDate:
    def test_iso_date(self):
        result = parse_date("2025-01-15")
        assert result.year == 2025
        assert result.month == 1
        assert result.day == 15

    def test_iso_datetime(self):
        result = parse_date("2025-01-15T10:30:00")
        assert result.year == 2025
        assert result.hour == 10
        assert result.minute == 30

    def test_none_input(self):
        result = parse_date(None)
        assert result is None

    def test_empty_string(self):
        result = parse_date("")
        assert result is None

    def test_invalid_format(self):
        with pytest.raises(ValueError):
            parse_date("not-a-date")
