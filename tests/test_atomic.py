"""Atomic output: a failure at any point leaves the previous file untouched and no temporary file behind."""

from __future__ import annotations

import os
import stat
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from phpp_tool import atomic
from phpp_tool.atomic import atomic_copy, atomic_write_bytes, atomic_write_text
from phpp_tool.surgical_writer import apply_surgical_writes


def _only(folder: Path, name: str) -> bool:
    return sorted(p.name for p in folder.iterdir()) == [name]


class TestAtomicWrite:
    def test_writes_a_new_file_and_creates_its_folder(self, tmp_path: Path):
        target = tmp_path / "records" / "out.json"
        atomic_write_text(target, '{"a": 1}')
        assert target.read_text(encoding="utf-8") == '{"a": 1}'
        assert _only(target.parent, "out.json")
        assert stat.S_IMODE(target.stat().st_mode) == 0o644

    def test_replaces_an_existing_file_and_keeps_its_mode(self, tmp_path: Path):
        target = tmp_path / "out.bin"
        target.write_bytes(b"old")
        target.chmod(0o600)
        atomic_write_bytes(target, b"new")
        assert target.read_bytes() == b"new"
        assert stat.S_IMODE(target.stat().st_mode) == 0o600
        assert _only(tmp_path, "out.bin")

    def test_a_failure_while_replacing_keeps_the_old_file(
        self, tmp_path: Path, monkeypatch
    ):
        target = tmp_path / "out.bin"
        target.write_bytes(b"old")

        def boom(*_args, **_kwargs):
            raise OSError("cannot replace")

        monkeypatch.setattr(atomic.os, "replace", boom)
        with pytest.raises(OSError, match="cannot replace"):
            atomic_write_bytes(target, b"new")
        assert target.read_bytes() == b"old"
        assert _only(tmp_path, "out.bin")

    def test_a_failure_while_writing_keeps_the_old_file(
        self, tmp_path: Path, monkeypatch
    ):
        target = tmp_path / "out.bin"
        target.write_bytes(b"old")

        def boom(_fd):
            raise OSError("disk full")

        monkeypatch.setattr(atomic.os, "fsync", boom)
        with pytest.raises(OSError, match="disk full"):
            atomic_write_bytes(target, b"new")
        assert target.read_bytes() == b"old"
        assert _only(tmp_path, "out.bin")

    def test_a_failure_with_no_previous_file_leaves_nothing(
        self, tmp_path: Path, monkeypatch
    ):
        target = tmp_path / "out.bin"
        monkeypatch.setattr(
            atomic.os, "replace", lambda *_a, **_k: (_ for _ in ()).throw(OSError("x"))
        )
        with pytest.raises(OSError):
            atomic_write_bytes(target, b"new")
        assert list(tmp_path.iterdir()) == []


class TestAtomicCopy:
    def test_copies_contents(self, tmp_path: Path):
        source = tmp_path / "in.bin"
        source.write_bytes(b"payload")
        atomic_copy(source, tmp_path / "sub" / "out.bin")
        assert (tmp_path / "sub" / "out.bin").read_bytes() == b"payload"
        assert _only(tmp_path / "sub", "out.bin")

    def test_a_failure_keeps_the_old_destination(self, tmp_path: Path, monkeypatch):
        source = tmp_path / "in.bin"
        source.write_bytes(b"payload")
        target = tmp_path / "out.bin"
        target.write_bytes(b"old")
        monkeypatch.setattr(
            atomic.os, "replace", lambda *_a, **_k: (_ for _ in ()).throw(OSError("x"))
        )
        with pytest.raises(OSError):
            atomic_copy(source, target)
        assert target.read_bytes() == b"old"
        assert sorted(p.name for p in tmp_path.iterdir()) == ["in.bin", "out.bin"]


class TestWorkbookOutput:
    def _template(self, tmp_path: Path) -> Path:
        template = tmp_path / "template.xlsx"
        wb = Workbook()
        wb.active.title = "Sheet1"
        wb.save(template)
        return template

    def test_a_failed_write_keeps_an_existing_output_workbook(
        self, tmp_path: Path, monkeypatch
    ):
        template = self._template(tmp_path)
        output = tmp_path / "output.xlsx"
        output.write_bytes(b"the designer's earlier output")
        monkeypatch.setattr(
            atomic.os, "replace", lambda *_a, **_k: (_ for _ in ()).throw(OSError("x"))
        )
        with pytest.raises(OSError):
            apply_surgical_writes(template, output, [("Sheet1", "B", 2, 5)])
        assert output.read_bytes() == b"the designer's earlier output"
        assert sorted(p.name for p in tmp_path.iterdir()) == [
            "output.xlsx",
            "template.xlsx",
        ]

    def test_a_failed_copy_path_keeps_an_existing_output_workbook(
        self, tmp_path: Path, monkeypatch
    ):
        template = self._template(tmp_path)
        output = tmp_path / "output.xlsx"
        output.write_bytes(b"earlier")
        monkeypatch.setattr(
            atomic.os, "replace", lambda *_a, **_k: (_ for _ in ()).throw(OSError("x"))
        )
        with pytest.raises(OSError):
            apply_surgical_writes(
                template, output, []
            )  # nothing to patch: the template is copied as it is
        assert output.read_bytes() == b"earlier"

    def test_a_good_write_replaces_the_output_and_leaves_no_temp_file(
        self, tmp_path: Path
    ):
        template = self._template(tmp_path)
        output = tmp_path / "output.xlsx"
        output.write_bytes(b"earlier")
        apply_surgical_writes(template, output, [("Sheet1", "B", 2, 5)])
        assert zipfile.is_zipfile(output)
        assert sorted(p.name for p in tmp_path.iterdir()) == [
            "output.xlsx",
            "template.xlsx",
        ]

    def test_writing_over_the_template_itself_works(self, tmp_path: Path):
        template = self._template(tmp_path)
        apply_surgical_writes(template, template, [("Sheet1", "B", 2, 5)])
        assert zipfile.is_zipfile(template)
        assert os.listdir(tmp_path) == ["template.xlsx"]


class TestCompareJsonOutput:
    def test_reports_are_written_atomically(self, tmp_path: Path, monkeypatch):
        from click.testing import CliRunner

        from compare_json.cli import main

        a, b = tmp_path / "a.json", tmp_path / "b.json"
        a.write_text('{"x": 1}')
        b.write_text('{"x": 2}')
        report = tmp_path / "report.txt"
        report.write_text("earlier report")
        runner = CliRunner()

        monkeypatch.setattr(
            atomic.os, "replace", lambda *_a, **_k: (_ for _ in ()).throw(OSError("x"))
        )
        failed = runner.invoke(main, [str(a), str(b), "-o", str(report)])
        assert failed.exit_code != 0
        assert report.read_text() == "earlier report"

        monkeypatch.undo()
        ok = runner.invoke(main, [str(a), str(b), "-o", str(report)])
        assert ok.exit_code == 0
        assert report.read_text() != "earlier report"
        assert sorted(p.name for p in tmp_path.iterdir()) == [
            "a.json",
            "b.json",
            "report.txt",
        ]
