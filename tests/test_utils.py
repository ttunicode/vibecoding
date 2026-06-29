from io import BytesIO
from types import SimpleNamespace

import pandas as pd

from utils import combine_excel_files


def make_uploaded_file(name: str, rows: dict) -> SimpleNamespace:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    return SimpleNamespace(name=name, getvalue=lambda: buffer.getvalue(), size=len(buffer.getvalue()))


def test_simple_accumulate_adds_source_column():
    files = [
        make_uploaded_file("file_a.xlsx", {"id": [1, 2], "name": ["A", "B"]}),
        make_uploaded_file("file_b.xlsx", {"id": [3], "name": ["C"]}),
    ]

    result = combine_excel_files(files, mode="단순 누적 (행 방향)")

    assert "원본_파일명" in result.columns
    assert len(result) == 3


def test_key_merge_uses_selected_key():
    files = [
        make_uploaded_file("file_a.xlsx", {"id": [1, 2], "value_a": [10, 20]}),
        make_uploaded_file("file_b.xlsx", {"id": [2, 3], "value_b": [30, 40]}),
    ]

    result = combine_excel_files(files, mode="키(Key) 기준 병합 (열 방향)", key_column="id")

    assert "id" in result.columns
    assert "원본_파일명" in result.columns
    assert result.shape[0] >= 3
