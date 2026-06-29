import io
from typing import List, Optional

import pandas as pd


def normalize_column_name(column_name: object) -> str:
    if column_name is None:
        return ""
    return str(column_name).strip().lower()


def _read_excel_dataframe(uploaded_file) -> pd.DataFrame:
    if uploaded_file is None:
        raise ValueError("업로드된 파일이 없습니다.")

    if getattr(uploaded_file, "size", 0) in (None, 0):
        raise ValueError(f"{uploaded_file.name} 파일이 비어 있습니다.")

    file_bytes = uploaded_file.getvalue()
    if not file_bytes:
        raise ValueError(f"{uploaded_file.name} 파일이 비어 있습니다.")

    try:
        if uploaded_file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
    except Exception as exc:
        raise ValueError(f"{uploaded_file.name} 파일을 읽는 중 문제가 발생했습니다: {exc}") from exc

    if df is None or df.empty:
        raise ValueError(f"{uploaded_file.name} 파일에 데이터가 없습니다.")

    df = df.copy()
    if not list(df.columns):
        raise ValueError(f"{uploaded_file.name} 파일의 컬럼을 확인할 수 없습니다.")

    df.columns = [normalize_column_name(col) for col in df.columns]
    return df


def get_common_columns(uploaded_files) -> List[str]:
    if not uploaded_files:
        return []

    parsed_dataframes = []
    for uploaded_file in uploaded_files:
        parsed_dataframes.append(_read_excel_dataframe(uploaded_file))

    if not parsed_dataframes:
        return []

    common = set(parsed_dataframes[0].columns)
    for df in parsed_dataframes[1:]:
        common &= set(df.columns)

    return sorted(list(common))


def combine_excel_files(uploaded_files, mode: str = "단순 누적 (행 방향)", key_column: Optional[str] = None) -> pd.DataFrame:
    if not uploaded_files:
        raise ValueError("업로드된 파일이 없습니다.")

    parsed_dataframes = []
    for uploaded_file in uploaded_files:
        parsed_dataframes.append(_read_excel_dataframe(uploaded_file))

    if mode == "단순 누적 (행 방향)":
        all_columns = []
        for df in parsed_dataframes:
            all_columns.extend(list(df.columns))
        union_columns = list(dict.fromkeys(all_columns))

        for df in parsed_dataframes:
            for col in union_columns:
                if col not in df.columns:
                    df[col] = pd.NA

        for idx, df in enumerate(parsed_dataframes):
            df.insert(0, "원본_파일명", uploaded_files[idx].name)

        combined_df = pd.concat(parsed_dataframes, ignore_index=True, sort=False)
        combined_df = combined_df.loc[:, ["원본_파일명"] + [col for col in combined_df.columns if col != "원본_파일명"]]
        return combined_df

    if mode == "키(Key) 기준 병합 (열 방향)":
        if not key_column:
            raise ValueError("병합 기준 열을 선택해 주세요.")

        for df in parsed_dataframes:
            if key_column not in df.columns:
                raise ValueError(f"선택한 기준 열 '{key_column}'이(가) 파일에 없습니다.")

        merged_df = None
        for idx, df in enumerate(parsed_dataframes):
            df_work = df.copy()
            df_work = df_work.set_index(key_column)
            if merged_df is None:
                merged_df = df_work
            else:
                merged_df = merged_df.join(df_work, how="outer", rsuffix=f"_{idx}")

        merged_df = merged_df.reset_index()
        merged_df.insert(0, "원본_파일명", "; ".join(file.name for file in uploaded_files))
        return merged_df

    raise ValueError("지원하지 않는 취합 방식입니다.")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="취합결과")
    output.seek(0)
    return output.getvalue()
