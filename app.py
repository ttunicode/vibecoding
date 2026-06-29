import io

import streamlit as st

from utils import (
    combine_excel_files,
    get_common_columns,
    to_excel_bytes,
)


st.set_page_config(page_title="엑셀 취합 앱", page_icon="📊", layout="wide")

st.title("엑셀 파일 취합 앱")
st.write("여러 개의 엑셀 파일을 업로드해 하나의 파일로 취합하고 다운로드하세요.")

with st.sidebar:
    st.header("설정")
    merge_mode = st.radio(
        "취합 방식",
        ["단순 누적 (행 방향)", "키(Key) 기준 병합 (열 방향)"],
        horizontal=False,
    )

uploaded_files = st.file_uploader(
    "엑셀 파일을 선택하세요",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.success(f"{len(uploaded_files)}개의 파일이 선택되었습니다.")
    file_names = [file.name for file in uploaded_files]
    st.write("선택된 파일:")
    st.code("\n".join(file_names), language="text")

    if merge_mode == "키(Key) 기준 병합 (열 방향)":
        common_columns = get_common_columns(uploaded_files)
        if common_columns:
            selected_key = st.selectbox(
                "병합 기준 열을 선택하세요",
                common_columns,
                index=0,
            )
        else:
            st.warning("공통 컬럼이 없어 키 기반 병합을 진행할 수 없습니다. 다른 양식의 파일을 선택해 주세요.")
            selected_key = None
    else:
        selected_key = None

    if st.button("취합 시작", use_container_width=True):
        if merge_mode == "키(Key) 기준 병합 (열 방향)" and not selected_key:
            st.error("병합 기준 열을 선택해 주세요.")
        else:
            with st.spinner("파일을 읽고 취합하는 중입니다..."):
                try:
                    result_df = combine_excel_files(
                        uploaded_files,
                        mode=merge_mode,
                        key_column=selected_key,
                    )
                except ValueError as exc:
                    st.error(str(exc))
                except Exception as exc:  # pragma: no cover - UI safety
                    st.error(f"처리 중 오류가 발생했습니다: {exc}")
                else:
                    st.success("취합이 완료되었습니다.")
                    st.subheader("미리보기")
                    st.dataframe(result_df.head(5), use_container_width=True)

                    excel_bytes = to_excel_bytes(result_df)
                    st.download_button(
                        label="취합 결과 다운로드",
                        data=excel_bytes,
                        file_name="merged_result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
else:
    st.info("업로드할 엑셀 파일을 선택해 주세요.")
