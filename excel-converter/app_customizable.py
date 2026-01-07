# app_upload_fix.py
# 실행: streamlit run app_upload_fix.py
# 필요: pip install streamlit pandas openpyxl
# (.xls 읽기 필요 시) pip install "xlrd==1.2.0"

import io
import re
import json
import zipfile
from datetime import datetime
from typing import Optional, List

import pandas as pd
import streamlit as st

st.set_page_config(page_title="엑셀 양식 변환기 (1→2)", layout="centered")

st.title("엑셀 양식 변환기 (1 → 2)")
st.caption("라오라 / 쿠팡 / 스마트스토어(키워드) / 떠리몰(S&V 규칙) 형식을 2번 템플릿으로 변환합니다. (전화번호 0 보존)")

# -------------------------- Helpers --------------------------
def excel_col_to_index(col_letters: str) -> int:
    col_letters = str(col_letters).strip().upper()
    if not re.fullmatch(r"[A-Z]+", col_letters):
        raise ValueError(f"Invalid Excel column letters: {col_letters}")
    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1  # 0-based

def index_to_excel_col(n: int) -> str:
    s = ""
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(r + 65) + s
    return s

def excel_letters(max_cols=104):
    return [index_to_excel_col(i) for i in range(max_cols)]

def read_first_sheet_template(file) -> pd.DataFrame:
    """템플릿(2.xlsx)은 일반적으로 읽기"""
    return pd.read_excel(file, sheet_name=0, header=0, engine="openpyxl")

def read_first_sheet_source_as_text(file) -> pd.DataFrame:
    """소스는 전 컬럼을 문자열로 읽어 전화번호 앞 0 보존"""
    return pd.read_excel(
        file,
        sheet_name=0,
        header=0,
        engine="openpyxl",
        dtype=str,
        keep_default_na=False,  # 빈값을 NaN 대신 빈 문자열로 유지
    )

def ensure_mapping_initialized(template_columns, default_mapping):
    m = st.session_state.get("mapping")
    if not isinstance(m, dict):
        m = {}
    synced = {k: str(v).upper() for k, v in m.items() if k in template_columns and v}
    for k in template_columns:
        if k not in synced and k in default_mapping:
            synced[k] = default_mapping[k]
    st.session_state["mapping"] = synced
    return st.session_state["mapping"]

def norm_header(s: str) -> str:
    return re.sub(r"[\s\(\)\[\]{}:：/\\\-]", "", str(s).strip().lower())

# -------------------- Defaults --------------------
DEFAULT_TEMPLATE_COLUMNS = [
    "주문번호",
    "받는분 이름",
    "받는분 주소",
    "받는분 전화번호",
    "상품명",
    "수량",
    "메모",
]

# 라오라 기본 매핑 (열 문자)
DEFAULT_MAPPING = {
    "주문번호": "A",
    "받는분 이름": "I",
    "받는분 주소": "L",
    "받는분 전화번호": "J",
    "상품명": "D",
    "수량": "G",
    "메모": "M",
}

# 쿠팡 고정 매핑 (열 문자) — 주문번호 C
COUPANG_MAPPING = {
    "주문번호": "C",
    "받는분 이름": "AA",
    "받는분 주소": "AD",
    "받는분 전화번호": "AB",
    "상품명": "P",
    "수량": "W",
    "메모": "AE",
}

# 스마트스토어 키워드 매핑용 후보
SS_NAME_MAP = {
    "주문번호": ["주문번호"],
    "받는분 이름": ["수취인명"],
    "받는분 주소": ["통합배송지"],
    "받는분 전화번호": ["수취인연락처1", "수취인연락처", "수취인휴대폰", "연락처1"],
    "상품명_left": ["상품명"],
    "상품명_right": ["옵션정보", "옵션명", "옵션내용"],
    "수량": ["수량", "구매수량"],
    "메모": ["배송메세지", "배송메시지", "배송요청사항"],
}

# 떠리몰 고정 매핑 (열 문자) + 상품명 S&V 규칙
TTARIMALL_FIXED_LETTER_MAPPING = {
    "주문번호": "H",
    "받는분 이름": "AB",
    "받는분 주소": "AE",
    "받는분 전화번호": "AC",
    "상품명": "V",  # 비교는 S와 수행
    "수량": "Y",
    "메모": "AA",
}

# -------------------------- Sidebar --------------------------
st.sidebar.header("템플릿 옵션")
use_uploaded_template = st.sidebar.checkbox("템플릿(2.xlsx) 직접 업로드", value=False)
max_letter_cols = st.sidebar.slider(
    "라오라용 최대 열 범위(Excel 문자)",
    min_value=52,
    max_value=156,
    value=104,
    step=26,
    help="라오라 매핑 드롭다운의 열 문자 개수",
)
st.sidebar.divider()
st.sidebar.subheader("라오라 매핑 저장/불러오기")
mapping_upload = st.sidebar.file_uploader("매핑 JSON 불러오기 (라오라)", type=["json"], key="mapping_json")
prepare_download = st.sidebar.button("현재 라오라 매핑 JSON 다운로드 준비")

# -------------------------- 템플릿 설정 (공용) --------------------------
st.subheader("템플릿 설정 (2.xlsx)")
tpl_df = None
if use_uploaded_template:
    tpl_file = st.file_uploader("2와 같은 템플릿 파일 업로드 (예: 2.xlsx)", type=["xlsx"], key="tpl")
    if tpl_file:
        try:
            tpl_df = read_first_sheet_template(tpl_file)
            st.success(f"템플릿 업로드 완료. 컬럼 수: {len(tpl_df.columns)}")
        except Exception as e:
            st.warning(f"템플릿 파일을 읽는 중 오류가 발생했습니다: {e}")
            tpl_df = None
else:
    tpl_df = pd.DataFrame(columns=DEFAULT_TEMPLATE_COLUMNS)
    st.info("업로드된 템플릿이 없으므로 기본 템플릿을 사용합니다. (주문번호, 받는분 이름, 받는분 주소, 받는분 전화번호, 상품명, 수량, 메모)")

template_columns = list(tpl_df.columns) if tpl_df is not None else []

# ======================================================================
# 1) 라오라 파일 변환 (열 문자 매핑)
# ======================================================================
st.markdown("## 라오라 파일 변환")

current_mapping = ensure_mapping_initialized(template_columns, DEFAULT_MAPPING)
letters = excel_letters(max_letter_cols)

if mapping_upload is not None:
    try:
        loaded = json.load(mapping_upload)
        if not isinstance(loaded, dict):
            raise ValueError("JSON 루트가 객체(dict)가 아닙니다.")
        new_map = {}
        for k, v in loaded.items():
            if k in template_columns and isinstance(v, str) and re.fullmatch(r"[A-Za-z]+", v):
                new_map[k] = v.upper()
        for k in template_columns:
            if k not in new_map:
                new_map[k] = current_mapping.get(k, DEFAULT_MAPPING.get(k, ""))
        st.session_state["mapping"] = new_map
        current_mapping = new_map
        st.success("라오라 매핑 JSON을 불러왔습니다.")
    except Exception as e:
        st.warning(f"라오라 매핑 JSON 불러오기 실패: {e}")

edited_mapping = {}
with st.form("mapping_form_laora"):
    for col in template_columns:
        default_val = current_mapping.get(col, "")
        if default_val not in letters:
            default_val = ""
        options = [""] + letters
        sel = st.selectbox(
            f"{col} ⟶ 1.xlsx(라오라) 열 문자 선택",
            options=options,
            index=(options.index(default_val) if default_val in options else 0),
            key=f"map_laora_{col}",
        )
        edited_mapping[col] = sel
    if st.form_submit_button("라오라 매핑 저장"):
        st.session_state["mapping"] = {k: v for k, v in edited_mapping.items() if v}
        current_mapping = st.session_state["mapping"]
        st.success("라오라 매핑을 저장했습니다.")

if prepare_download:
    mapping_bytes = json.dumps(current_mapping, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button(
        label="현재 라오라 매핑 JSON 다운로드",
        data=mapping_bytes,
        file_name="mapping_laora.json",
        mime="application/json",
    )

st.subheader("라오라 소스 파일 업로드")
src_file_laora = st.file_uploader("라오라 형식의 파일 업로드 (예: 1.xlsx)", type=["xlsx"], key="src_laora")
run_laora = st.button("라오라 변환 실행")
if run_laora:
    if not src_file_laora:
        st.error("라오라 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_src = read_first_sheet_source_as_text(src_file_laora)
        except Exception as e:
            st.exception(RuntimeError(f"라오라 소스 파일을 읽는 중 오류: {e}"))
        else:
            result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
            mapping = st.session_state.get("mapping", {})
            if not isinstance(mapping, dict) or not mapping:
                st.error("라오라 매핑이 없습니다. 먼저 저장해 주세요.")
            else:
                src_cols_by_index = list(df_src.columns)
                resolved_map = {}
                try:
                    for tpl_header, xl_letters in mapping.items():
                        if not xl_letters:
                            continue
                        idx = excel_col_to_index(xl_letters)
                        if idx >= len(src_cols_by_index):
                            raise IndexError(
                                f"소스 파일에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                                f"소스 컬럼 수: {len(src_cols_by_index)}"
                            )
                        resolved_map[tpl_header] = src_cols_by_index[idx]
                except Exception as e:
                    st.exception(RuntimeError(f"라오라 매핑 인덱스 계산 중 오류: {e}"))
                else:
                    for tpl_header, src_colname in resolved_map.items():
                        try:
                            if tpl_header == "수량":
                                result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
                            elif tpl_header == "받는분 전화번호":
                                series = df_src[src_colname].astype(str)
                                result[tpl_header] = series.where(series.str.lower() != "nan", "")
                            else:
                                result[tpl_header] = df_src[src_colname]
                        except KeyError:
                            st.warning(f"소스 컬럼 '{src_colname}'(매핑: {tpl_header})을(를) 찾을 수 없습니다. 해당 필드는 비워집니다.")

                    # 템플릿 숫자형 정렬(전화번호 제외)
                    for col in template_columns:
                        if col in tpl_df.columns and tpl_df[col].notna().any():
                            if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                                result[col] = pd.to_numeric(result[col], errors="coerce")

                    st.success(f"라오라 변환 완료: 총 {len(result)}행")
                    st.dataframe(result.head(50))

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        out_df = result[template_columns + [c for c in result.columns if c not in template_columns]]
                        out_df.to_excel(writer, index=False)

                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label=f"라오라 변환 결과 다운로드 (라오 3pl발주용_{ts}.xlsx)",
                        data=buffer.getvalue(),
                        file_name=f"라오 3pl발주용_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


st.markdown("---")

# ======================================================================
# 2) 쿠팡 파일 변환 (고정 매핑)
# ======================================================================
st.markdown("## 쿠팡 파일 변환")

with st.expander("쿠팡 → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **쿠팡 소스열 → 템플릿 컬럼**  
        - `C` → **주문번호**  
        - `AA` → **받는분 이름**  
        - `AD` → **받는분 주소**  
        - `AB` → **받는분 전화번호**  
        - `P` → **상품명** (최초등록상품명/옵션명)  
        - `W` → **수량** (구매수)  
        - `AE` → **메모** (배송메시지)
        """
    )

st.subheader("쿠팡 소스 파일 업로드")
src_file_coupang = st.file_uploader("쿠팡 형식의 파일 업로드 (예: 쿠팡.xlsx)", type=["xlsx"], key="src_coupang")
run_coupang = st.button("쿠팡 변환 실행")
if run_coupang:
    if not src_file_coupang:
        st.error("쿠팡 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_src_cp = read_first_sheet_source_as_text(src_file_coupang)
        except Exception as e:
            st.exception(RuntimeError(f"쿠팡 소스 파일을 읽는 중 오류: {e}"))
        else:
            result_cp = pd.DataFrame(index=range(len(df_src_cp)), columns=template_columns)
            mapping_cp = COUPANG_MAPPING.copy()

            src_cols_by_index_cp = list(df_src_cp.columns)
            resolved_map_cp = {}
            try:
                for tpl_header, xl_letters in mapping_cp.items():
                    idx = excel_col_to_index(xl_letters)
                    if idx >= len(src_cols_by_index_cp):
                        raise IndexError(
                            f"쿠팡 소스에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                            f"소스 컬럼 수: {len(src_cols_by_index_cp)}"
                        )
                    resolved_map_cp[tpl_header] = src_cols_by_index_cp[idx]
            except Exception as e:
                st.exception(RuntimeError(f"쿠팡 매핑 인덱스 계산 중 오류: {e}"))
            else:
                for tpl_header, src_colname in resolved_map_cp.items():
                    try:
                        if tpl_header == "수량":
                            result_cp[tpl_header] = pd.to_numeric(df_src_cp[src_colname], errors="coerce")
                        elif tpl_header == "받는분 전화번호":
                            series = df_src_cp[src_colname].astype(str)
                            result_cp[tpl_header] = series.where(series.str.lower() != "nan", "")
                        else:
                            result_cp[tpl_header] = df_src_cp[src_colname]
                    except KeyError:
                        st.warning(f"[쿠팡] 소스 컬럼 '{src_colname}'(매핑: {tpl_header})을(를) 찾을 수 없습니다. 해당 필드는 비워집니다.")

                # 템플릿 숫자형 정렬(전화번호 제외)
                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_cp[col] = pd.to_numeric(result_cp[col], errors="coerce")

                st.success(f"쿠팡 변환 완료: 총 {len(result_cp)}행")
                st.dataframe(result_cp.head(50))

                buffer_cp = io.BytesIO()
                with pd.ExcelWriter(buffer_cp, engine="openpyxl") as writer:
                    out_df_cp = result_cp[template_columns + [c for c in result_cp.columns if c not in template_columns]]
                    out_df_cp.to_excel(writer, index=False)

                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label=f"쿠팡 변환 결과 다운로드 (쿠팡 3pl발주용_{ts}.xlsx)",
                    data=buffer_cp.getvalue(),
                    file_name=f"쿠팡 3pl발주용_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")

# ======================================================================
# 3) 스마트스토어 파일 변환 (키워드 매핑)
# ======================================================================
st.markdown("## 스마트스토어 파일 변환 (키워드 매핑)")

with st.expander("스마트스토어(키워드) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **스마트스토어 컬럼명(헤더) → 템플릿 컬럼**  
        - `주문번호` → **주문번호**  
        - `수취인명` → **받는분 이름**  
        - `통합배송지` → **받는분 주소**  
        - `수취인연락처1` → **받는분 전화번호**  
        - `=상품명&옵션정보` → **상품명** (두 값을 그대로 연결)  
        - `수량` → **수량**  
        - `배송메세지` → **메모**  (※ 일부 파일은 `배송메시지` 표기)
        """
    )

st.subheader("스마트스토어 소스 파일 업로드 (키워드 매핑)")
src_file_ss_fixed = st.file_uploader(
    "스마트스토어 형식의 파일 업로드 (예: 스마트스토어.xlsx)",
    type=["xlsx"],
    key="src_smartstore_fixed",
)

run_ss_fixed = st.button("스마트스토어 변환 실행 (키워드 매핑)")
if run_ss_fixed:
    if not src_file_ss_fixed:
        st.error("스마트스토어 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_ss = read_first_sheet_source_as_text(src_file_ss_fixed)
        except Exception as e:
            st.exception(RuntimeError(f"스마트스토어 소스 파일을 읽는 중 오류: {e}"))
        else:
            def find_col(preferred_names, df):
                norm_cols = {norm_header(c): c for c in df.columns}
                cand_norm = [norm_header(x) for x in preferred_names]
                for n in cand_norm:
                    if n in norm_cols:
                        return norm_cols[n]
                for want in cand_norm:
                    hits = [orig for k, orig in norm_cols.items() if want in k]
                    if hits:
                        return sorted(hits, key=len)[0]
                raise KeyError(f"해당 키워드에 맞는 컬럼을 찾을 수 없습니다: {preferred_names}")

            try:
                col_order = find_col(SS_NAME_MAP["주문번호"], df_ss)
                col_name  = find_col(SS_NAME_MAP["받는분 이름"], df_ss)
                col_addr  = find_col(SS_NAME_MAP["받는분 주소"], df_ss)
                col_phone = find_col(SS_NAME_MAP["받는분 전화번호"], df_ss)
                col_prod_l = find_col(SS_NAME_MAP["상품명_left"], df_ss)
                col_prod_r = find_col(SS_NAME_MAP["상품명_right"], df_ss)
                col_qty   = find_col(SS_NAME_MAP["수량"], df_ss)
                col_memo  = find_col(SS_NAME_MAP["메모"], df_ss)
            except Exception as e:
                st.exception(RuntimeError(f"스마트스토어 키워드 매핑 해석 중 오류: {e}"))
            else:
                result_ss = pd.DataFrame(index=range(len(df_ss)), columns=template_columns)

                result_ss["주문번호"] = df_ss[col_order]
                result_ss["받는분 이름"] = df_ss[col_name]
                result_ss["받는분 주소"] = df_ss[col_addr]

                series_phone = df_ss[col_phone].astype(str)
                result_ss["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")

                left_raw = df_ss[col_prod_l].astype(str)
                right_raw = df_ss[col_prod_r].astype(str)
                left = left_raw.where(left_raw.str.lower() != "nan", "")
                right = right_raw.where(right_raw.str.lower() != "nan", "")
                result_ss["상품명"] = (left.fillna("") + right.fillna(""))

                result_ss["수량"] = pd.to_numeric(df_ss[col_qty], errors="coerce")
                result_ss["메모"] = df_ss[col_memo]

                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_ss[col] = pd.to_numeric(result_ss[col], errors="coerce")

                st.success(f"스마트스토어(키워드) 변환 완료: 총 {len(result_ss)}행")
                st.dataframe(result_ss.head(50))

                buffer_ss = io.BytesIO()
                with pd.ExcelWriter(buffer_ss, engine="openpyxl") as writer:
                    out_df_ss = result_ss[template_columns + [c for c in result_ss.columns if c not in template_columns]]
                    out_df_ss.to_excel(writer, index=False)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label=f"스마트스토어 변환 결과 다운로드 (스마트스토어 3pl발주용_{ts}.xlsx)",
                    data=buffer_ss.getvalue(),
                    file_name=f"스마트스토어 3pl발주용_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")

# ======================================================================
# 4) 떠리몰 파일 변환 (고정 매핑: 열 문자)
# ======================================================================
st.markdown("## 떠리몰 파일 변환 (고정 매핑: 열 문자)")

with st.expander("떠리몰(고정) → 템플릿 매핑 보기", expanded=False):
    st.markdown(
        """
        **떠리몰 소스열 → 템플릿 컬럼**  
        - `H` → **주문번호**  
        - `AB` → **받는분 이름** (수령자명)  
        - `AE` → **받는분 주소** (주소)  
        - `AC` → **받는분 전화번호** (수령자연락처)  
        - `S & V` → **상품명** (S와 V가 같으면 V만, 다르면 S&V로 연결)  
        - `Y` → **수량**  
        - `AA` → **메모** (배송메시지)
        """
    )

st.subheader("떠리몰 소스 파일 업로드 (고정 매핑)")
src_file_ttarimall = st.file_uploader("떠리몰 형식의 파일 업로드 (예: 떠리몰.xlsx)", type=["xlsx"], key="src_ttarimall")

run_ttarimall = st.button("떠리몰 변환 실행 (고정 매핑)")
if run_ttarimall:
    if not src_file_ttarimall:
        st.error("떠리몰 소스 파일을 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        try:
            df_tm = read_first_sheet_source_as_text(src_file_ttarimall)
        except Exception as e:
            st.exception(RuntimeError(f"떠리몰 소스 파일을 읽는 중 오류: {e}"))
        else:
            result_tm = pd.DataFrame(index=range(len(df_tm)), columns=template_columns)

            src_cols_by_index_tm = list(df_tm.columns)

            def resolve(letter: str) -> str:
                idx = excel_col_to_index(letter)
                if idx >= len(src_cols_by_index_tm):
                    raise IndexError(
                        f"떠리몰 소스에 {letter} 열(0-based index {idx})이 없습니다. "
                        f"소스 컬럼 수: {len(src_cols_by_index_tm)}"
                    )
                return src_cols_by_index_tm[idx]

            try:
                col_order = resolve(TTARIMALL_FIXED_LETTER_MAPPING["주문번호"])
                col_name = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 이름"])
                col_addr = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 주소"])
                col_phone = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 전화번호"])
                col_prod_v = resolve(TTARIMALL_FIXED_LETTER_MAPPING["상품명"])  # V
                col_prod_s = resolve("S")  # S 열도 함께 사용
                col_qty = resolve(TTARIMALL_FIXED_LETTER_MAPPING["수량"])
                col_memo = resolve(TTARIMALL_FIXED_LETTER_MAPPING["메모"])
            except Exception as e:
                st.exception(RuntimeError(f"떠리몰 고정 매핑 인덱스 계산 중 오류: {e}"))
            else:
                result_tm["주문번호"] = df_tm[col_order]
                result_tm["받는분 이름"] = df_tm[col_name]
                result_tm["받는분 주소"] = df_tm[col_addr]
                series_phone = df_tm[col_phone].astype(str)
                result_tm["받는분 전화번호"] = series_phone.where(series_phone.str.lower() != "nan", "")

                # 상품명: S와 V가 같으면 V, 다르면 S&V
                s_series_raw = df_tm[col_prod_s].astype(str)
                v_series_raw = df_tm[col_prod_v].astype(str)
                s_series = s_series_raw.where(s_series_raw.str.lower() != "nan", "")
                v_series = v_series_raw.where(v_series_raw.str.lower() != "nan", "")
                same_mask = (s_series == v_series)
                prod_series = v_series.copy()
                prod_series.loc[~same_mask] = s_series[~same_mask] + v_series[~same_mask]
                result_tm["상품명"] = prod_series

                result_tm["수량"] = pd.to_numeric(df_tm[col_qty], errors="coerce")
                result_tm["메모"] = df_tm[col_memo]

                for col in template_columns:
                    if col in tpl_df.columns and tpl_df[col].notna().any():
                        if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                            result_tm[col] = pd.to_numeric(result_tm[col], errors="coerce")

                st.success(f"떠리몰(고정) 변환 완료: 총 {len(result_tm)}행")
                st.dataframe(result_tm.head(50))

                buffer_tm = io.BytesIO()
                with pd.ExcelWriter(buffer_tm, engine="openpyxl") as writer:
                    out_df_tm = result_tm[template_columns + [c for c in result_tm.columns if c not in template_columns]]
                    out_df_tm.to_excel(writer, index=False)

                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label=f"떠리몰 변환 결과 다운로드 (떠리몰 3pl발주용_{ts}.xlsx)",
                    data=buffer_tm.getvalue(),
                    file_name=f"떠리몰 3pl발주용_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")

# ======================================================================
# 5) 배치 처리: 여러 파일 자동 분류 → 일괄 변환 → ZIP 다운로드
# ======================================================================
st.markdown("## 🗂️ 배치 처리 (여러 파일 한번에)")

batch_files = st.file_uploader("여러 엑셀 파일을 한번에 업로드하세요", type=["xlsx"], accept_multiple_files=True, key="batch_files")
run_batch = st.button("배치 변환 실행")

def detect_platform_by_headers(df: pd.DataFrame) -> str:
    headers = [norm_header(c) for c in df.columns]

    def has_any(keys):
        keys_norm = [norm_header(k) for k in keys]
        return any(k in headers for k in keys_norm)

    # 떠리몰 신호
    if has_any(["수령자명", "수령자연락처", "옵션명:옵션값"]):
        return "TTARIMALL"
    # 스마트스토어 신호
    if has_any(["수취인명", "수취인연락처1", "통합배송지"]):
        return "SMARTSTORE"
    # 쿠팡 신호
    if has_any(["최초등록상품명"]) or (has_any(["구매수"]) and has_any(["옵션명"])) or has_any(["배송메시지"]):
        return "COUPANG"
    # 그 외 → 라오라로 가정
    return "LAORA"

def convert_laora(df_src: pd.DataFrame) -> pd.DataFrame:
    mapping = st.session_state.get("mapping", {})
    if not isinstance(mapping, dict) or not mapping:
        raise RuntimeError("라오라 매핑이 없습니다. 사이드바에서 라오라 매핑을 먼저 저장해 주세요.")
    result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
    src_cols_by_index = list(df_src.columns)
    resolved_map = {}
    for tpl_header, xl_letters in mapping.items():
        if not xl_letters:
            continue
        idx = excel_col_to_index(xl_letters)
        if idx >= len(src_cols_by_index):
            raise IndexError(
                f"소스 파일에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                f"소스 컬럼 수: {len(src_cols_by_index)}"
            )
        resolved_map[tpl_header] = src_cols_by_index[idx]
    for tpl_header, src_colname in resolved_map.items():
        if tpl_header == "수량":
            result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
        elif tpl_header == "받는분 전화번호":
            series = df_src[src_colname].astype(str)
            result[tpl_header] = series.where(series.str.lower() != "nan", "")
        else:
            result[tpl_header] = df_src[src_colname]
    return result

def convert_coupang(df_src: pd.DataFrame) -> pd.DataFrame:
    result = pd.DataFrame(index=range(len(df_src)), columns=template_columns)
    src_cols_by_index = list(df_src.columns)
    resolved_map = {}
    for tpl_header, xl_letters in COUPANG_MAPPING.items():
        idx = excel_col_to_index(xl_letters)
        if idx >= len(src_cols_by_index):
            raise IndexError(
                f"쿠팡 소스에 {xl_letters} 열(0-based index {idx})이 존재하지 않습니다. "
                f"소스 컬럼 수: {len(src_cols_by_index)}"
            )
        resolved_map[tpl_header] = src_cols_by_index[idx]
    for tpl_header, src_colname in resolved_map.items():
        if tpl_header == "수량":
            result[tpl_header] = pd.to_numeric(df_src[src_colname], errors="coerce")
        elif tpl_header == "받는분 전화번호":
            series = df_src[src_colname].astype(str)
            result[tpl_header] = series.where(series.str.lower() != "nan", "")
        else:
            result[tpl_header] = df_src[src_colname]
    return result

def find_col(preferred_names, df):
    norm_cols = {norm_header(c): c for c in df.columns}
    cand_norm = [norm_header(x) for x in preferred_names]
    for n in cand_norm:
        if n in norm_cols:
            return norm_cols[n]
    for want in cand_norm:
        hits = [orig for k, orig in norm_cols.items() if want in k]
        if hits:
            return sorted(hits, key=len)[0]
    raise KeyError(f"해당 키워드에 맞는 컬럼을 찾을 수 없습니다: {preferred_names}")

def convert_smartstore_keywords(df_ss: pd.DataFrame) -> pd.DataFrame:
    col_order = find_col(SS_NAME_MAP["주문번호"], df_ss)
    col_name = find_col(SS_NAME_MAP["받는분 이름"], df_ss)
    col_addr = find_col(SS_NAME_MAP["받는분 주소"], df_ss)
    col_phone = find_col(SS_NAME_MAP["받는분 전화번호"], df_ss)
    col_prod_l = find_col(SS_NAME_MAP["상품명_left"], df_ss)
    col_prod_r = find_col(SS_NAME_MAP["상품명_right"], df_ss)
    col_qty = find_col(SS_NAME_MAP["수량"], df_ss)
    col_memo = find_col(SS_NAME_MAP["메모"], df_ss)

    result = pd.DataFrame(index=range(len(df_ss)), columns=template_columns)
    result["주문번호"] = df_ss[col_order]
    result["받는분 이름"] = df_ss[col_name]
    result["받는분 주소"] = df_ss[col_addr]
    phone = df_ss[col_phone].astype(str)
    result["받는분 전화번호"] = phone.where(phone.str.lower() != "nan", "")
    lraw = df_ss[col_prod_l].astype(str)
    rraw = df_ss[col_prod_r].astype(str)
    l = lraw.where(lraw.str.lower() != "nan", "")
    r = rraw.where(rraw.str.lower() != "nan", "")
    result["상품명"] = l.fillna("") + r.fillna("")
    result["수량"] = pd.to_numeric(df_ss[col_qty], errors="coerce")
    result["메모"] = df_ss[col_memo]
    return result

def convert_ttarimall(df_tm: pd.DataFrame) -> pd.DataFrame:
    src_cols_by_index = list(df_tm.columns)

    def resolve(letter: str) -> str:
        idx = excel_col_to_index(letter)
        if idx >= len(src_cols_by_index):
            raise IndexError(
                f"떠리몰 소스에 {letter} 열(0-based index {idx})이 없습니다. "
                f"소스 컬럼 수: {len(src_cols_by_index)}"
            )
        return src_cols_by_index[idx]

    col_order = resolve(TTARIMALL_FIXED_LETTER_MAPPING["주문번호"])
    col_name = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 이름"])
    col_addr = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 주소"])
    col_phone = resolve(TTARIMALL_FIXED_LETTER_MAPPING["받는분 전화번호"])
    col_v = resolve(TTARIMALL_FIXED_LETTER_MAPPING["상품명"])
    col_s = resolve("S")
    col_qty = resolve(TTARIMALL_FIXED_LETTER_MAPPING["수량"])
    col_memo = resolve(TTARIMALL_FIXED_LETTER_MAPPING["메모"])

    result = pd.DataFrame(index=range(len(df_tm)), columns=template_columns)
    result["주문번호"] = df_tm[col_order]
    result["받는분 이름"] = df_tm[col_name]
    result["받는분 주소"] = df_tm[col_addr]
    phone = df_tm[col_phone].astype(str)
    result["받는분 전화번호"] = phone.where(phone.str.lower() != "nan", "")

    s_raw = df_tm[col_s].astype(str)
    v_raw = df_tm[col_v].astype(str)
    s = s_raw.where(s_raw.str.lower() != "nan", "")
    v = v_raw.where(v_raw.str.lower() != "nan", "")
    same = (s == v)
    prod = v.copy()
    prod.loc[~same] = s[~same] + v[~same]
    result["상품명"] = prod

    result["수량"] = pd.to_numeric(df_tm[col_qty], errors="coerce")
    result["메모"] = df_tm[col_memo]
    return result

def post_numeric_alignment(result_df: pd.DataFrame):
    # 템플릿 숫자형 정렬(전화번호 제외)
    for col in template_columns:
        if col in result_df.columns and col in tpl_df.columns and tpl_df[col].notna().any():
            if pd.api.types.is_numeric_dtype(tpl_df[col]) and col != "받는분 전화번호":
                result_df[col] = pd.to_numeric(result_df[col], errors="coerce")

if run_batch:
    if not batch_files:
        st.error("엑셀 파일을 하나 이상 업로드해 주세요.")
    elif tpl_df is None or len(template_columns) == 0:
        st.error("유효한 템플릿이 필요합니다.")
    else:
        zip_buffer = io.BytesIO()
        logs = []
        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for f in batch_files:
                fname = getattr(f, "name", "uploaded.xlsx")
                try:
                    df = read_first_sheet_source_as_text(f)
                except Exception as e:
                    logs.append(f"[FAIL] {fname}: 파일 읽기 오류 - {e}")
                    continue

                platform = detect_platform_by_headers(df)
                try:
                    if platform == "TTARIMALL":
                        out_df = convert_ttarimall(df)
                    elif platform == "SMARTSTORE":
                        out_df = convert_smartstore_keywords(df)
                    elif platform == "COUPANG":
                        out_df = convert_coupang(df)
                    else:  # LAORA
                        out_df = convert_laora(df)
                    post_numeric_alignment(out_df)

                    # 파일별 엑셀 쓰기
                    xbuf = io.BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                        out_df_sorted = out_df[template_columns + [c for c in out_df.columns if c not in template_columns]]
                        out_df_sorted.to_excel(writer, index=False)
                    base = fname.rsplit(".", 1)[0]
                    out_name = f"{base}__{platform.lower()}_converted.xlsx"
                    zf.writestr(out_name, xbuf.getvalue())

                    logs.append(f"[OK]   {fname}: {platform} → rows={len(out_df)} → {out_name}")
                except Exception as e:
                    logs.append(f"[FAIL] {fname}: {platform} 처리 중 오류 - {e}")

            # 로그 파일 추가
            log_text = "Batch Convert Log - " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n" + "\n".join(logs)
            zf.writestr("batch_convert_log.txt", log_text)

        st.success("배치 변환이 완료되었습니다.")
        st.text_area("변환 로그", value="\n".join(logs), height=200)
        st.download_button(
            label="배치 변환 결과 ZIP 다운로드",
            data=zip_buffer.getvalue(),
            file_name=f"batch_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
        )

st.caption("라오라 / 쿠팡 / 스마트스토어(키워드) / 떠리몰(S&V) 외 양식도 추가 가능합니다. 규칙만 알려주시면 바로 넣어드릴게요.")

# ======================================================================
# 6) 송장등록: 송장파일(.xls/.xlsx) → 라오/스마트스토어/쿠팡 분류 & 생성
# ======================================================================

# 안전 로더 (.xls/.xlsx)
def _read_excel_any(file, header=0, dtype=str, keep_default_na=False) -> pd.DataFrame:
    """
    안전한 엑셀 로더 (.xlsx/.xls)
      - 업로드 바이트 확보 → BytesIO 로 매 시도마다 새로 읽음
      - .xlsx → openpyxl
      - .xls  → xlrd (권장 버전: 1.2.0)
    """
    name = (getattr(file, "name", "") or "").lower()

    data = None
    if hasattr(file, "getvalue"):
        try:
            data = file.getvalue()
        except Exception:
            data = None
    if data is None:
        try:
            cur = file.tell() if hasattr(file, "tell") else None
            if hasattr(file, "seek"):
                file.seek(0)
            data = file.read()
            if hasattr(file, "seek") and cur is not None:
                file.seek(cur)
        except Exception:
            data = None

    def _read_with(engine: Optional[str]):
        bio = io.BytesIO(data) if data is not None else file
        return pd.read_excel(
            bio, sheet_name=0, header=header, dtype=dtype,
            keep_default_na=keep_default_na, engine=engine,
        )

    try:
        if name.endswith(".xlsx"):
            return _read_with("openpyxl")
        elif name.endswith(".xls"):
            try:
                return _read_with("xlrd")
            except Exception as e:
                raise RuntimeError(
                    "'.xls' 파일을 읽으려면 xlrd가 필요합니다. 권장: pip install \"xlrd==1.2.0\"\n"
                    f"원본 오류: {e}"
                )
        else:
            try:
                return _read_with(None)
            except Exception:
                try:
                    return _read_with("openpyxl")
                except Exception:
                    try:
                        return _read_with("xlrd")
                    except Exception as e:
                        raise RuntimeError(
                            "엑셀 파일을 읽을 수 없습니다. (.xlsx는 openpyxl, .xls는 xlrd 필요)\n"
                            f"원본 오류: {e}"
                        )
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"엑셀 파일을 읽는 중 알 수 없는 오류: {e}")

# 숫자만 남기는 헬퍼 (쿠팡 매칭용)
def _digits_only(x: str) -> str:
    return re.sub(r"\D+", "", str(x or ""))

st.markdown("## 🚚 송장등록")

with st.expander("동작 요약", expanded=False):
    st.markdown(
        """
        - **분류 규칙**
          1) 주문번호에 **`LO`** 포함 → **라스트오더(라오)**
          2) (숫자 기준) **16자리** → **스마트스토어**
        - **라오 출력**: 템플릿 업로드 없이 고정 컬럼  
          **[`주문번호`, `택배사코드(08)`, `송장번호`]** → **라오 송장 완성_타임스탬프.xlsx**
        - **스마트스토어 출력**: 주문 파일과 **주문번호 매칭** → 송장번호 추가/갱신  
          (결과 **시트명: 배송처리**, `택배사` 기본값=**롯데택배**, 파일명에 타임스탬프)
        - **쿠팡 출력**: **송장파일의 P열(주문번호)** ↔ **쿠팡주문파일의 C열(주문번호)** 를  
          **숫자만 비교**하여 일치 시 **쿠팡주문파일 E열(운송장 번호)** 에 **송장파일의 송장번호** 입력
        """
    )

# 라오 고정 컬럼
LAO_FIXED_TEMPLATE_COLUMNS = ["주문번호", "택배사코드", "송장번호"]

st.subheader("1) 파일 업로드")
invoice_file = st.file_uploader("송장번호 포함 파일 업로드 (예: 송장파일.xls)", type=["xls", "xlsx"], key="inv_file")
ss_order_file = st.file_uploader("스마트스토어 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_ss_orders")
cp_order_file = st.file_uploader("쿠팡 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_cp_orders")

run_invoice = st.button("송장등록 실행")

# 헤더 후보
ORDER_KEYS_INVOICE = ["주문번호", "주문ID", "주문코드", "주문번호1"]
TRACKING_KEYS = ["송장번호", "운송장번호", "운송장", "등기번호", "운송장 번호", "송장번호1"]

SS_ORDER_KEYS = ["주문번호"]
SS_TRACKING_COL_NAME = "송장번호"

def build_order_tracking_map(df_invoice: pd.DataFrame):
    """송장파일에서 (주문번호 → 송장번호) 매핑 생성 (헤더명 기반)"""
    order_col = find_col(ORDER_KEYS_INVOICE, df_invoice)
    tracking_col = find_col(TRACKING_KEYS, df_invoice)
    orders = df_invoice[order_col].astype(str)
    tracks = df_invoice[tracking_col].astype(str)
    orders = orders.where(orders.str.lower() != "nan", "")
    tracks = tracks.where(tracks.str.lower() != "nan", "")
    mapping = {}
    for o, t in zip(orders, tracks):
        if o and t:
            mapping[str(o)] = str(t)
    return mapping

def classify_orders(mapping: dict):
    """
    분류:
      - 라오: 'LO' 포함
      - 스마트스토어: 숫자만 16자리
      (쿠팡은 자리수 무시 숫자매칭으로 별도 처리)
    """
    lao, ss = {}, {}
    for o, t in mapping.items():
        s = str(o).strip()
        if "LO" in s.upper():
            lao[s] = t
        elif len(_digits_only(s)) == 16:
            ss[s] = t
    return lao, ss

def make_lao_invoice_df_fixed(lao_map: dict) -> pd.DataFrame:
    """라오 송장: 고정 컬럼으로 DF 생성 (택배사코드=08, 컬럼 순서 고정)"""
    if not lao_map:
        return pd.DataFrame(columns=LAO_FIXED_TEMPLATE_COLUMNS)
    orders = list(lao_map.keys())
    tracks = [lao_map[o] for o in orders]
    out = pd.DataFrame(
        {"주문번호": orders, "택배사코드": ["08"] * len(orders), "송장번호": tracks},
        columns=LAO_FIXED_TEMPLATE_COLUMNS,
    )
    return out

def make_ss_filled_df(ss_map: dict, ss_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """스마트스토어 주문 파일에 송장번호를 매칭해 추가/갱신 (파일 없으면 2열 매핑만)"""
    if ss_df is None or ss_df.empty:
        if not ss_map:
            return pd.DataFrame()
        df = pd.DataFrame({"주문번호": list(ss_map.keys()), SS_TRACKING_COL_NAME: list(ss_map.values())})
        df["택배사"] = "롯데택배"
        return df

    col_order = find_col(SS_ORDER_KEYS, ss_df)
    out = ss_df.copy()
    if SS_TRACKING_COL_NAME not in out.columns:
        out[SS_TRACKING_COL_NAME] = ""

    existing = out[SS_TRACKING_COL_NAME].astype(str)
    is_empty = (existing.str.lower().eq("nan")) | (existing.str.strip().eq(""))
    mapped = out[col_order].astype(str).map(ss_map).fillna("")
    out.loc[is_empty, SS_TRACKING_COL_NAME] = mapped[is_empty]

    # 택배사 기본값=롯데택배
    if "택배사" not in out.columns:
        out["택배사"] = "롯데택배"
    else:
        ser = out["택배사"].astype(str)
        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
        out.loc[empty_mask, "택배사"] = "롯데택배"

    return out

# --- (쿠팡) 송장파일 P열 기반 매핑 생성: 키는 숫자만 ---
def build_inv_map_from_P(df_invoice: pd.DataFrame) -> dict:
    """
    송장파일: P열(주문번호) ↔ 송장번호(여러 헤더명 중 탐색) → {숫자키: 송장번호}
    """
    inv_cols = list(df_invoice.columns)
    try:
        inv_order_col = inv_cols[excel_col_to_index("P")]
    except Exception:
        raise RuntimeError("송장파일에 P열(주문번호)이 없습니다. 송장파일 양식을 확인해 주세요.")
    tracking_col = find_col(TRACKING_KEYS, df_invoice)

    orders = df_invoice[inv_order_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    tracks = df_invoice[tracking_col].astype(str).where(lambda s: s.str.lower() != "nan", "")

    inv_map = {}
    for o, t in zip(orders, tracks):
        key = _digits_only(o)
        if key and str(t):
            inv_map[key] = str(t)  # 중복 키는 마지막 값 우선
    return inv_map

def make_cp_filled_df_by_letters(df_invoice: Optional[pd.DataFrame],
                                 cp_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    쿠팡 송장등록:
      - 매칭 키: (숫자만 남긴) 송장파일의 **P열 주문번호** ↔ (숫자만 남긴) 쿠팡주문파일의 **C열 주문번호**
      - 쓰기 대상: 쿠팡주문파일의 **E열(운송장 번호)** ← 송장파일의 '송장번호' 값
      - 자리수/포맷 무시(숫자만 비교)
    """
    if cp_df is None or cp_df.empty:
        return pd.DataFrame()
    if df_invoice is None or df_invoice.empty:
        return cp_df

    inv_map = build_inv_map_from_P(df_invoice)

    cp_cols = list(cp_df.columns)
    try:
        cp_order_col = cp_cols[excel_col_to_index("C")]  # 매칭 키
    except Exception:
        raise RuntimeError("쿠팡 주문 파일에 C열(주문번호)이 없습니다. 쿠팡 주문파일 양식을 확인해 주세요.")
    try:
        cp_track_col = cp_cols[excel_col_to_index("E")]  # 쓰기 대상
    except Exception:
        cp_track_col = "운송장 번호"
        if cp_track_col not in cp_df.columns:
            cp_df = cp_df.copy()
            cp_df[cp_track_col] = ""
        cp_cols = list(cp_df.columns)

    out = cp_df.copy()
    cp_keys = out[cp_order_col].astype(str).map(_digits_only)
    mapped = cp_keys.map(inv_map)

    # 매칭된 행에만 덮어쓰기
    mask = mapped.notna() & mapped.astype(str).str.len().gt(0)
    out.loc[mask, cp_track_col] = mapped[mask]

    return out


if run_invoice:
    # NameError 방지용 초기화
    df_invoice = None
    df_ss_orders = None
    df_cp_orders = None

    if not invoice_file:
        st.error("송장번호가 포함된 송장파일을 업로드해 주세요. (예: 송장파일.xls)")
    else:
        # 1) 송장파일 읽기
        try:
            df_invoice = _read_excel_any(invoice_file, header=0, dtype=str, keep_default_na=False)
        except Exception as e:
            st.exception(RuntimeError(f"송장파일 읽기 오류: {e}"))
            df_invoice = None

        # 2) (선택) 스마트스토어/쿠팡 주문 파일 읽기
        if ss_order_file:
            try:
                df_ss_orders = read_first_sheet_source_as_text(ss_order_file)
            except Exception as e:
                st.warning(f"스마트스토어 주문 파일을 읽는 중 오류: {e}")
                df_ss_orders = None

        if cp_order_file:
            try:
                df_cp_orders = read_first_sheet_source_as_text(cp_order_file)
            except Exception as e:
                st.warning(f"쿠팡 주문 파일을 읽는 중 오류: {e}")
                df_cp_orders = None

        # 3) 처리
        if df_invoice is None:
            st.error("송장파일을 읽지 못했습니다. 파일 형식 및 내용(주문번호/송장번호 컬럼)을 확인해 주세요.")
        else:
            try:
                # (주문번호 → 송장번호) 매핑 & 분류(라오/스마트스토어만)
                order_track_map = build_order_tracking_map(df_invoice)
                lao_map, ss_map = classify_orders(order_track_map)

                # 결과 DF 생성
                lao_out_df = make_lao_invoice_df_fixed(lao_map)                 # 라오: 택배사코드=08, 컬럼 순서 고정
                ss_out_df = make_ss_filled_df(ss_map, df_ss_orders)             # 스마트스토어: 주문번호 매칭(+택배사 기본값)
                cp_out_df = make_cp_filled_df_by_letters(df_invoice, df_cp_orders)  # 쿠팡: P↔C(숫자비교), E열 채움

                # 쿠팡 업데이트 예정 건수(숫자비교 기준)
                cp_update_cnt = 0
                if df_cp_orders is not None and not df_cp_orders.empty:
                    try:
                        inv_map_tmp = build_inv_map_from_P(df_invoice)
                        cp_cols_tmp = list(df_cp_orders.columns)
                        cp_order_col_tmp = cp_cols_tmp[excel_col_to_index("C")]
                        mapped_tmp = df_cp_orders[cp_order_col_tmp].astype(str).map(_digits_only).map(inv_map_tmp)
                        cp_update_cnt = int((mapped_tmp.notna() & mapped_tmp.astype(str).str.len().gt(0)).sum())
                    except Exception:
                        cp_update_cnt = 0

                # 미리보기
                st.success(f"분류 완료: 라오 {len(lao_map)}건 / 스마트스토어 {len(ss_map)}건 / 쿠팡 업데이트 예정 {cp_update_cnt}건")
                with st.expander("라오 송장 미리보기", expanded=True):
                    st.dataframe(lao_out_df.head(50))
                with st.expander("스마트스토어 송장 미리보기 (시트명: 배송처리)", expanded=False):
                    st.dataframe(ss_out_df.head(50))
                with st.expander("쿠팡 송장 미리보기", expanded=False):
                    st.dataframe(cp_out_df.head(50))

                # 파일명 타임스탬프
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")

                # 라오 송장 완성.xlsx
                buf_lao = io.BytesIO()
                with pd.ExcelWriter(buf_lao, engine="openpyxl") as writer:
                    lao_out_df.to_excel(writer, index=False)
                st.download_button(
                    label="라오 송장 완성.xlsx 다운로드",
                    data=buf_lao.getvalue(),
                    file_name=f"라오 송장 완성_{ts}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # 스마트스토어 송장 완성.xlsx — 시트명: 배송처리 / 택배사=롯데택배 기본값
                if ss_out_df is not None and not ss_out_df.empty:
                    ss_out_export = ss_out_df.copy()
                    if "택배사" not in ss_out_export.columns:
                        ss_out_export["택배사"] = "롯데택배"
                    else:
                        ser = ss_out_export["택배사"].astype(str)
                        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
                        ss_out_export.loc[empty_mask, "택배사"] = "롯데택배"

                    buf_ss = io.BytesIO()
                    with pd.ExcelWriter(buf_ss, engine="openpyxl") as writer:
                        ss_out_export.to_excel(writer, index=False, sheet_name="배송처리")
                    st.download_button(
                        label="스마트스토어 송장 완성.xlsx 다운로드",
                        data=buf_ss.getvalue(),
                        file_name=f"스마트스토어 송장 완성_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                # 쿠팡 송장 완성.xlsx
                if cp_out_df is not None and not cp_out_df.empty:
                    buf_cp = io.BytesIO()
                    with pd.ExcelWriter(buf_cp, engine="openpyxl") as writer:
                        cp_out_df.to_excel(writer, index=False)
                    st.download_button(
                        label="쿠팡 송장 완성.xlsx 다운로드",
                        data=buf_cp.getvalue(),
                        file_name=f"쿠팡 송장 완성_{ts}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                if (ss_out_df is None or ss_out_df.empty) and (cp_out_df is None or cp_out_df.empty):
                    st.info("스마트스토어/쿠팡 대상 건이 없거나, 매칭할 주문 파일이 없어 생성 결과가 없습니다.")

            except Exception as e:
                st.exception(RuntimeError(f"송장등록 처리 중 오류: {e}"))
