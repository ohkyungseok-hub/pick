# app_upload_fix.py
# 실행: streamlit run app_upload_fix.py
# 필요: pip install streamlit pandas openpyxl
# (.xls 읽기 필요 시) pip install "xlrd==1.2.0"

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

st.set_page_config(page_title="송장등록", layout="centered")

st.title("송장등록")
st.caption("송장번호를 라오/스마트스토어/쿠팡/떠리몰 형식으로 등록합니다.")

# -------------------------- Helpers --------------------------
def excel_col_to_index(col_letters: str) -> int:
    col_letters = str(col_letters).strip().upper()
    if not re.fullmatch(r"[A-Z]+", col_letters):
        raise ValueError(f"Invalid Excel column letters: {col_letters}")
    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1  # 0-based

def norm_header(s: str) -> str:
    return re.sub(r"[\s\(\)\[\]{}:：/\\\-]", "", str(s).strip().lower())

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

def read_first_sheet_source_as_text(file) -> pd.DataFrame:
    """전 컬럼 문자열로 읽어 전화번호 앞 0 보존"""
    return pd.read_excel(
        file, sheet_name=0, header=0, engine="openpyxl",
        dtype=str, keep_default_na=False,
    )

# Excel이 CSV를 열 때 숫자로 오인되지 않도록 텍스트 보호
def _guard_excel_text(s: str) -> str:
    s = "" if s is None else str(s)
    if s == "" or s.startswith('="'):
        return s
    return f'="{s}"'

# -------------------- CSV 출력 설정(구분자/인코딩) --------------------
CSV_SEPARATORS = {"쉼표(,)": ",", "세미콜론(;)": ";", "탭(\\t)": "\t", "파이프(|)": "|"}
CSV_ENCODINGS = {
    "UTF-8-SIG (권장)": "utf-8-sig",
    "UTF-8 (BOM 없음)": "utf-8",
    "CP949 (윈도우)": "cp949",
    "EUC-KR": "euc-kr",
}

def _get_csv_prefs():
    # 기본 CP949, 쉼표
    sep = st.session_state.get("csv_sep", ",")
    enc = st.session_state.get("csv_encoding", "cp949")
    label_sep = st.session_state.get("csv_sep_label", "쉼표(,)")
    label_enc = st.session_state.get("csv_enc_label", "CP949 (윈도우)")
    return sep, enc, label_sep, label_enc

def download_df(
    df: pd.DataFrame,
    base_label: str,
    filename_stem: str,
    widget_key: str,
    sheet_name: Optional[str] = None,
    csv_sep_override: Optional[str] = None,
    csv_encoding_override: Optional[str] = None,
):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    col_csv, col_xlsx = st.columns(2)

    def _labels_from_sep(sep: str) -> str:
        return {",": "쉼표(,)", ";": "세미콜론(;)", "\t": "탭(\\t)", "|": "파이프(|)"}.get(sep, f"사용자({repr(sep)})")

    def _labels_from_enc(enc: str) -> str:
        rev = {v: k for k, v in CSV_ENCODINGS.items()}
        return rev.get(enc, enc)

    default_sep, default_enc, _, _ = _get_csv_prefs()
    csv_sep = csv_sep_override if csv_sep_override is not None else default_sep
    csv_enc = csv_encoding_override if csv_encoding_override is not None else default_enc
    label_sep = _labels_from_sep(csv_sep)
    label_enc = _labels_from_enc(csv_enc)

    # CSV (전화번호 보호)
    with col_csv:
        df_safe = df.copy()
        phone_like_cols = [c for c in df_safe.columns if re.search(r"(전화번호|연락처|휴대폰)", str(c))]
        for c in phone_like_cols:
            df_safe[c] = df_safe[c].astype(str).map(_guard_excel_text)

        csv_str = df_safe.to_csv(index=False, sep=csv_sep, lineterminator="\n")
        csv_bytes = csv_str.encode(csv_enc, errors="replace")
        st.download_button(
            label=f"{base_label} (CSV · {label_sep} · {label_enc})",
            data=csv_bytes,
            file_name=f"{filename_stem}_{ts}.csv",
            mime="text/csv",
            key=f"btn_{widget_key}_csv",
            help="선택한/강제된 구분자·인코딩으로 CSV 저장합니다.",
        )

    # XLSX
    with col_xlsx:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            if sheet_name:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
            else:
                df.to_excel(writer, index=False)
        st.download_button(
            label=f"{base_label} (XLSX)",
            data=buf.getvalue(),
            file_name=f"{filename_stem}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"btn_{widget_key}_xlsx",
            help="서식 유지가 필요할 때 XLSX로 저장하세요.",
        )

# ======================================================================
# 송장등록: 송장파일 → 라오/스마트스토어/쿠팡/떠리몰
# ======================================================================

def _get_bytes(file) -> bytes:
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
    if data is None:
        raise RuntimeError("업로드 파일 바이트를 읽을 수 없습니다.")
    return data

def _read_excel_any(file, header=0, dtype=str, keep_default_na=False) -> pd.DataFrame:
    name = (getattr(file, "name", "") or "").lower()
    data = _get_bytes(file)

    def _read_with(engine: Optional[str]):
        bio = io.BytesIO(data)
        return pd.read_excel(bio, sheet_name=0, header=header, dtype=dtype, keep_default_na=keep_default_na, engine=engine)

    try:
        if name.endswith(".xlsx"):
            return _read_with("openpyxl")
        elif name.endswith(".xls"):
            try:
                return _read_with("xlrd")
            except Exception as e:
                raise RuntimeError("'.xls' 파일을 읽으려면 xlrd가 필요합니다. 권장: pip install \"xlrd==1.2.0\"; 원본 오류: "+str(e))
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
                        raise RuntimeError("엑셀 파일을 읽을 수 없습니다. (.xlsx는 openpyxl, .xls는 xlrd 필요) 원본 오류: "+str(e))
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"엑셀 파일을 읽는 중 알 수 없는 오류: {e}")

def _digits_only(x: str) -> str:
    return re.sub(r"\D+", "", str(x or ""))

st.markdown("## 🚚 new 송장등록")

with st.expander("동작 요약", expanded=False):
    st.markdown(
        """
        - **송장파일 형식**: 주문번호/고객주문번호와 운송장번호 컬럼을 자동으로 찾아 처리합니다.
        - **분류 규칙**
          1) 주문번호에 **`LO`** 포함 → **라스트오더(라오)**
          2) (숫자 기준) **16자리** → **스마트스토어**
        - **라오 출력**: [`주문번호`, `택배사코드(04)`, `송장번호`]
        - **스마트스토어 출력**: 주문 파일과 주문번호 매칭 → 송장번호 추가/갱신  
          (결과 **시트명: 발송처리**, `택배사` 기본값=**CJ대한통운**)
        - **쿠팡 출력**: 송장 주문번호(**P열 또는 헤더 자동탐색**) ↔ 쿠팡 C열(숫자만 비교) 일치 시 E열에 입력
        - **떠리몰 출력(키워드)**: 주문번호 매칭 후 송장번호 자동 기입
        """
    )

LAO_FIXED_TEMPLATE_COLUMNS = ["주문번호", "택배사코드", "송장번호"]

st.subheader("1) 파일 업로드")
invoice_file = st.file_uploader("송장번호 포함 파일 업로드 (예: 송장파일.xls)", type=["xls", "xlsx"], key="inv_file")
ss_order_file = st.file_uploader("스마트스토어 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_ss_orders")
cp_order_file = st.file_uploader("쿠팡 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_cp_orders")
tm_order_file = st.file_uploader("떠리몰 주문 파일 업로드 (선택)", type=["xlsx"], key="inv_tm_orders")

run_invoice = st.button("송장등록 실행")

ORDER_KEYS_INVOICE = ["주문번호", "주문ID", "주문코드", "주문번호1", "고객주문번호"]
TRACKING_KEYS = ["송장번호", "운송장번호", "운송장", "등기번호", "운송장 번호", "송장번호1"]

SS_ORDER_KEYS = ["주문번호"]
SS_TRACKING_COL_NAME = "송장번호"
TM_ORDER_KEYS = ["주문번호", "주문ID", "주문코드", "주문번호1"]

def build_order_tracking_map(df_invoice: pd.DataFrame):
    order_col = find_col(ORDER_KEYS_INVOICE, df_invoice)
    tracking_col = find_col(TRACKING_KEYS, df_invoice)
    orders = df_invoice[order_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    tracks = df_invoice[tracking_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    mapping = {}
    for o, t in zip(orders, tracks):
        if o and t:
            mapping[str(o)] = str(t)
    return mapping

def classify_orders(mapping: dict):
    lao, ss = {}, {}
    for o, t in mapping.items():
        s = str(o).strip()
        if "LO" in s.upper():
            lao[s] = t
        elif len(_digits_only(s)) == 16:
            ss[s] = t
    return lao, ss

def make_lao_invoice_df_fixed(lao_map: dict) -> pd.DataFrame:
    if not lao_map:
        return pd.DataFrame(columns=LAO_FIXED_TEMPLATE_COLUMNS)
    orders = list(lao_map.keys())
    tracks = [lao_map[o] for o in orders]
    return pd.DataFrame({"주문번호": orders, "택배사코드": ["04"] * len(orders), "송장번호": tracks}, columns=LAO_FIXED_TEMPLATE_COLUMNS)

def make_ss_filled_df(ss_map: dict, ss_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if ss_df is None or ss_df.empty:
        if not ss_map:
            return pd.DataFrame()
        df = pd.DataFrame({"주문번호": list(ss_map.keys()), SS_TRACKING_COL_NAME: list(ss_map.values())})
        df["택배사"] = "CJ대한통운"
        return df
    col_order = find_col(SS_ORDER_KEYS, ss_df)
    out = ss_df.copy()
    if SS_TRACKING_COL_NAME not in out.columns:
        out[SS_TRACKING_COL_NAME] = ""
    existing = out[SS_TRACKING_COL_NAME].astype(str)
    is_empty = (existing.str.lower().eq("nan")) | (existing.str.strip().eq(""))
    mapped = out[col_order].astype(str).map(ss_map).fillna("")
    out.loc[is_empty, SS_TRACKING_COL_NAME] = mapped[is_empty]
    if "택배사" not in out.columns:
        out["택배사"] = "CJ대한통운"
    else:
        ser = out["택배사"].astype(str)
        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
        out.loc[empty_mask, "택배사"] = "CJ대한통운"
    return out

# --- (쿠팡) 송장파일에서 주문번호 매핑 생성: P열 우선, 없으면 헤더 자동탐색 ---
def build_inv_map_from_P(df_invoice: pd.DataFrame) -> dict:
    """
    송장파일: (우선) P열(주문번호) 또는 (대안) 헤더 키워드(ORDER_KEYS_INVOICE)로 주문번호 열을 찾아
    송장번호(TRACKING_KEYS)와 매핑을 만든다. 반환: {숫자만 남긴 주문번호: 송장번호}
    """
    inv_cols = list(df_invoice.columns)
    tracking_col = find_col(TRACKING_KEYS, df_invoice)
    try:
        inv_order_col = inv_cols[excel_col_to_index("P")]
    except Exception:
        try:
            inv_order_col = find_col(ORDER_KEYS_INVOICE, df_invoice)
        except Exception:
            raise RuntimeError("송장파일에서 주문번호 열을 찾지 못했습니다. (P열 또는 헤더: 주문번호/주문ID/주문코드/주문번호1)")
    orders = df_invoice[inv_order_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    tracks = df_invoice[tracking_col].astype(str).where(lambda s: s.str.lower() != "nan", "")
    inv_map = {}
    for o, t in zip(orders, tracks):
        key = _digits_only(o)
        if key and str(t):
            inv_map[key] = str(t)
    return inv_map

def make_cp_filled_df_by_letters(df_invoice: Optional[pd.DataFrame], cp_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if cp_df is None or cp_df.empty:
        return pd.DataFrame()
    if df_invoice is None or df_invoice.empty:
        return cp_df
    inv_map = build_inv_map_from_P(df_invoice)
    cp_cols = list(cp_df.columns)
    try:
        cp_order_col = cp_cols[excel_col_to_index("C")]
    except Exception:
        raise RuntimeError("쿠팡 주문 파일에 C열(주문번호)이 없습니다.")
    try:
        cp_track_col = cp_cols[excel_col_to_index("E")]
    except Exception:
        cp_track_col = "운송장 번호"
        if cp_track_col not in cp_df.columns:
            cp_df = cp_df.copy()
            cp_df[cp_track_col] = ""
    out = cp_df.copy()
    cp_keys = out[cp_order_col].astype(str).map(_digits_only)
    mapped = cp_keys.map(inv_map)
    mask = mapped.notna() & mapped.astype(str).str.len().gt(0)
    out.loc[mask, cp_track_col] = mapped[mask]
    return out

def make_tm_filled_df(tm_df: Optional[pd.DataFrame], inv_map: dict) -> pd.DataFrame:
    if tm_df is None or tm_df.empty:
        return pd.DataFrame()
    tm_order_col = find_col(TM_ORDER_KEYS, tm_df)
    tracking_col_candidates = [c for c in TRACKING_KEYS if c in list(tm_df.columns)]
    if tracking_col_candidates:
        tm_tracking_col = tracking_col_candidates[0]
        out = tm_df.copy()
    else:
        tm_tracking_col = "송장번호"
        out = tm_df.copy()
        if tm_tracking_col not in out.columns:
            out[tm_tracking_col] = ""
    keys = out[tm_order_col].astype(str)
    mapped = keys.map(inv_map)
    mask = mapped.notna() & mapped.astype(str).str.len().gt(0)
    out.loc[mask, tm_tracking_col] = mapped[mask]
    return out

if run_invoice:
    df_invoice = None
    df_ss_orders = None
    df_cp_orders = None
    df_tm_orders = None

    if not invoice_file:
        st.error("송장번호가 포함된 송장파일을 업로드해 주세요. (예: 송장파일.xls)")
    else:
        try:
            df_invoice = _read_excel_any(invoice_file, header=0, dtype=str, keep_default_na=False)
        except Exception as e:
            st.exception(RuntimeError(f"송장파일 읽기 오류: {e}"))
            df_invoice = None

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

        if tm_order_file:
            try:
                df_tm_orders = read_first_sheet_source_as_text(tm_order_file)
            except Exception as e:
                st.warning(f"떠리몰 주문 파일을 읽는 중 오류: {e}")
                df_tm_orders = None

        if df_invoice is None:
            st.error("송장파일을 읽지 못했습니다. 파일 형식 및 내용(주문번호/송장번호 컬럼)을 확인해 주세요.")
        else:
            try:
                order_track_map = build_order_tracking_map(df_invoice)
                lao_map, ss_map = classify_orders(order_track_map)

                lao_out_df = make_lao_invoice_df_fixed(lao_map)
                ss_out_df = make_ss_filled_df(ss_map, df_ss_orders)
                cp_out_df = make_cp_filled_df_by_letters(df_invoice, df_cp_orders)
                tm_out_df = make_tm_filled_df(df_tm_orders, order_track_map)

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

                tm_update_cnt = 0
                if df_tm_orders is not None and not df_tm_orders.empty and tm_out_df is not None and not tm_out_df.empty:
                    try:
                        tm_track_col = next((c for c in TRACKING_KEYS if c in tm_out_df.columns), "송장번호")
                        before = df_tm_orders.get(tm_track_col, pd.Series([""]*len(df_tm_orders))).astype(str).fillna("")
                        after  = tm_out_df.get(tm_track_col, pd.Series([""]*len(tm_out_df))).astype(str).fillna("")
                        tm_update_cnt = int((before != after).sum())
                    except Exception:
                        tm_update_cnt = 0

                st.success(f"분류/매칭 완료: 라오 {len(lao_map)}건 / 스마트스토어 {len(ss_map)}건 / 쿠팡 업데이트 예정 {cp_update_cnt}건 / 떠리몰 갱신 {tm_update_cnt}건")
                with st.expander("라오 송장 미리보기", expanded=True):
                    st.dataframe(lao_out_df.head(50))
                with st.expander("스마트스토어 송장 미리보기 (시트명: 발송처리)", expanded=False):
                    st.dataframe(ss_out_df.head(50))
                with st.expander("쿠팡 송장 미리보기", expanded=False):
                    st.dataframe(cp_out_df.head(50))
                with st.expander("떠리몰 송장 미리보기", expanded=False):
                    st.dataframe(tm_out_df.head(50))

                # 다운로드 (CSV 전부 CP949)
                download_df(lao_out_df, "라오 송장 완성 다운로드", "라오 송장 완성", "lao_inv",
                            csv_encoding_override="cp949")
                if ss_out_df is not None and not ss_out_df.empty:
                    ss_out_export = ss_out_df.copy()
                    if "택배사" not in ss_out_export.columns:
                        ss_out_export["택배사"] = "CJ대한통운"
                    else:
                        ser = ss_out_export["택배사"].astype(str)
                        empty_mask = ser.str.lower().eq("nan") | ser.str.strip().eq("")
                        ss_out_export.loc[empty_mask, "택배사"] = "CJ대한통운"
                    download_df(ss_out_export, "스마트스토어 송장 완성 다운로드", "스마트스토어 송장 완성", "ss_inv",
                                sheet_name="발송처리", csv_sep_override=",", csv_encoding_override="cp949")
                if cp_out_df is not None and not cp_out_df.empty:
                    download_df(cp_out_df, "쿠팡 송장 완성 다운로드", "쿠팡 송장 완성", "cp_inv",
                                csv_encoding_override="cp949")
                if tm_out_df is not None and not tm_out_df.empty:
                    download_df(tm_out_df, "떠리몰 송장 완성 다운로드", "떠리몰 송장 완성", "tm_inv",
                                csv_encoding_override="cp949")

                if (ss_out_df is None or ss_out_df.empty) and (cp_out_df is None or cp_out_df.empty) and (tm_out_df is None or tm_out_df.empty):
                    st.info("스마트스토어/쿠팡/떠리몰 대상 건이 없거나, 매칭할 주문 파일이 없어 생성 결과가 없습니다.")

            except Exception as e:
                st.exception(RuntimeError(f"송장등록 처리 중 오류: {e}"))
