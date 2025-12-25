import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter


def excel_col_to_zero_index(col_letter: str) -> int:
    col_letter = col_letter.strip().upper()
    n = 0
    for ch in col_letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letter: {col_letter}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def normalize_address(s: pd.Series) -> pd.Series:
    s = s.fillna("").astype(str)
    s = s.str.replace("\n", " ", regex=False).str.strip()
    s = s.str.replace(r"\s+", " ", regex=True)
    return s


def build_picking_sheet(src_path: str, out_path: str, colmap=None):
    if colmap is None:
        colmap = {
            "상품연동코드": "J",
            "주문상품": "K",
            "옵션": "L",
            "주문수량": "N",
            "주문회원": "Q",
            "주소": "V",
            "주문요청사항": "W",
        }

    df = pd.read_excel(src_path)

    needed = ["상품연동코드", "주문상품", "옵션", "주문수량", "주문회원", "주소", "주문요청사항"]
    idxs = [excel_col_to_zero_index(colmap[k]) for k in needed]

    max_idx = max(idxs)
    if df.shape[1] <= max_idx:
        raise ValueError(
            f"원본 파일 열 수({df.shape[1]})가 부족합니다. "
            f"필요한 최대 열: {get_column_letter(max_idx+1)}"
        )

    df_sel = df.iloc[:, idxs].copy()
    df_sel.columns = needed

    df_sel["주소"] = normalize_address(df_sel["주소"])

    df_sorted = df_sel.sort_values(
        by=["주소", "상품연동코드"],
        ascending=[True, False],
        kind="mergesort",
    )

    out_chunks = []
    for addr, g in df_sorted.groupby("주소", sort=False, dropna=False):
        out_chunks.append(g)

        subtotal = {c: "" for c in df_sorted.columns}
        subtotal["주문상품"] = "합계"
        qty = pd.to_numeric(g["주문수량"], errors="coerce").fillna(0).sum()
        qty = int(qty) if float(qty).is_integer() else float(qty)

        subtotal["주문수량"] = qty
        subtotal["주소"] = addr
        out_chunks.append(pd.DataFrame([subtotal]))

    df_final = pd.concat(out_chunks, ignore_index=True)
    df_final.to_excel(out_path, index=False)

    # 서식/인쇄 설정
    wb = load_workbook(out_path)
    ws = wb.active

    header_font = Font(bold=True)
    header_align = Alignment(wrap_text=True, vertical="center")
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.alignment = header_align

    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    addr_col = headers["주소"]

    wrap_top = Alignment(wrap_text=True, vertical="top")
    for r in range(2, ws.max_row + 1):
        for name in ["주문상품", "옵션", "주소", "주문요청사항"]:
            ws.cell(r, headers[name]).alignment = wrap_top

    widths = {
        "상품연동코드": 18,
        "주문상품": 60,
        "옵션": 50,
        "주문수량": 10,
        "주문회원": 18,
        "주소": 50,
        "주문요청사항": 40,
    }
    for name, w in widths.items():
        ws.column_dimensions[get_column_letter(headers[name])].width = w

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    ws.row_breaks.brk = []
    if ws.max_row >= 2:
        prev_addr = ws.cell(2, addr_col).value
        for r in range(3, ws.max_row + 1):
            curr_addr = ws.cell(r, addr_col).value
            if curr_addr != prev_addr:
                ws.row_breaks.append(Break(id=r - 1))
                prev_addr = curr_addr

    ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    wb.save(out_path)


# ---------------- Streamlit UI ----------------

st.set_page_config(page_title="피킹시트 생성기", layout="centered")
st.title("📦 피킹시트 생성기")
st.caption("엑셀 업로드 → 주소별 정렬/합계/페이지나누기 적용 → 결과 엑셀 다운로드")

with st.expander("원본 컬럼 위치 설정(기본값: J,K,L,N,Q,V,W)", expanded=False):
    colmap = {
        "상품연동코드": st.text_input("상품연동코드 컬럼(예: J)", value="J"),
        "주문상품": st.text_input("주문상품 컬럼(예: K)", value="K"),
        "옵션": st.text_input("옵션 컬럼(예: L)", value="L"),
        "주문수량": st.text_input("주문수량 컬럼(예: N)", value="N"),
        "주문회원": st.text_input("주문회원 컬럼(예: Q)", value="Q"),
        "주소": st.text_input("주소 컬럼(예: V)", value="V"),
        "주문요청사항": st.text_input("주문요청사항 컬럼(예: W)", value="W"),
    }

uploaded = st.file_uploader("원본 엑셀(.xlsx)을 업로드하세요", type=["xlsx"])

if uploaded is not None:
    st.info(f"업로드 파일: {uploaded.name}")

    out_name = st.text_input("결과 파일명", value=f"picking_{Path(uploaded.name).stem}.xlsx")

    if st.button("✅ 피킹시트 만들기", use_container_width=True):
        try:
            with st.spinner("처리 중..."):
                with tempfile.TemporaryDirectory() as td:
                    src_path = Path(td) / "src.xlsx"
                    out_path = Path(td) / "out.xlsx"

                    # 업로드 파일 저장
                    src_path.write_bytes(uploaded.getbuffer())

                    # 생성
                    build_picking_sheet(str(src_path), str(out_path), colmap=colmap)

                    # 결과를 bytes로 읽기
                    data = out_path.read_bytes()

            st.success("완료! 아래 버튼으로 다운로드하세요.")
            st.download_button(
                label="⬇️ 결과 엑셀 다운로드",
                data=data,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error("처리 중 오류가 발생했습니다.")
            st.exception(e)
else:
    st.warning("엑셀 파일을 업로드하면 시작할 수 있어요.")
