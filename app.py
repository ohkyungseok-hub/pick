# app.py
# 사용법:
#   python app.py "원본.xlsx" "결과.xlsx"
#   python app.py "원본.xlsx" "결과.xlsx" --docx "결과.docx"
#   python app.py "원본.xlsx" "결과.xlsx" --docx "결과.docx" --skip-xlsx
#
# 요구 라이브러리:
#   pip install pandas openpyxl python-docx

import sys
import argparse
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.text import WD_BREAK
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, RGBColor


def excel_col_to_zero_index(col_letter: str) -> int:
    """Excel column letter (e.g., 'A', 'J') -> pandas zero-based index"""
    col_letter = col_letter.strip().upper()
    n = 0
    for ch in col_letter:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letter: {col_letter}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def build_picking_dataframe(src_path: str, colmap: dict) -> pd.DataFrame:
    """원본 엑셀 -> 피킹용 DF(주소 정렬 + 주소별 합계행 포함)"""
    df = pd.read_excel(src_path)

    need_keys = ["상품연동코드", "주문상품", "옵션", "주문수량", "주문회원", "주소", "주문요청사항"]
    idxs = [excel_col_to_zero_index(colmap[k]) for k in need_keys]

    df_sel = df.iloc[:, idxs].copy()
    df_sel.columns = need_keys

    # 정렬: 주소(오름), 상품연동코드(내림)
    df_sorted = df_sel.sort_values(
        by=["주소", "상품연동코드"],
        ascending=[True, False],
        kind="mergesort"
    )

    # 주소별 합계행 추가
    out_chunks = []
    for addr, g in df_sorted.groupby("주소", sort=False):
        out_chunks.append(g)

        subtotal = {c: "" for c in df_sorted.columns}
        subtotal["주문상품"] = "합계"
        subtotal_qty = pd.to_numeric(g["주문수량"], errors="coerce").fillna(0).sum()
        subtotal["주문수량"] = subtotal_qty
        subtotal["주소"] = addr
        out_chunks.append(pd.DataFrame([subtotal]))

    df_final = pd.concat(out_chunks, ignore_index=True)
    return df_final


def build_picking_xlsx(df_final: pd.DataFrame, out_path: str) -> None:
    """DF -> 피킹용 엑셀 저장 + 인쇄/서식 설정(openpyxl)"""
    df_final.to_excel(out_path, index=False)

    wb = load_workbook(out_path)
    ws = wb.active

    # 헤더: 굵게 + 줄바꿈
    header_font = Font(bold=True)
    header_align = Alignment(wrap_text=True, vertical="center")
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.font = header_font
        cell.alignment = header_align

    headers = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    addr_col = headers["주소"]

    # 긴 텍스트 줄바꿈 + 위쪽 정렬
    wrap_top = Alignment(wrap_text=True, vertical="top")
    for r in range(2, ws.max_row + 1):
        for name in ["주문상품", "옵션", "주소", "주문요청사항"]:
            ws.cell(r, headers[name]).alignment = wrap_top

    # 열 너비(기존 유지)
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

    # 인쇄 설정(기존 유지: 가로)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    # 주소 바뀔 때마다 페이지 나누기
    ws.row_breaks.brk = []
    if ws.max_row >= 2:
        prev_addr = ws.cell(2, addr_col).value
        for r in range(3, ws.max_row + 1):
            curr_addr = ws.cell(r, addr_col).value
            if curr_addr != prev_addr:
                ws.row_breaks.append(Break(id=r - 1))  # 이전 행 뒤에서 끊기
                prev_addr = curr_addr

    ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    wb.save(out_path)


# -----------------------
# 워드 출력 (세로모드 최종)
# -----------------------
def _docx_set_row_height(row, pt: int = 26) -> None:
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(pt * 20)))  # twips
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)

    # 셀 내부 문단 여백 최소화
    for cell in row.cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)


def _docx_shade_row(row, fill: str = "EFEFEF") -> None:
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)


def build_picking_docx(df_final: pd.DataFrame, out_docx: str) -> None:
    """
    DF -> 피킹용 워드(.docx)
    - A4 세로
    - 행높이 26pt(정확히)
    - 주소별 1페이지
    - 코드 변경 시 음영 토글
    - 폰트 규칙(최종 확정):
        상단 주소 10
        표안 주소열 5
        주문상품 8, 옵션 8
        주문수량 12(2이상 빨강)
        상품연동코드 14 Bold
        합계행 16 Bold + 합계행 주소칸 비움
    """
    required_cols = ["상품연동코드", "주문상품", "옵션", "주문수량", "주문회원", "주소", "주문요청사항"]
    for c in required_cols:
        if c not in df_final.columns:
            raise ValueError(f"df_final에 '{c}' 컬럼이 없습니다. 현재 컬럼: {list(df_final.columns)}")

    doc = Document()

    # 페이지(세로) + 여백
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)   # A4
    section.page_height = Inches(11.69) # A4
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)

    # 기본 폰트
    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "맑은 고딕")
    style.font.size = Pt(9)

    # 주소별로 끊기(주소가 변경될 때마다 한 페이지)
    groups = []
    current_addr = None
    current_rows = []

    for _, row in df_final.iterrows():
        addr = "" if pd.isna(row["주소"]) else str(row["주소"]).strip()
        if current_addr is None:
            current_addr = addr
            current_rows = [row]
        elif addr == current_addr:
            current_rows.append(row)
        else:
            groups.append((current_addr, current_rows))
            current_addr = addr
            current_rows = [row]
    if current_rows:
        groups.append((current_addr, current_rows))

    # 컬럼 순서 고정
    cols = required_cols[:]

    # 세로모드 열 너비(인치)
    col_widths = {
        "상품연동코드": Inches(0.8),
        "주문상품": Inches(2.4),
        "옵션": Inches(1.4),
        "주문수량": Inches(0.6),
        "주문회원": Inches(1.0),
        "주소": Inches(1.2),
        "주문요청사항": Inches(1.1),
    }

    for gi, (addr, rows_in_addr) in enumerate(groups):
        # 상단 주소(10pt)
        p = doc.add_paragraph(f"주소: {addr}")
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(10)
        p.paragraph_format.space_after = Pt(4)

        table = doc.add_table(rows=1, cols=len(cols))
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.allow_autofit = False

        # 헤더행(8pt) + 행높이 26
        hdr = table.rows[0]
        _docx_set_row_height(hdr, 26)
        for ci, name in enumerate(cols):
            cell = hdr.cells[ci]
            cell.text = name
            cell.width = col_widths[name]
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].font.size = Pt(8)

        # 코드 변경 시 음영 토글
        last_code = None
        shade_on = False

        for r in rows_in_addr:
            is_sum = (str(r.get("주문상품", "")) == "합계") or ("합계" in str(r.get("주문상품", "")))

            code_val = "" if pd.isna(r["상품연동코드"]) else str(r["상품연동코드"])
            if not is_sum:
                if code_val != last_code:
                    shade_on = not shade_on
                    last_code = code_val

            row = table.add_row()
            _docx_set_row_height(row, 26)

            if shade_on and not is_sum:
                _docx_shade_row(row, "EFEFEF")

            for ci, name in enumerate(cols):
                cell = row.cells[ci]
                cell.width = col_widths[name]

                # 합계행: 주소칸 비움
                if is_sum and name == "주소":
                    cell.text = ""
                    continue

                val = r.get(name, "")
                text = "" if pd.isna(val) else str(val)

                run = cell.paragraphs[0].add_run(text)

                if name == "주소":
                    run.font.size = Pt(5)
                elif name in ("주문상품", "옵션"):
                    run.font.size = Pt(8)
                elif name == "주문수량":
                    run.font.size = Pt(12)
                    try:
                        if int(float(text)) >= 2:
                            run.font.color.rgb = RGBColor(255, 0, 0)
                    except:
                        pass
                elif name == "상품연동코드":
                    run.font.size = Pt(14)
                    run.bold = True
                else:
                    run.font.size = Pt(8)

                if is_sum:
                    run.font.size = Pt(16)
                    run.bold = True

        if gi != len(groups) - 1:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    doc.save(out_docx)


def build_picking_sheet(
    src_path: str,
    out_xlsx_path: str,
    colmap=None,
    out_docx_path: str | None = None,
    skip_xlsx: bool = False,
):
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

    df_final = build_picking_dataframe(src_path, colmap)

    if not skip_xlsx:
        build_picking_xlsx(df_final, out_xlsx_path)

    if out_docx_path:
        build_picking_docx(df_final, out_docx_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("src", help='원본 엑셀 경로 (예: "원본.xlsx")')
    parser.add_argument("out_xlsx", help='결과 엑셀 경로 (예: "결과.xlsx")')
    parser.add_argument("--docx", dest="out_docx", default=None, help='결과 워드 경로 (예: "결과.docx")')
    parser.add_argument("--skip-xlsx", action="store_true", help="엑셀 저장 생략(워드만 생성할 때)")
    args = parser.parse_args()

    build_picking_sheet(
        src_path=args.src,
        out_xlsx_path=args.out_xlsx,
        out_docx_path=args.out_docx,
        skip_xlsx=args.skip_xlsx,
    )

    if not args.skip_xlsx:
        print(f"엑셀 완료: {args.out_xlsx}")
    if args.out_docx:
        print(f"워드 완료: {args.out_docx}")


if __name__ == "__main__":
    main()
