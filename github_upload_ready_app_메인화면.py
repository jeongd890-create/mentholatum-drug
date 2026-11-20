import importlib.util
import os
import sys
import streamlit as st

st.set_page_config(
    page_title="거래명세표 PDF 일괄 생성 시스템",
    page_icon="📄",
    layout="wide",
)


def _load_pdf_engine():
    """
    pdf변환_엔진.py 또는 기존 pdf_generator 계열 모듈을 우선 탐색하여 로드하고,
    없을 경우 내장 엔진 소스로 자동 실행하여 파일명 변경에 관계없이 동작하도록 보장합니다.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "pdf변환_엔진.py"),
        os.path.join(base_dir, "github_upload_ready_pdf변환_엔진.py"),
        os.path.join(base_dir, "pdf_generator.py"),
        os.path.join(base_dir, "gemini-code-pdf_generator.py"),
    ]
    for path in candidates:
        if os.path.exists(path):
            spec = importlib.util.spec_from_file_location("pdf_generator", path)
            module = importlib.util.module_from_spec(spec)
            sys.modules["pdf_generator"] = module
            spec.loader.exec_module(module)
            return module

    fallback_path = os.path.join(base_dir, "pdf변환_엔진.py")
    embedded_source = """import datetime
import io
import math
import os
import re
import subprocess
import tempfile
import zipfile
from typing import Dict, List, Tuple

import openpyxl
import pandas as pd


def parse_yyyymmdd(val) -> datetime.datetime:
    \"\"\"수주일자/납품일자(20261120 또는 datetime)를 datetime 객체로 변환\"\"\"
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime(val.year, val.month, val.day)
    s = re.sub(r"[^0-9]", "", str(val))
    if len(s) >= 8:
        return datetime.datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))
    return datetime.datetime.today()


def load_orders_from_v7(excel_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame]:
    \"\"\"
    거래처별_주문입력_및_수주업로드_v7.xlsx에서
    1) 백제약품_주문입력, 지오영_주문입력, 타거래처_주문입력 시트의 양수 주문 직접 추출
    2) 수주업로드_서식 시트 캐시값과 교차 검증용 데이터 반환
    \"\"\"
    wb_val = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)

    ws_code = wb_val["거래처코드"]
    bill_to_name_map: Dict[int, str] = {}
    ship_to_center_map: Dict[int, str] = {}
    for r in range(2, ws_code.max_row + 1):
        cust_nm = ws_code.cell(r, 2).value
        b_cd = ws_code.cell(r, 3).value
        center_nm = ws_code.cell(r, 5).value
        s_cd = ws_code.cell(r, 6).value
        if b_cd is not None and cust_nm:
            b_int = int(float(b_cd))
            if b_int == 86030001:
                bill_to_name_map[b_int] = "백제약품"
            elif b_int == 86049000:
                bill_to_name_map[b_int] = "지오영"
            else:
                bill_to_name_map[b_int] = str(cust_nm).strip()
        if s_cd is not None and center_nm:
            ship_to_center_map[int(float(s_cd))] = str(center_nm).strip()

    ws_price = wb_val["단가정리"]
    pack_map: Dict[str, int] = {}
    baekje_price_map: Dict[str, int] = {}
    geoyoung_price_map: Dict[str, int] = {}
    other_price_map: Dict[Tuple[str, str], int] = {}

    for r in range(2, ws_price.max_row + 1):
        bc = ws_price.cell(r, 1).value
        if bc:
            bc_s = str(bc).strip()
            pack_map[bc_s] = int(ws_price.cell(r, 3).value or 1)
            baekje_price_map[bc_s] = int(ws_price.cell(r, 4).value or 0)
        gc = ws_price.cell(r, 6).value
        if gc:
            gc_s = str(gc).strip()
            pack_map[gc_s] = int(ws_price.cell(r, 8).value or 1)
            geoyoung_price_map[gc_s] = int(ws_price.cell(r, 9).value or 0)
        ocust = ws_price.cell(r, 11).value
        ocode = ws_price.cell(r, 12).value
        if ocust and ocode:
            oc_s = str(ocust).strip()
            op_s = str(ocode).strip()
            pack_map[op_s] = int(ws_price.cell(r, 14).value or 1)
            other_price_map[(oc_s, op_s)] = int(ws_price.cell(r, 15).value or 0)

    ws_prod = wb_val["상품"]
    prod_name_map: Dict[str, str] = {}
    for r in range(2, ws_prod.max_row + 1):
        pc = ws_prod.cell(r, 2).value
        pn = ws_prod.cell(r, 3).value
        if pc and pn:
            prod_name_map[str(pc).strip()] = str(pn).strip()

    order_sheets = [
        ("백제약품_주문입력", 86030001),
        ("지오영_주문입력", 86049000),
        ("타거래처_주문입력", None),
    ]

    records: List[dict] = []
    today_int = int(datetime.datetime.today().strftime("%Y%m%d"))
    tomorrow_int = int((datetime.datetime.today() + datetime.timedelta(days=1)).strftime("%Y%m%d"))

    for sh_name, fixed_bill_to in order_sheets:
        ws_v = wb_val[sh_name]

        for r in range(2, ws_v.max_row + 1):
            if str(ws_v.cell(r, 1).value or "").strip() == "합계":
                continue

            pcode = ws_v.cell(r, 8).value
            if not pcode or str(pcode).startswith("="):
                continue
            pcode = str(pcode).strip()

            sales_qty = ws_v.cell(r, 11).value
            promo_qty = ws_v.cell(r, 12).value
            sales_qty = int(float(sales_qty)) if isinstance(sales_qty, (int, float)) else 0
            promo_qty = int(float(promo_qty)) if isinstance(promo_qty, (int, float)) else 0

            if sales_qty <= 0 and promo_qty <= 0:
                continue

            odt = ws_v.cell(r, 2).value
            ddt = ws_v.cell(r, 3).value
            if odt is None or isinstance(odt, str):
                odt = today_int
            if ddt is None or isinstance(ddt, str):
                ddt = tomorrow_int

            if fixed_bill_to is not None:
                bill_to = fixed_bill_to
                ship_to = int(float(ws_v.cell(r, 6).value))
                cust_name = bill_to_name_map.get(bill_to, "백제약품" if bill_to == 86030001 else "지오영")
                center_name = ship_to_center_map.get(ship_to, str(ws_v.cell(r, 5).value or ""))
            else:
                cust_name = str(ws_v.cell(r, 4).value or "").strip()
                bill_raw = ws_v.cell(r, 5).value
                ship_raw = ws_v.cell(r, 6).value
                bill_to = int(float(bill_raw))
                ship_to = int(float(ship_raw)) if ship_raw is not None else bill_to
                center_name = ""

            raw_price = ws_v.cell(r, 13).value
            if isinstance(raw_price, (int, float)) and raw_price > 0:
                unit_price = int(round(float(raw_price)))
            else:
                if bill_to == 86030001:
                    unit_price = baekje_price_map.get(pcode, 0)
                elif bill_to == 86049000:
                    unit_price = geoyoung_price_map.get(pcode, 0)
                else:
                    unit_price = other_price_map.get((cust_name, pcode), 0)

            pack_size = ws_v.cell(r, 10).value
            pack_size = int(float(pack_size)) if isinstance(pack_size, (int, float)) and pack_size > 0 else pack_map.get(pcode, 1)
            pname = prod_name_map.get(pcode, str(ws_v.cell(r, 9).value or pcode).strip())

            is_zpk = (bill_to == 86020000) or (cust_name.upper() == "ZPK")

            if bill_to in (86030001, 86049000):
                group_key = f"{bill_to}_{ship_to}"
                display_customer = f"{cust_name} ({center_name})" if center_name else cust_name
                file_label = f"{cust_name}_{center_name}" if center_name else cust_name
            else:
                group_key = f"{bill_to}"
                display_customer = cust_name
                file_label = cust_name

            if sales_qty > 0:
                records.append({
                    "group_key": group_key,
                    "is_zpk": is_zpk,
                    "bill_to": bill_to,
                    "ship_to": ship_to,
                    "cust_name": cust_name,
                    "center_name": center_name,
                    "display_customer": display_customer,
                    "file_label": file_label,
                    "order_date": odt,
                    "delivery_date": ddt,
                    "pcode": pcode,
                    "pname": pname,
                    "pack_size": pack_size,
                    "qty": sales_qty,
                    "unit_price": unit_price,
                    "amount": sales_qty * unit_price,
                    "is_promo": False,
                })

            if promo_qty > 0:
                records.append({
                    "group_key": group_key,
                    "is_zpk": is_zpk,
                    "bill_to": bill_to,
                    "ship_to": ship_to,
                    "cust_name": cust_name,
                    "center_name": center_name,
                    "display_customer": display_customer,
                    "file_label": file_label,
                    "order_date": odt,
                    "delivery_date": ddt,
                    "pcode": pcode,
                    "pname": f"{pname} (할증분)",
                    "pack_size": pack_size,
                    "qty": promo_qty,
                    "unit_price": 0,
                    "amount": 0,
                    "is_promo": True,
                })

    df_all = pd.DataFrame(records)

    ws_up = wb_val["수주업로드_서식"]
    up_rows = []
    for r in range(2, ws_up.max_row + 1):
        b_val = ws_up.cell(r, 4).value
        q_val = ws_up.cell(r, 10).value
        a_val = ws_up.cell(r, 12).value
        if b_val and isinstance(q_val, (int, float)) and q_val > 0:
            up_rows.append({
                "bill_to": int(float(b_val)),
                "ship_to": int(float(ws_up.cell(r, 6).value or b_val)),
                "qty": int(float(q_val)),
                "amount": int(round(float(a_val or 0))),
            })
    df_upload_cache = pd.DataFrame(up_rows)
    return df_all, df_upload_cache


def fill_statement_workbook(template_path: str, group_df: pd.DataFrame) -> Tuple[bytes, dict]:
    \"\"\"
    단일 거래처(또는 단일 센터)의 주문 데이터만 격리하여 거래명세서 엑셀을 채우고
    수량/공급가액 총계 더블체크 결과를 반환합니다.
    (출력 PDF에 제품코드는 기입하지 않음)
    \"\"\"
    wb = openpyxl.load_workbook(template_path)
    base_ws = wb.worksheets[0]

    first_row = group_df.iloc[0]
    order_dt = parse_yyyymmdd(first_row["order_date"])
    deliv_dt = parse_yyyymmdd(first_row["delivery_date"])
    bill_to = int(first_row["bill_to"])
    display_customer = str(first_row["display_customer"])
    mmdd = deliv_dt.strftime("%m%d")

    items = group_df.to_dict("records")
    max_slots = 22
    num_pages = max(1, math.ceil(len(items) / max_slots))

    base_ws["K5"].value = None
    for r in range(15, 37):
        base_ws.cell(r, 2).value = None
        base_ws.cell(r, 3).value = None
        base_ws.cell(r, 9).value = None
        base_ws.cell(r, 10).value = None
        base_ws.cell(r, 11).value = None
        base_ws.cell(r, 12).value = None
        base_ws.cell(r, 13).value = f"=K{r}*L{r}"

    sheets_to_fill = [base_ws]
    base_ws.title = f"{display_customer[:20]}({mmdd})" if num_pages == 1 else f"{display_customer[:17]}_1({mmdd})"
    for p in range(1, num_pages):
        copied = wb.copy_worksheet(base_ws)
        copied.title = f"{display_customer[:17]}_{p+1}({mmdd})"
        sheets_to_fill.append(copied)

    written_total_qty = 0
    written_total_amount = 0

    for p_idx, ws in enumerate(sheets_to_fill):
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

        ws["C8"].value = order_dt
        ws["C8"].number_format = "yyyy-mm-dd"
        ws["G8"].value = deliv_dt
        ws["G8"].number_format = "yyyy-mm-dd"
        ws["M8"].value = bill_to

        ws["I10"].value = display_customer
        ws["A38"].value = f"*배송처: {display_customer}"

        chunk = items[p_idx * max_slots : (p_idx + 1) * max_slots]
        for idx, item in enumerate(chunk):
            r = 15 + idx
            qty = int(item["qty"])
            price = int(item["unit_price"])
            amt = qty * price

            ws.cell(r, 1).value = p_idx * max_slots + idx + 1
            ws.cell(r, 2).value = None
            ws.cell(r, 3).value = item["pname"]
            ws.cell(r, 9).value = int(item["pack_size"])
            ws.cell(r, 11).value = qty
            ws.cell(r, 12).value = price
            ws.cell(r, 13).value = f"=K{r}*L{r}"

            written_total_qty += qty
            written_total_amount += amt

        ws["L37"].value = "=SUM(M15:M36)"
        ws["L38"].value = "=ROUND(L37*0.1,0)"
        ws["L39"].value = "=L37+L38"

    expected_qty = int(group_df["qty"].sum())
    expected_amount = int(group_df["amount"].sum())
    is_matched = (expected_qty == written_total_qty) and (expected_amount == written_total_amount)

    check_info = {
        "display_customer": display_customer,
        "bill_to": bill_to,
        "item_count": len(items),
        "expected_qty": expected_qty,
        "written_qty": written_total_qty,
        "expected_amount": expected_amount,
        "written_amount": written_total_amount,
        "vat": round(written_total_amount * 0.1),
        "gross_total": written_total_amount + round(written_total_amount * 0.1),
        "matched": is_matched,
    }

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue(), check_info


def convert_xlsx_to_pdf(xlsx_bytes: bytes, base_filename: str) -> bytes:
    \"\"\"LibreOffice Headless를 호출하여 스타일·이미지가 보존된 PDF 바이트 생성\"\"\"
    with tempfile.TemporaryDirectory() as tmpdir:
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", base_filename)
        xlsx_path = os.path.join(tmpdir, f"{safe_name}.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_bytes)

        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                xlsx_path,
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )

        pdf_path = os.path.join(tmpdir, f"{safe_name}.pdf")
        with open(pdf_path, "rb") as f:
            return f.read()


def generate_all_statements_zip(
    order_excel_bytes: bytes, template_path: str
) -> Tuple[bytes, pd.DataFrame, List[dict]]:
    \"\"\"
    전체 거래처·센터별 거래명세서 PDF를 일괄 생성하고 ZIP 아카이브와 더블체크 리포트를 반환
    \"\"\"
    df_all, df_upload_cache = load_orders_from_v7(order_excel_bytes)
    if df_all.empty:
        return b"", pd.DataFrame(), []

    df_target = df_all[~df_all["is_zpk"]].copy()
    if df_target.empty:
        return b"", pd.DataFrame(), []

    zip_buf = io.BytesIO()
    summary_rows: List[dict] = []
    individual_pdfs: List[dict] = []

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for group_key, grp in df_target.groupby("group_key", sort=False):
            isolated_grp = grp.copy().reset_index(drop=True)
            xlsx_bytes, check_info = fill_statement_workbook(template_path, isolated_grp)

            if not df_upload_cache.empty:
                b_to = int(isolated_grp.iloc[0]["bill_to"])
                if b_to in (86030001, 86049000):
                    s_to = int(isolated_grp.iloc[0]["ship_to"])
                    up_sub = df_upload_cache[
                        (df_upload_cache["bill_to"] == b_to) & (df_upload_cache["ship_to"] == s_to)
                    ]
                else:
                    up_sub = df_upload_cache[df_upload_cache["bill_to"] == b_to]

                if not up_sub.empty:
                    up_qty = int(up_sub["qty"].sum())
                    up_amt = int(up_sub["amount"].sum())
                    check_info["upload_sheet_qty"] = up_qty
                    check_info["upload_sheet_amount"] = up_amt
                    if up_qty != check_info["written_qty"] or up_amt != check_info["written_amount"]:
                        check_info["matched"] = False
                else:
                    check_info["upload_sheet_qty"] = check_info["expected_qty"]
                    check_info["upload_sheet_amount"] = check_info["expected_amount"]
            else:
                check_info["upload_sheet_qty"] = check_info["expected_qty"]
                check_info["upload_sheet_amount"] = check_info["expected_amount"]

            deliv_dt = parse_yyyymmdd(isolated_grp.iloc[0]["delivery_date"])
            mmdd = deliv_dt.strftime("%m%d")
            safe_label = re.sub(r'[\\/*?:"<>|]', "_", str(isolated_grp.iloc[0]["file_label"]))
            pdf_filename = f"거래명세서_{safe_label}({mmdd}).pdf"

            pdf_bytes = convert_xlsx_to_pdf(xlsx_bytes, f"거래명세서_{safe_label}({mmdd})")
            zf.writestr(pdf_filename, pdf_bytes)

            summary_rows.append(check_info)
            individual_pdfs.append({
                "filename": pdf_filename,
                "customer": check_info["display_customer"],
                "pdf_bytes": pdf_bytes,
            })

    return zip_buf.getvalue(), pd.DataFrame(summary_rows), individual_pdfs
"""
    try:
        with open(fallback_path, "w", encoding="utf-8") as f:
            f.write(embedded_source)
    except OSError:
        pass

    module = type(sys)("pdf_generator")
    exec(compile(embedded_source, "pdf변환_엔진.py", "exec"), module.__dict__)
    sys.modules["pdf_generator"] = module
    return module


try:
    _engine = _load_pdf_engine()
    generate_all_statements_zip = _engine.generate_all_statements_zip
    load_orders_from_v7 = _engine.load_orders_from_v7
except Exception as e:
    st.error(f"PDF 변환 엔진 초기화 오류: {e}")
    st.stop()


def _find_template_path() -> str:
    """양식 엑셀 파일 경로 탐색 (루트, 양식폴더, templates 폴더 모두 지원)"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "거래명세서_양식_동원약품(1123).xlsx"),
        os.path.join(base_dir, "github_upload_ready_거래명세서_양식_동원약품(1123).xlsx"),
        os.path.join(base_dir, "거래명세서_동원약품(1123).xlsx"),
        os.path.join(base_dir, "양식", "거래명세서_양식_동원약품(1123).xlsx"),
        os.path.join(base_dir, "templates", "거래명세서_동원약품(1123).xlsx"),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


TEMPLATE_PATH = _find_template_path()

st.title("📄 거래명세표 PDF 자동 생성 및 검증 시스템")
st.caption(
    "`거래처별_주문입력_및_수주업로드_v7.xlsx` 파일을 업로드하면 거래처별·센터별 거래명세표 PDF를 일괄 생성합니다."
)

with st.expander("📌 적용된 생성 규칙 및 보안 검증 기준 확인", expanded=False):
    st.markdown(
        """
        - **양식 변환**: 거래명세서 엑셀 양식 서식·로고·직인 그대로 PDF 변환 (제품코드 미기입)
        - **출력 단위**: 일반 거래처는 **거래처별 1부**, **백제약품·지오영은 배송센터별 1부** 분리 생성
        - **ZPK 제외**: `ZPK` 주문 건은 거래명세표 출력 대상에서 자동 제외
        - **일자 및 상호**: `주문일자 = 수주일자`, `배송일자 = 납품일자`, `상호 = 거래처명(센터명)` 자동 기입
        - **총계 더블체크**: 주문파일의 거래처별 수량·공급가액 총계와 명세서 기입 총계가 100% 일치할 때만 다운로드 허용
        - **단가 유출 방지**: 거래처/센터별 데이터를 메모리상에서 완전 격리하여 타 거래처 단가 혼입 원천 차단
        """
    )

uploaded_order = st.file_uploader(
    "1️⃣ 수주 파일 업로드 (`거래처별_주문입력_및_수주업로드_v7.xlsx`)",
    type=["xlsx"],
)

if not os.path.exists(TEMPLATE_PATH):
    st.error(
        "거래명세서 양식 엑셀 파일(`거래명세서_양식_동원약품(1123).xlsx` 또는 `거래명세서_동원약품(1123).xlsx`)을 찾을 수 없습니다. "
        "GitHub 레포지토리에 양식 엑셀 파일을 함께 업로드해 주세요."
    )
    st.stop()

if uploaded_order is not None:
    order_bytes = uploaded_order.getvalue()

    with st.spinner("주문 데이터 분석, 단가 격리 검증 및 거래명세표 PDF 변환 중..."):
        df_all, _ = load_orders_from_v7(order_bytes)

        if df_all.empty:
            st.warning("업로드된 파일에 양수(`>0`)로 입력된 주문 내역이 없습니다.")
            st.stop()

        zpk_df = df_all[df_all["is_zpk"]]
        if not zpk_df.empty:
            st.info(
                f"ℹ️ **ZPK 제외 안내**: ZPK 주문 {len(zpk_df)}건(총 {int(zpk_df['qty'].sum()):,}개)은 규정에 따라 명세서 출력 대상에서 자동 제외되었습니다."
            )

        zip_bytes, df_summary, individual_pdfs = generate_all_statements_zip(
            order_bytes, TEMPLATE_PATH
        )

    if df_summary.empty:
        st.warning("ZPK 제외 후 발행할 거래명세표 대상 주문이 없습니다.")
        st.stop()

    st.subheader("2️⃣ 거래처·센터별 수량 및 공급가액 총계 더블체크")

    all_matched = bool(df_summary["matched"].all())
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("발행 명세서 수", f"{len(df_summary)}건")
    col2.metric("총 출고 수량", f"{int(df_summary['written_qty'].sum()):,}개")
    col3.metric("총 공급가액(VAT별도)", f"{int(df_summary['written_amount'].sum()):,}원")
    col4.metric("검증 상태", "✅ 전체 일치" if all_matched else "❌ 불일치 발생")

    display_df = df_summary.rename(
        columns={
            "display_customer": "상호(거래처/센터)",
            "bill_to": "발주처코드",
            "item_count": "품목행수",
            "expected_qty": "주문파일 수량계",
            "written_qty": "명세서 수량계",
            "expected_amount": "주문파일 공급가액",
            "written_amount": "명세서 공급가액",
            "vat": "부가세(10%)",
            "gross_total": "공급가총액",
            "matched": "더블체크 일치",
        }
    )[
        [
            "상호(거래처/센터)",
            "발주처코드",
            "품목행수",
            "주문파일 수량계",
            "명세서 수량계",
            "주문파일 공급가액",
            "명세서 공급가액",
            "부가세(10%)",
            "공급가총액",
            "더블체크 일치",
        ]
    ]

    st.dataframe(
        display_df.style.format(
            {
                "주문파일 수량계": "{:,}",
                "명세서 수량계": "{:,}",
                "주문파일 공급가액": "{:,}",
                "명세서 공급가액": "{:,}",
                "부가세(10%)": "{:,}",
                "공급가총액": "{:,}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("3️⃣ 거래명세표 PDF 다운로드")
    if not all_matched:
        st.error(
            "수주파일 총계와 거래명세서 총계가 일치하지 않는 항목이 있어 다운로드가 차단되었습니다. 수식 및 단가를 확인해주세요."
        )
    else:
        st.success("모든 거래처·센터의 수량 및 공급가액 총계가 100% 일치합니다.")
        st.download_button(
            label=f"📦 거래명세표 PDF 일괄 다운로드 ({len(individual_pdfs)}건 ZIP)",
            data=zip_bytes,
            file_name="거래명세표_일괄다운로드.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

        with st.expander("개별 거래처·센터 PDF 따로 받기", expanded=False):
            for idx, item in enumerate(individual_pdfs):
                c_left, c_right = st.columns([4, 1])
                c_left.write(f"• **{item['customer']}** (`{item['filename']}`)")
                c_right.download_button(
                    label="PDF 받기",
                    data=item["pdf_bytes"],
                    file_name=item["filename"],
                    mime="application/pdf",
                    key=f"dl_{idx}",
                )
