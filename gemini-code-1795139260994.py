import importlib.util
import os
import sys
import streamlit as st

st.set_page_config(
    page_title="거래명세표 PDF 일괄 생성 시스템",
    page_icon="📄",
    layout="wide",
)


def _load_pdf_generator():
    """pdf_generator.py 또는 gemini-code-pdf_generator.py를 자동으로 찾아 로드"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
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

    raise FileNotFoundError(
        "pdf_generator.py 또는 gemini-code-pdf_generator.py 파일을 찾을 수 없습니다."
    )


try:
    _pdf_gen = _load_pdf_generator()
    generate_all_statements_zip = _pdf_gen.generate_all_statements_zip
    load_orders_from_v7 = _pdf_gen.load_orders_from_v7
except Exception as e:
    st.error(f"모듈 로드 오류: {e}")
    st.stop()


def _find_template_path() -> str:
    """양식 엑셀 파일 경로 탐색 (templates 폴더 또는 루트 폴더 지원)"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "templates", "거래명세서_동원약품(1123).xlsx"),
        os.path.join(base_dir, "거래명세서_동원약품(1123).xlsx"),
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
        - **양식 변환**: `거래명세서_동원약품(1123).xlsx` 원본 서식 그대로 PDF 변환
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
        "거래명세서 양식 파일(`거래명세서_동원약품(1123).xlsx`)을 찾을 수 없습니다. "
        "GitHub 레포지토리의 `templates/거래명세서_동원약품(1123).xlsx` 또는 최상단 경로에 업로드해 주세요."
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

    # 2️⃣ 더블체크 검증 결과 표시
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

    # 3️⃣ 일괄 ZIP 다운로드 및 개별 PDF 다운로드
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