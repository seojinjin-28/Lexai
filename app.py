
import streamlit as st
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from io import BytesIO

st.set_page_config(
    page_title="Lexai",
    page_icon="🏗️",
    layout="wide"
)

# -------------------------
# Lexai 기본 화면
# -------------------------

st.title("🏗️ Lexai")
st.caption("Construction Contract Risk Analyzer")

menu = st.sidebar.radio(
    "분석 메뉴",
    [
        "건설 하도급 계약서 분석",
        "외국인 근로자 노무계약 미스매치 실드"
    ]
)

# -------------------------
# 문서 텍스트 추출 기능
# -------------------------

def extract_text(file):
    text = ""
    name = file.name.lower()

    if name.endswith(".pdf"):
        reader = PdfReader(file)
        for page in reader.pages:
            text += page.extract_text() or ""

    elif name.endswith(".docx"):
        document = Document(file)
        for paragraph in document.paragraphs:
            text += paragraph.text + "\n"

    elif name.endswith(".pptx"):
        presentation = Presentation(file)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text += shape.text + "\n"

    return text

# ==================================================
# Lexai Word 보고서 공통 디자인
# ==================================================

def style_report(report):

    # 기본 글꼴 설정
    styles = report.styles

    styles["Normal"].font.name = "맑은 고딕"
    styles["Normal"].font.size = Pt(10.5)
    styles["Normal"].font.color.rgb = RGBColor(0, 0, 0)

    # 제목 스타일
    styles["Title"].font.name = "맑은 고딕"
    styles["Title"].font.size = Pt(22)
    styles["Title"].font.bold = True
    styles["Title"].font.color.rgb = RGBColor(0, 0, 0)

    # 큰 소제목
    styles["Heading 1"].font.name = "맑은 고딕"
    styles["Heading 1"].font.size = Pt(16)
    styles["Heading 1"].font.bold = True
    styles["Heading 1"].font.color.rgb = RGBColor(0, 0, 0)

    # 작은 소제목
    styles["Heading 2"].font.name = "맑은 고딕"
    styles["Heading 2"].font.size = Pt(12)
    styles["Heading 2"].font.bold = True
    styles["Heading 2"].font.color.rgb = RGBColor(0, 0, 0)

    # 문서 전체 문단 및 글자 설정
    for paragraph in report.paragraphs:

        paragraph.paragraph_format.space_after = Pt(8)
        paragraph.paragraph_format.line_spacing = 1.25

        for run in paragraph.runs:
            run.font.name = "맑은 고딕"
            run.font.color.rgb = RGBColor(0, 0, 0)

    # 첫 번째 문단을 메인 제목으로 가운데 정렬
    if report.paragraphs:

        title = report.paragraphs[0]

        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title.paragraph_format.space_before = Pt(20)
        title.paragraph_format.space_after = Pt(18)

        for run in title.runs:
            run.font.name = "맑은 고딕"
            run.font.size = Pt(22)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)

    # 페이지 여백 설정
    for section in report.sections:
        section.top_margin = Pt(60)
        section.bottom_margin = Pt(60)
        section.left_margin = Pt(65)
        section.right_margin = Pt(65)

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def set_korean_font(report):

    for paragraph in report.paragraphs:
        for run in paragraph.runs:

            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.rFonts

            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                r_pr.insert(0, r_fonts)

            r_fonts.set(qn("w:eastAsia"), "맑은 고딕")

# -------------------------
# 기존 하도급 계약서 분석 규칙
# -------------------------

risk_rules = {
    "민원 처리 비용 전가": {
        "keywords": [
            "모든 민원",
            "전액 수급사업자가 부담",
            "귀책사유와 관계없이"
        ],
        "level": "높음",
        "reason": "민원 발생 원인과 책임을 구분하지 않고 모든 비용을 수급사업자에게 부담시키는 내용입니다.",
        "suggestion": "민원 발생 원인과 각 당사자의 귀책사유를 확인한 후 책임이 있는 당사자가 합리적인 범위에서 비용을 부담하도록 수정합니다."
    },
    "자재 가격 변동에 따른 대금 조정 제한": {
        "keywords": [
            "어떠한 경우에도 조정하지 않는다",
            "대금 조정을 요청할 수 없다"
        ],
        "level": "높음",
        "reason": "자재 가격 등 계약금액에 영향을 주는 변동이 발생해도 대금 조정을 전면 제한하는 내용입니다.",
        "suggestion": "관련 법령과 계약 조건에 따른 대금 조정 요건 및 절차를 명시합니다."
    },
    "과도한 지체상금": {
        "keywords": [
            "계약금액의 3%",
            "지체일수 1일마다"
        ],
        "level": "높음",
        "reason": "높은 지체상금을 부과하고 책임 범위를 구분하지 않는 내용입니다.",
        "suggestion": "합리적인 지체상금률과 상한을 정하고 지연 원인 및 귀책사유를 구분하도록 수정합니다."
    },
    "추가 공사 비용 전가": {
        "keywords": [
            "별도의 비용을 청구할 수 없다",
            "추가 작업"
        ],
        "level": "높음",
        "reason": "추가 작업이 발생하더라도 비용을 청구하지 못하도록 제한하는 내용입니다.",
        "suggestion": "추가 작업의 범위와 비용을 사전에 협의하고 서면으로 합의하도록 수정합니다."
    },
    "일방적인 손해배상 책임": {
        "keywords": [
            "모든 손해를 배상한다",
            "어떠한 경우에도 손해배상 책임을 부담하지 않는다"
        ],
        "level": "높음",
        "reason": "당사자별 책임 범위를 구분하지 않고 손해배상 책임을 일방에게 부담시키는 내용입니다.",
        "suggestion": "각 당사자의 귀책사유와 실제 손해 범위에 따라 책임을 부담하도록 수정합니다."
    }
}


# 위험 항목별 관련 법률 조항 데이터
LAW_REFERENCES = {
    "민원 처리 비용 전가": [
        {
            "law": "하도급거래 공정화에 관한 법률",
            "article": "제3조의4 제2항 제2호",
            "reason": "원사업자의 의무사항이나 책임을 수급사업자에게 전가하는 부당한 특약인지 검토"
        }
    ],
    "자재 가격 상승에 따른 대금 조정 배제": [
        {
            "law": "하도급거래 공정화에 관한 법률",
            "article": "제16조의2",
            "reason": "공급원가 변동에 따른 하도급대금 조정 관련 규정"
        }
    ],
    "과도한 지체상금": [
        {
            "law": "민법",
            "article": "제398조 제2항",
            "reason": "손해배상액의 예정이 부당하게 과다한 경우 감액 가능성 검토"
        },
        {
            "law": "건설산업기본법",
            "article": "제22조",
            "reason": "건설공사 도급계약의 원칙 관련 규정"
        }
    ],
    "추가 공사비 전가": [
        {
            "law": "하도급거래 공정화에 관한 법률",
            "article": "제3조의4",
            "reason": "부당한 특약에 해당하는지 검토"
        }
    ],
    "일방적인 책임 전가": [
        {
            "law": "하도급거래 공정화에 관한 법률",
            "article": "제3조의4",
            "reason": "수급사업자에게 부당하게 불리한 특약인지 검토"
        },
        {
            "law": "건설산업기본법",
            "article": "제22조",
            "reason": "건설공사 도급계약의 공정한 체결 및 이행 원칙 검토"
        }
    ]
}

# 법률 조항별 원문 데이터
LAW_TEXTS = {
    "하도급거래 공정화에 관한 법률|제3조의4 제2항 제2호": {
        "text": (
            "원사업자가 부담하여야 할 민원처리, 산업재해 등과 "
            "관련된 비용을 수급사업자에게 부담시키는 약정"
        ),
        "date": "2026년 10월 2일 시행 기준"
    },

    "하도급거래 공정화에 관한 법률|제3조의4": {
        "text": (
            "① 원사업자는 수급사업자의 이익을 부당하게 침해하거나 "
            "제한하는 계약조건(이하 '부당한 특약'이라 한다)을 설정하여서는 아니 된다.\n\n"
            "② 다음 각 호의 어느 하나에 해당하는 약정은 부당한 특약으로 본다.\n"
            "1. 서면에 기재되지 아니한 사항을 요구함에 따라 발생된 비용을 "
            "수급사업자에게 부담시키는 약정\n"
            "2. 원사업자가 부담하여야 할 민원처리, 산업재해 등과 관련된 "
            "비용을 수급사업자에게 부담시키는 약정\n"
            "3. 입찰내역에 없는 사항을 요구함에 따라 발생된 비용을 "
            "수급사업자에게 부담시키는 약정",
        ),
        "date": "2026년 10월 2일 시행 기준"
    },

    "하도급거래 공정화에 관한 법률|제16조의2": {
        "text": (
            "① 수급사업자는 제조등의 위탁을 받은 후 다음 각 호의 "
            "어느 하나에 해당하여 하도급대금의 조정이 불가피한 경우에는 "
            "원사업자에게 하도급대금의 조정을 신청할 수 있다.\n\n"
            "1. 목적물등의 공급원가가 변동되는 경우\n"
            "2. 수급사업자의 책임으로 돌릴 수 없는 사유로 목적물등의 "
            "납품등 시기가 지연되어 관리비 등 공급원가 외의 비용이 "
            "변동되는 경우"
        ),
        "date": "2026년 10월 2일 시행 기준"
    },

    
    "건설산업기본법|제22조": {
        "text": (
            "① 건설공사에 관한 도급계약(하도급계약을 포함한다)의 "
            "당사자는 대등한 입장에서 합의에 따라 공정하게 계약을 체결하고 "
            "신의를 지켜 성실하게 계약을 이행하여야 한다.\n\n"

            "⑤ 건설공사 도급계약의 내용이 당사자 일방에게 현저하게 "
            "불공정한 경우로서 다음 각 호의 어느 하나에 해당하는 경우에는 "
            "그 부분에 한정하여 무효로 한다.\n\n"

            "1. 계약체결 이후 설계변경, 경제상황의 변동에 따라 발생하는 "
            "계약금액의 변경을 상당한 이유 없이 인정하지 아니하거나 "
            "그 부담을 상대방에게 떠넘기는 경우\n\n"

            "2. 계약체결 이후 공사내용의 변경에 따른 계약기간의 변경을 "
            "상당한 이유 없이 인정하지 아니하거나 그 부담을 상대방에게 "
            "떠넘기는 경우\n\n"

            "3. 도급계약의 형태, 건설공사의 내용 등 관련된 모든 사정에 "
            "비추어 계약체결 당시 예상하기 어려운 내용에 대하여 "
            "상대방에게 책임을 떠넘기는 경우\n\n"

            "4. 계약내용에 대하여 구체적인 정함이 없거나 당사자 간 이견이 "
            "있을 경우 계약내용을 일방의 의사에 따라 정함으로써 "
            "상대방의 정당한 이익을 침해한 경우\n\n"

            "5. 계약불이행에 따른 당사자의 손해배상책임을 과도하게 "
            "경감하거나 가중하여 정함으로써 상대방의 정당한 이익을 "
            "침해한 경우\n\n"

            "6. 민법 등 관계 법령에서 인정하고 있는 상대방의 권리를 "
            "상당한 이유 없이 배제하거나 제한하는 경우"
        ),
        "date": "2025년 11월 27일 시행 기준"
    },


    "민법|제398조 제2항": {
        "text": (
            "손해배상의 예정액이 부당히 과다한 경우에는 "
            "법원은 적당히 감액할 수 있다."
        ),
        "date": "2026년 3월 17일 시행 기준"
    }
}

def analyze_contract(text):
    results = []

    for risk_name, rule in risk_rules.items():
        matched = [
            keyword for keyword in rule["keywords"]
            if keyword in text
        ]

        if matched:
            results.append({
                "name": risk_name,
                "keywords": matched,
                "level": rule["level"],
                "reason": rule["reason"],
                "suggestion": rule["suggestion"],
                "laws": LAW_REFERENCES.get(risk_name, [])
            })

    return results


# ==================================================
# 메뉴 1. 건설 하도급 계약서 분석
# ==================================================

if menu == "건설 하도급 계약서 분석":

    st.header("📄 건설 하도급 계약서 분석")
    st.write("건설공사 계약서의 불공정 특약 및 위험 가능성을 검토합니다.")

    uploaded_file = st.file_uploader(
        "분석할 계약서를 업로드하세요.",
        type=["pdf", "docx", "pptx"],
        key="subcontract_file"
    )

    if uploaded_file is not None:

        text = extract_text(uploaded_file)

        if text.strip():

            st.success("계약서 업로드 및 텍스트 추출 완료!")

            if st.button("🔍 계약서 위험 조항 분석하기"):

                results = analyze_contract(text)

                st.subheader("📊 분석 결과")

                if results:

                    st.warning(
                        f"위험 검토 대상 조항: {len(results)}개"
                    )

                    for i, result in enumerate(results, 1):

                        with st.container(border=True):

                            st.markdown(
                                f"### {i}. {result['name']}"
                            )

                            st.error(
                                f"위험도: {result['level']}"
                            )

                            st.write("**탐지된 표현**")
                            for keyword in result["keywords"]:
                                st.write(f"- {keyword}")

                            st.write("**위험 사유**")
                            st.write(result["reason"])

                            st.write("**개선 제안 문구**")
                            st.info(result["suggestion"])


                            # 관련 법률 조항 표시
                            st.write("**관련 법률 조항**")

                            if result.get("laws"):
                                
                                for law in result["laws"]:
                                    st.write(f"법률명: {law['law']}")
                                    st.write(f"관련 조항: {law['article']}")
                                    st.write(f"검토 이유: {law['reason']}")

                                    # 법률 원문 데이터 연결
                                    law_key = f"{law['law']}|{law['article']}"
                                    law_data = LAW_TEXTS.get(law_key)

                                    if law_data:
                                        with st.expander("법률 조항 내용 확인하기"):
                                            st.caption(f"기준: {law_data['date']}")
                                            st.write(law_data["text"])
                                            st.caption("※ 법률 내용 중 관련 부분을 발췌한 참고 자료입니다. 공식 현행 법령 원문과 대조가 필요합니다.")
                                    else:
                                        st.warning("등록된 법률 원문이 없습니다.")

                            else:
                                st.write("연결된 법률 조항이 없습니다.")


                    # Word 보고서 생성
                    report = Document()

                    report.add_heading(
                        "Lexai 건설공사 계약서 분석 보고서",
                        0
                    )

                    report.add_paragraph(
                        f"분석 대상: {uploaded_file.name}"
                    )

                    report.add_paragraph(
                        f"위험 검토 대상: {len(results)}개"
                    )

                    for i, result in enumerate(results, 1):

                        report.add_heading(
                            f"{i}. {result['name']}",
                            level=1
                        )

                        report.add_paragraph(
                            f"위험도: {result['level']}"
                        )

                        report.add_paragraph("탐지된 표현")
                        for keyword in result["keywords"]:
                            report.add_paragraph(
                                keyword,
                                style="List Bullet"
                            )

                        report.add_paragraph("위험 사유")
                        report.add_paragraph(result["reason"])

                        report.add_paragraph("개선 제안")
                        report.add_paragraph(result["suggestion"])

                        # 관련 법률 조항 및 원문
                        report.add_heading("관련 법률 조항", level=2)

                        if result.get("laws"):
                            for law in result["laws"]:
                                report.add_paragraph(f"법률명: {law['law']}")
                                report.add_paragraph(f"관련 조항: {law['article']}")
                                report.add_paragraph(f"검토 이유: {law['reason']}")

                                # 법률 원문 검색
                                law_key = f"{law['law']}|{law['article']}"
                                law_data = LAW_TEXTS.get(law_key)

                                if law_data:
                                    report.add_paragraph("법률 원문")
                                    report.add_paragraph(law_data["text"])
                                    report.add_paragraph(
                                        f"법령 기준일: {law_data['date']}"
                                    )
                                else:
                                    report.add_paragraph(
                                        "등록된 법률 원문이 없습니다."
                                    )

                        else:
                            report.add_paragraph("연결된 법률 조항이 없습니다.")

                        report.add_paragraph(
                            "※ 법률 원문은 참고용 발췌 자료이며, "
                            "실제 법률 위반 여부는 구체적인 계약 내용과 "
                            "현행 법령을 추가로 검토해야 합니다."
                        )

                    report.add_paragraph(
                        "본 결과는 기초 키워드 분석이며 "
                        "법률상 위법 여부를 확정하지 않습니다. "
                        "실제 계약 적용 전 전문가 검토가 필요합니다."
                    )

                    # 보고서 디자인 적용
                    style_report(report)
                    set_korean_font(report)

                    # Word 파일 저장
                    buffer = BytesIO()
                    report.save(buffer)
                    buffer.seek(0)

                    st.download_button(
                        "📥 Word 분석 보고서 다운로드",
                        data=buffer.getvalue(),
                        file_name="Lexai_계약서_분석보고서.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                else:
                    st.success(
                        "등록된 위험 표현이 발견되지 않았습니다."
                    )

        else:
            st.error("문서에서 텍스트를 추출하지 못했습니다.")


# ==================================================
# 메뉴 2. 외국인 근로자 노무계약 미스매치 실드
# ==================================================


elif menu == "외국인 근로자 노무계약 미스매치 실드":

    st.header("🌏 외국인 근로자 노무계약 미스매치 실드")

    st.write(
        "한국어 근로계약서와 외국어 계약서의 내용 차이 및 "
        "안전보건 서약서의 위험 가능성을 검토합니다."
    )

    st.divider()

    st.subheader("1. 한국어 표준근로계약서")

    korean_file = st.file_uploader(
        "한국어 계약서를 업로드하세요.",
        type=["pdf", "docx", "pptx"],
        key="korean_file"
    )

    st.subheader("2. 외국어 근로계약서")

    foreign_file = st.file_uploader(
        "외국어 계약서를 업로드하세요.",
        type=["pdf", "docx", "pptx"],
        key="foreign_file"
    )

    st.subheader("3. 안전보건 서약서")

    safety_file = st.file_uploader(
        "안전보건 서약서를 업로드하세요.",
        type=["pdf", "docx", "pptx"],
        key="safety_file"
    )

    st.divider()

    if st.button("📑 문서 읽기 및 분석 준비"):

        if korean_file and foreign_file and safety_file:

            korean_text = extract_text(korean_file)
            foreign_text = extract_text(foreign_file)
            safety_text = extract_text(safety_file)

            st.session_state["korean_text"] = korean_text
            st.session_state["foreign_text"] = foreign_text
            st.session_state["safety_text"] = safety_text

            st.session_state["documents_loaded"] = True

        else:
            st.session_state["documents_loaded"] = False
            st.warning("세 문서를 모두 업로드해 주세요.")

    if st.session_state.get("documents_loaded", False):

        korean_text = st.session_state["korean_text"]
        foreign_text = st.session_state["foreign_text"]
        safety_text = st.session_state["safety_text"]

        st.divider()
        st.subheader("📊 문서 추출 결과")

        documents = [
            ("한국어 근로계약서", korean_text),
            ("외국어 근로계약서", foreign_text),
            ("안전보건 서약서", safety_text)
        ]

        for name, content in documents:

            with st.container(border=True):

                st.write(f"**{name}**")

                if content.strip():
                    st.success("텍스트 추출 성공")
                    st.write(f"추출된 글자 수: {len(content):,}자")
                else:
                    st.error(
                        "텍스트를 추출하지 못했습니다. "
                        "스캔된 문서인지 확인해 주세요."
                    )

        if (
            korean_text.strip()
            and foreign_text.strip()
            and safety_text.strip()
        ):
            st.success(
                "3개 문서의 내용 추출이 완료되었습니다. "
                "다음 단계의 비교 분석을 진행할 수 있습니다."
            )


        else:

            st.warning(
                "한국어 계약서, 외국어 계약서, 안전보건 서약서를 "
                "모두 업로드해 주세요."
            )

# ==================================================
# 한국어 / 외국어 근로계약서 비교 기능
# ==================================================

# 외국인 근로계약서 관련 법률 매칭 데이터
WORKER_LAW_REFERENCES = {
    "기본 월급": [
        {
            "law": "근로기준법",
            "article": "제17조",
            "reason": "근로계약 체결 시 임금의 명시 여부와 계약서 간 임금 조건 일치 여부를 확인해야 합니다."
        },
        {
            "law": "근로기준법",
            "article": "제43조",
            "reason": "임금 전액 지급 원칙에 따라 실제 지급 금액과 공제 내역을 확인해야 합니다."
        }
    ],

    "근무 종료 시간": [
        {
            "law": "근로기준법",
            "article": "제17조",
            "reason": "근로계약서에 명시된 소정근로시간이 서로 일치하는지 확인해야 합니다."
        },
        {
            "law": "근로기준법",
            "article": "제50조",
            "reason": "법정 근로시간을 초과하는 근무인지 확인해야 합니다."
        },
        {
            "law": "근로기준법",
            "article": "제53조",
            "reason": "연장근로가 법에서 정한 요건과 한도에 맞는지 확인해야 합니다."
        }
    ],

    "연장근로수당": [
        {
            "law": "근로기준법",
            "article": "제56조",
            "reason": "연장근로에 대한 가산임금 지급 조건이 적절하게 반영되어 있는지 확인해야 합니다."
        }
    ],

    "안전보호구 지급": [
        {
            "law": "산업안전보건기준에 관한 규칙",
            "article": "제32조",
            "reason": "해당 작업에 필요한 보호구를 사업주가 지급해야 하는 작업인지 확인해야 합니다."
        }
    ],

    "숙소 비용 공제": [
        {
            "law": "근로기준법",
            "article": "제43조",
            "reason": "임금에서 숙소 비용을 공제하는 경우 공제의 법적 근거와 동의 여부 등을 확인해야 합니다."
        },
        {
            "law": "외국인근로자의 고용 등에 관한 법률",
            "article": "제22조의2",
            "reason": "외국인 근로자에게 제공하는 기숙사의 기준과 관련 정보를 사전에 제공했는지 확인해야 합니다."
        }
    ]
}
# 외국인 근로계약서 관련 법률 원문
WORKER_LAW_TEXTS = {
    "근로기준법|제17조": {
        "date": "2026년 10월 2일 시행 기준",
        "text": (
            "사용자는 근로계약을 체결할 때 근로자에게 "
            "임금과 소정근로시간 등의 사항을 명시해야 합니다. "
            "임금의 구성항목, 계산방법, 지급방법 등은 "
            "서면으로 교부해야 합니다."
        )
    },

    "근로기준법|제43조": {
        "date": "2026년 10월 2일 시행 기준",
        "text": (
            "임금은 통화로 직접 근로자에게 전액 지급해야 합니다. "
            "다만 법령 또는 단체협약에 특별한 규정이 있는 경우 "
            "일부 공제가 가능합니다. "
            "임금은 매월 1회 이상 일정한 날짜에 지급해야 합니다."
        )
    },

    "근로기준법|제50조": {
        "date": "2026년 10월 8일 시행 기준",
        "text": (
            "1주 간의 근로시간은 휴게시간을 제외하고 "
            "40시간을 초과할 수 없습니다. "
            "1일의 근로시간은 휴게시간을 제외하고 "
            "8시간을 초과할 수 없습니다."
        )
    },

    "근로기준법|제56조": {
        "date": "2026년 10월 2일 시행 기준",
        "text": (
            "연장근로에 대해서는 통상임금의 50퍼센트 이상을 "
            "가산하여 지급해야 합니다. "
            "야간근로와 휴일근로에도 법정 가산 기준이 적용됩니다."
        )
    },

    "산업안전보건기준에 관한 규칙|제32조": {
        "date": "2026년 3월 2일 시행 기준",
        "text": (
            "사업주는 법에서 정한 위험 작업을 하는 근로자에게 "
            "작업조건에 맞는 보호구를 지급하고 착용하도록 해야 합니다."
        )
    },

    "외국인근로자의 고용 등에 관한 법률|제22조의2": {
        "date": "2025년 10월 1일 시행 기준",
        "text": (
            "사용자가 외국인근로자에게 기숙사를 제공하는 경우 "
            "기숙사 기준을 준수하고 건강과 안전을 지켜야 합니다. "
            "근로계약 체결 시 기숙사 구조, 설비, 위치, 환경, "
            "면적 등의 정보를 사전에 제공해야 합니다."
        )
    }
}
def compare_documents(korean_text, foreign_text):

    comparisons = [
        {
            "name": "기본 월급",
            "korean": ["2,800,000", "2800000"],
            "foreign": ["2.600.000", "2600000"],
            "reason": "한국어 계약서와 외국어 계약서의 기본급 금액이 다를 수 있습니다."
        },
        {
            "name": "근무 종료 시간",
            "korean": ["17:00"],
            "foreign": ["18:00"],
            "reason": "하루 근무 종료 시간이 서로 다릅니다."
        },
        {
            "name": "연장근로수당",
            "korean": ["별도로 산정", "별도로"],
            "foreign": ["bao gồm", "포함"],
            "reason": "연장근로수당이 별도 지급되는지 여부가 다를 수 있습니다."
        },
        {
            "name": "안전보호구 지급",
            "korean": ["사용자는", "지급한다"],
            "foreign": ["tự mua", "tự chịu chi phí"],
            "reason": "보호구를 회사가 지급하는지 근로자가 직접 구매하는지 확인해야 합니다."
        },
        {
            "name": "숙소 비용 공제",
            "korean": ["별도 서면 동의", "동의를 받아"],
            "foreign": ["200.000", "khấu trừ"],
            "reason": "숙소 비용 공제 방식과 동의 조건이 다를 수 있습니다."
        }
    ]

    results = []

    for item in comparisons:

        korean_found = any(
            word in korean_text
            for word in item["korean"]
        )

        foreign_found = any(
            word in foreign_text
            for word in item["foreign"]
        )

        
        if korean_found and foreign_found:
            result_item = item.copy()

            result_item["laws"] = WORKER_LAW_REFERENCES.get(
                item["name"], []
            )

            results.append(result_item)

    return results


# ==================================================
# 비교 결과 화면
# ==================================================

if menu == "외국인 근로자 노무계약 미스매치 실드":

    if st.button("🔍 한국어·외국어 계약서 비교하기"):

        if (
            st.session_state.get("documents_loaded", False)
            and st.session_state.get("korean_text", "").strip()
            and st.session_state.get("foreign_text", "").strip()
        ):

            comparison_results = compare_documents(
                st.session_state["korean_text"],
                st.session_state["foreign_text"]
            )

            st.session_state["comparison_results"] = comparison_results

        else:
            st.warning(
                "먼저 세 문서를 업로드하고 문서 읽기 및 분석 준비를 눌러주세요."
            )

    if "comparison_results" in st.session_state:

        results = st.session_state["comparison_results"]

        st.divider()
        st.subheader("📊 계약서 비교 결과")

        if results:

            st.warning(
                f"추가 확인이 필요한 항목: {len(results)}개"
            )

            for i, result in enumerate(results, 1):

                with st.container(border=True):

                    st.markdown(
                        f"### {i}. {result['name']}"
                    )

                    st.error("문서 간 차이 확인 필요")

                    st.write(result["reason"])
                    if result.get("laws"):
                        for law in result["laws"]:
                            st.write(f"법률명: {law['law']}")
                            st.write(f"관련 조항: {law['article']}")
                            st.write(f"검토 이유: {law['reason']}")

                            law_key = f"{law['law']}|{law['article']}"
                            law_data = globals().get("WORKER_LAW_TEXTS", {}).get(law_key)

                            if law_data:
                                with st.expander("법률 조항 내용 확인하기"):
                                    st.caption(f"기준: {law_data['date']}")
                                    st.write(law_data["text"])
                                    st.caption(
                                        "※ 법률 조항의 핵심을 요약한 참고 자료입니다. "
                                        "정확한 법적 판단은 현행 법령 원문을 확인해야 합니다."
                                    )
                            else:
                                st.warning("등록된 법률 원문이 없습니다.")

                    else:
                        st.caption("연결된 법률 조항이 없습니다.")

        else:

            st.success(
                "현재 등록된 비교 규칙에서 차이 의심 항목이 발견되지 않았습니다."
            )

        st.caption(
            "본 결과는 정해진 표현을 찾는 기초 비교 기능입니다. "
            "전체 문맥과 번역의 정확성을 판단하는 기능은 아닙니다."
        )

# ==================================================
# 안전보건 서약서 위험 조항 분석 기능
# ==================================================

# 안전보건 서약서 관련 법률 매칭 데이터
SAFETY_LAW_REFERENCES = {

    "사고 책임의 일괄 전가": [
        {
            "law": "산업안전보건법",
            "article": "제38조",
            "reason": "사업주의 안전조치 의무와 관련된 조항입니다. 사고 원인과 사업주의 안전조치 이행 여부를 함께 확인해야 합니다."
        },
        {
            "law": "산업안전보건법",
            "article": "제39조",
            "reason": "사업주의 보건조치 의무와 관련된 조항입니다."
        }
    ],

    "손해배상 비용의 전액 부담": [
        {
            "law": "근로기준법",
            "article": "제20조",
            "reason": "근로계약 불이행에 대한 위약금이나 손해배상액을 미리 정하는 계약을 금지합니다. 실제 손해에 대한 배상 책임과는 구분해서 검토해야 합니다."
        }
    ],

    "개인보호구 구입 비용 부담": [
        {
            "law": "산업안전보건기준에 관한 규칙",
            "article": "제32조",
            "reason": "법령에서 정한 위험 작업에 대해 사업주가 작업조건에 맞는 보호구를 지급하도록 규정하고 있습니다."
        }
    ],

    "안전관리 책임의 불명확성": [
        {
            "law": "산업안전보건법",
            "article": "제38조",
            "reason": "사업주의 안전조치 의무를 규정하고 있습니다."
        },
        {
            "law": "산업안전보건법",
            "article": "제39조",
            "reason": "사업주의 보건조치 의무를 규정하고 있습니다."
        }
    ]
}

# 안전보건 서약서 관련 법률 내용
SAFETY_LAW_TEXTS = {

    "근로기준법|제20조": {
        "date": "2026년 10월 8일 시행 기준",
        "text": (
            "사용자는 근로계약 불이행에 대한 위약금 또는 "
            "손해배상액을 미리 정하는 계약을 체결할 수 없습니다. "
            "다만 실제 발생한 손해에 대한 모든 배상 책임이 "
            "금지된다는 의미는 아닙니다."
        )
    },

    "산업안전보건법|제38조": {
        "date": "2026년 10월 27일 시행 기준",
        "text": (
            "사업주는 기계·설비, 위험물질, 전기 및 작업방법 등으로 "
            "인한 산업재해를 예방하기 위해 필요한 안전조치를 "
            "해야 합니다."
        )
    },

    "산업안전보건법|제39조": {
        "date": "2026년 10월 27일 시행 기준",
        "text": (
            "사업주는 유해물질, 소음, 진동, 고열 등으로 발생할 수 "
            "있는 근로자의 건강장해를 예방하기 위해 필요한 "
            "보건조치를 해야 합니다."
        )
    },

    "산업안전보건기준에 관한 규칙|제32조": {
        "date": "2026년 3월 2일 시행 기준",
        "text": (
            "사업주는 법령에서 정한 위험 작업을 하는 근로자에게 "
            "작업조건에 맞는 보호구를 지급하고 착용하도록 해야 합니다. "
            "근로자도 지급받은 보호구를 착용해야 합니다."
        )
    }
}

def analyze_safety_pledge(safety_text):

    safety_rules = [
        {
            "name": "사고 책임의 일괄 전가",
            "priority": "높음",
            "keyword_groups": [
                ["모든 사고 책임", "모든 책임", "사고에 대한 책임"],
                ["근로자 부담", "근로자가 부담", "근로자 책임"]
            ],
            "reason": "사고 원인이나 책임 소재를 구분하지 않고 근로자에게 책임을 부담시키는지 확인해야 합니다.",
            "suggestion": "사고 원인과 각 당사자의 책임을 조사한 후 관련 법령과 계약에 따라 책임을 판단하도록 수정하는 것이 좋습니다."
        },
        {
            "name": "손해배상 비용의 전액 부담",
            "priority": "높음",
            "keyword_groups": [
                ["모든 손해", "손해 전액", "모든 피해"],
                ["부담", "배상", "지급"]
            ],
            "reason": "근로자에게 손해배상 비용을 일괄적으로 부담시키는 내용인지 확인해야 합니다.",
            "suggestion": "손해 발생 원인과 책임 범위를 확인하고 관련 법령에 따라 배상 책임을 판단하도록 수정하는 것이 좋습니다."
        },
        {
            "name": "개인보호구 구입 비용 부담",
            "priority": "보통",
            "keyword_groups": [
                ["보호구", "안전장비", "안전용품"],
                ["구입 비용", "구매 비용", "구입비", "구매비", "비용을 부담"]
            ],
            "reason": "사업주가 부담해야 할 수 있는 보호구 지급 및 비용 책임을 근로자에게 전가하는지 확인해야 합니다.",
            "suggestion": "사업주의 보호구 지급 의무와 비용 부담 기준을 명확하게 작성하는 것이 좋습니다."
        },
        {
            "name": "안전관리 책임의 불명확성",
            "priority": "높음",
            "keyword_groups": [
                ["안전교육", "안전관리", "안전보건 조치"],
                ["사업주는 책임지지", "근로자 책임", "근로자가 전적으로 책임"]
            ],
            "reason": "사업주의 안전보건 조치 및 교육 책임이 불명확하거나 근로자에게 전가되는지 확인해야 합니다.",
            "suggestion": "사업주와 근로자의 안전보건 의무를 구분하여 명확하게 작성하는 것이 좋습니다."
        }
    ]

    results = []

    # 문장을 나누어 검사
    sentences = [
        sentence.strip()
        for sentence in safety_text.replace("。", ".").split(".")
        if sentence.strip()
    ]

    for rule in safety_rules:

        matched_keywords = []

        for group in rule["keyword_groups"]:
            found = [
                word for word in group
                if word in safety_text
            ]

            if found:
                matched_keywords.extend(found)
            else:
                break

        if all(
            any(word in safety_text for word in group)
            for group in rule["keyword_groups"]
        ):

            # 탐지된 키워드가 포함된 실제 문장 추출
            matched_sentences = [
                sentence
                for sentence in sentences
                if any(
                    word in sentence
                    for word in matched_keywords
                )
            ]

            results.append({
                "name": rule["name"],
                "priority": rule["priority"],
                "suggestion": rule["suggestion"],
                "found_keywords": matched_keywords,
                "original_text": matched_sentences,
                "reason": rule["reason"],

                "laws": SAFETY_LAW_REFERENCES.get(
                    rule["name"], []
                )
            })

    return results


# ==================================================
# 안전보건 서약서 분석 결과 화면
# ==================================================

if menu == "외국인 근로자 노무계약 미스매치 실드":

    if st.button("🦺 안전보건 서약서 위험 조항 분석하기"):

        safety_text = st.session_state.get("safety_text", "")

        if safety_text.strip():
            safety_results = analyze_safety_pledge(safety_text)
            st.session_state["safety_results"] = safety_results
        else:
            st.warning("먼저 안전보건 서약서를 업로드하고 문서를 읽어주세요.")

    if "safety_results" in st.session_state:

        results = st.session_state["safety_results"]

        st.divider()
        st.subheader("🦺 안전보건 서약서 분석 결과")

        if results:
            st.warning(f"추가 확인이 필요한 항목: {len(results)}개")

            for i, result in enumerate(results, 1):

                with st.container(border=True):

                    st.markdown(f"### {i}. {result['name']}")

                    if result["priority"] == "높음":
                        st.error("검토 우선순위: 높음")
                    else:
                        st.warning("검토 우선순위: 보통")

                    st.write("**탐지된 표현**")
                    st.write(", ".join(result["found_keywords"]))

                    st.write("**탐지된 원문**")

                    if result["original_text"]:
                        for sentence in result["original_text"]:
                            st.warning(sentence)
                    else:
                        st.write("원문 문장을 추출하지 못했습니다.")
                    
                     # 관련 법률 조항 표시
                    st.write("**관련 법률 조항**")

                    if result.get("laws"):
                        for law in result["laws"]:
                            st.write(f"법률명: {law['law']}")
                            st.write(f"관련 조항: {law['article']}")
                            st.write(f"검토 이유: {law['reason']}")
    
                            law_key = f"{law['law']}|{law['article']}"
                            law_data = globals().get("SAFETY_LAW_TEXTS", {}).get(law_key)

                            if law_data:
                                with st.expander("법률 조항 보기"):
                                    st.caption(f"기준: {law_data['date']}")
                                    st.write(law_data["text"])
                                    st.caption("※ 법률 조항의 핵심을 요약한 참고 자료입니다.")
                            else:
                                st.warning("등록된 법률 내용이 없습니다.")
                    else:
                        st.caption("연결된 법률 조항이 없습니다.")

        else:
            st.success("현재 등록된 규칙에서 추가 확인이 필요한 표현을 찾지 못했습니다.")

        st.caption(
            "본 분석은 정해진 표현을 탐지하는 기초 기능입니다. "
            "위법 여부를 확정하지 않으며 전문가의 검토가 필요합니다."
        )


# ==================================================
# 통합 분석 보고서 다운로드 기능
# ==================================================

if menu == "외국인 근로자 노무계약 미스매치 실드":

    st.divider()
    st.subheader("📄 통합 분석 보고서")

    if (
        "comparison_results" in st.session_state
        and "safety_results" in st.session_state
    ):

        report = Document()

        # 보고서 제목
        report.add_heading(
            "Lexai 외국인 근로자 계약 분석 보고서",
            level=0
        )

        report.add_paragraph(
            "Construction Contract Risk Analyzer"
        )

        report.add_paragraph(
            "본 보고서는 계약서 비교 및 안전보건 서약서 분석 결과를 정리한 자료입니다."
        )

        # 1. 근로계약서 비교 결과
        report.add_heading(
            "1. 한국어·외국어 근로계약서 비교 결과",
            level=1
        )

        comparison_results = st.session_state["comparison_results"]

        if comparison_results:
            for i, result in enumerate(comparison_results, 1):
                report.add_heading(
                    f"{i}. {result['name']}",
                    level=2
                )
                report.add_paragraph(
                    "확인 사유: " + result["reason"]
                )
        else:
            report.add_paragraph(
                "등록된 비교 규칙에서 차이 의심 항목이 발견되지 않았습니다."
            )

        # 2. 안전보건 서약서 분석 결과
        report.add_heading(
            "2. 안전보건 서약서 분석 결과",
            level=1
        )

        safety_results = st.session_state["safety_results"]

        if safety_results:
            for i, result in enumerate(safety_results, 1):
                report.add_heading(
                    f"{i}. {result['name']}",
                    level=2
                )
                report.add_paragraph(
                    "검토 우선순위: "
                    + str(result.get("priority", "미분류"))
                )
                report.add_paragraph(
                    "탐지된 표현: "
                    + ", ".join(result["found_keywords"])
                )

                report.add_paragraph(
                    "확인 사유: " + str(result.get("reason", "확인 사유 없음"))
                )

                report.add_paragraph(
                    "수정 제안: " + result["suggestion"]
                )
        else:
            report.add_paragraph(
                "등록된 규칙에서 추가 확인이 필요한 표현을 찾지 못했습니다."
            )

        # 3. 종합 안내
        report.add_heading("3. 종합 안내", level=1)

        report.add_paragraph(
            "본 보고서는 정해진 검색 규칙에 따라 위험 가능성이 있는 "
            "표현을 탐지한 기초 분석 자료입니다. "
            "실제 위법 여부나 법적 책임을 확정하지 않으며, "
            "계약서 전체 문맥과 관련 법령을 바탕으로 전문가의 "
            "추가 검토가 필요합니다."
        )

        # Word 파일 생성
        style_report(report)
        set_korean_font(report)

        report_buffer = BytesIO()
        report.save(report_buffer)
        report_buffer.seek(0)

        st.download_button(
            label="📥 통합 분석 보고서 다운로드",
            data=report_buffer,
            file_name="Lexai_통합분석보고서.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    else:
        st.info(
            "먼저 근로계약서 비교와 안전보건 서약서 분석을 "
            "모두 실행해 주세요."
        )

# -------------------------
# 하단 안내
# -------------------------

st.divider()

st.caption(
    "Lexai | AI-Powered Construction Contract Review"
)
