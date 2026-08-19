import io
import time
import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ----------------------------------------------------
# 1. 페이지 기본 설정
# ----------------------------------------------------
st.set_page_config(
    page_title="삼성생명 기획팀 동향분석 Agent",
    page_icon="📊",
    layout="wide"
)

# ----------------------------------------------------
# 2. Gem 시스템 지침 (Instructions) 설정
# ----------------------------------------------------
SYSTEM_INSTRUCTION = """
# Role
당신은 삼성생명 기획팀의 "동향분석 및 보고서 작성 전문 Agent"입니다.
국내 1위 생명보험사인 삼성생명의 기획팀 직원 관점에서 금융/보험/경제 동향 기사를 분석하고, 정교하고 전략적인 2페이지 규격 보고서를 작성하는 역할을 수행합니다.

# Core Workflow (4단계 보고서 작성 프로세스)
[Step 1. 메인 기사 분석]
- 입력된 URL 본문 또는 텍스트를 분석하여 주요 사실관계, 핵심 수치, 이슈 맥락을 정확히 파악합니다.

[Step 2. 심층 데이터 분석 및 배경 맥락 연계]
- 메인 이슈 관련 최신 업계 동향 및 금융당국 규제/제도 수치 (CSM, K-ICS 비율, 할인율, 무·저해지 해약률 등)
- 타 생보사 대응 현황 및 삼성생명 관련 사업/재무 비교 데이터 (CSM 잔액/신계약, 순이익, K-ICS, FC 규모 등)

[Step 3. 삼성생명 전략적 시사점 도출]
- 업계 1위 삼성생명 기획팀 관점에서 'So What(전략적 영향 및 대응 방안)'을 도출합니다.
- 아래 4대 핵심 요소를 반드시 포함하여 전략을 수립합니다:
  1) CSM(계약서비스마진) 질적 가치 확보 및 고마진 상품 포트폴리오 전략 (CSM 배수 관리 등)
  2) K-ICS(신지급여력제도) 비율 관리, 자본 효율성 극대화 및 장기 ALM 운용 전략
  3) 독보적 전속 FC 인프라 기반 컨설팅 역량 및 AI/디지털 영업 지원 차별화
  4) '2035 라이프케어 복합금융 플랫폼' 연결: 보험을 넘어 고객의 평생 리스크·건강·자산을 관리하는 삼성생명의 총체적 복합금융 생태계 구축

[Step 4. 규격화된 2페이지 보고서 작성]
- 아래 Output Structure & Formatting Rules를 철저히 준수하여 보고서를 출력합니다.

# Output Structure & Formatting Rules (출력 및 작성 규격)
1. 문체 및 작성 원칙:
   - 위계 구조: `ㅁ` (대항목) -> `-` (중항목) -> `.` (소항목)
   - 대항목 명칭: 반드시 `ㅁ [요약]` 및 `ㅁ [시사점]` 표기 준수
   - 문체: 명확하고 격식 있는 보고서용 개조식 명사형 종결문 (~확대, ~추진, ~견지, ~구축, ~달성 등)
   - 문장 길이 및 밀도: 워드 기준 15pt로 작성 시 각 항목이 한 줄~최대 2줄 내로 들어오도록 불필요한 수식어를 배제하고 고밀도·컴팩트하게 서술

2. 보고서 본문 구성:
---
# [Page 1] 동향 보고서

ㅁ [요약] 이슈 핵심 제목
  - 메인 기사로 확인된 핵심 사실관계 (수치 및 일자 포함)
    . 삼성생명 및 1위사 관련 세부 통계 데이터
    . 신계약 드라이브 및 주요 세부 지표
  - 업계/시장 변화 및 주요 타사 대응 동향
    . 경쟁사 약진 및 판매 전략 동향
    . 금융당국 규제/제도 기조 및 건전성 영향

ㅁ [시사점] 초격차 확대를 위한 4대 핵심 전략 방향
  - CSM 질적 가치 제고 및 업계 최고 K-ICS 기반 자본 효율성 극대화
    . 상품 세분화 및 시니어 특화 라인업을 통한 신계약 CSM 배수 관리
    . 장기 ALM 매칭 정밀화 및 글로벌 대체투자·배당수익 다변화
  - 4.5만 전속 FC 파워 및 '2035 라이프케어 복합금융 플랫폼' 생태계 선점
    . 독보적 FC망 기반 AI 컨설팅 인프라 탑재 및 유지율 개선(예실차 관리)
    . 보험을 넘어 고객의 평생 리스크·건강·자산을 관리하는 총체적 복합금융 생태계 구축

<표 1> 대형 생보 3사(또는 타사/이슈 대상 vs 삼성생명) 주요 재무 및 경영 지표 비교
(Markdown Table 형식으로 작성: 지표 구분 | 타사 1 | 타사 2 | 삼성생명 (1위) | 시사점 및 격차 분석)

---
# [Page 2] 참고 자료 및 심층 출처 리스트 (References & Deep Dive Data)

1. [메인 기사 출처]
  - 출처명: 기사 제목
    . 핵심 요약: 1~2줄 컴팩트 요약

2. [심층 배경 및 비교 데이터]
  - 주요 지표 및 제도적 맥락
    . 핵심 데이터: 1~2줄 핵심 수치 요약

3. [기획팀 종합 평가 및 향후 모니터링 포인트]
  - 초격차 지배력: 시장 리더십 및 절대 규모 격차 평가
  - 리스크 선제 대응: 금리/규제 환경 변화에 따른 중점 관리 요소
  - 미래 성장동력: 2035 라이프케어 복합금융 플랫폼 조기 구축 과제
"""

# ----------------------------------------------------
# 3. Word 문서 생성 유틸리티
# ----------------------------------------------------
def create_docx(text_content):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    title = doc.add_heading("삼성생명 기획팀 동향분석 보고서", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    lines = text_content.split("\n")
    for line in lines:
        if line.startswith("# "):
            doc.add_heading(line.replace("# ", ""), level=1)
        elif line.startswith("## "):
            doc.add_heading(line.replace("## ", ""), level=2)
        elif line.startswith("ㅁ"):
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 51, 102)
        elif line.strip().startswith("-"):
            p = doc.add_paragraph(line)
            p.paragraph_format.left_indent = Inches(0.2)
        elif line.strip().startswith("."):
            p = doc.add_paragraph(line)
            p.paragraph_format.left_indent = Inches(0.4)
        else:
            if line.strip():
                doc.add_paragraph(line)
                
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# ----------------------------------------------------
# 4. 부서원 비밀번호 인증
# ----------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 기획팀 전용 분석 Agent")
    pwd = st.text_input("부서 접속 비밀번호를 입력하세요", type="password")
    correct_pwd = st.secrets.get("AUTH_PASSWORD", "1234")
    
    if st.button("접속", use_container_width=True):
        if pwd == correct_pwd:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")
    st.stop()

# ----------------------------------------------------
# 5. 메인 앱 화면 및 API 호출 (Fallback & Retry 로직 적용)
# ----------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY")
if not api_key:
    st.error("GEMINI_API_KEY가 설정되지 않았습니다.")
    st.stop()

client = genai.Client(api_key=api_key)

st.title("📊 삼성생명 기획팀 동향분석 Agent")
st.caption("기사 URL 또는 본문 텍스트를 입력하면 정규 2페이지 보고서 및 Word/TXT 파일을 생성합니다.")

user_input = st.text_area(
    "분석할 기사 내용 또는 기사 URL을 입력하세요:",
    height=150,
    placeholder="예: 기사 전문을 붙여넣거나 핵심 이슈 텍스트 및 링크를 입력하세요."
)

if st.button("보고서 생성 시작", type="primary", use_container_width=True):
    if not user_input.strip():
        st.warning("분석할 기사 내용이나 URL을 입력해주세요.")
    else:
        with st.spinner("기획팀 규격에 맞추어 2페이지 동향분석 보고서를 작성 중입니다..."):
            # 우선순위 순서대로 모델 시도
            models_to_try = [
                "gemini-2.5-pro",
                "gemini-2.5-flash",
                "gemini-1.5-pro",
                "gemini-1.5-flash"
            ]
            response_success = False
            last_error_msg = ""

            for model_name in models_to_try:
                for attempt in range(2):  # 각 모델별 2회 시도
                    try:
                        response = client.models.generate_content(
                            model=model_name,
                            contents=user_input,
                            config=types.GenerateContentConfig(
                                system_instruction=SYSTEM_INSTRUCTION,
                                temperature=0.3
                            )
                        )
                        st.session_state["last_report"] = response.text
                        response_success = True
                        break
                    except Exception as e:
                        last_error_msg = str(e)
                        # 일시적인 503 또는 과부하 시 2초 대기 후 재시도
                        if "503" in last_error_msg or "UNAVAILABLE" in last_error_msg:
                            time.sleep(2)
                            continue
                        else:
                            break
                if response_success:
                    break

            if not response_success:
                st.error(f"구글 API 서버 응답 지연으로 생성을 완료하지 못했습니다: {last_error_msg}")

# ----------------------------------------------------
# 6. 생성된 보고서 표시 및 다운로드
# ----------------------------------------------------
if "last_report" in st.session_state:
    st.divider()
    st.markdown("### 📄 작성된 동향분석 보고서")
    st.markdown(st.session_state["last_report"])
    
    st.divider()
    col1, col2 = st.columns(2)
    
    with col1:
        docx_file = create_docx(st.session_state["last_report"])
        st.download_button(
            label="📥 Word 보고서 (.docx) 다운로드",
            data=docx_file,
            file_name="삼성생명_동향분석보고서.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with col2:
        st.download_button(
            label="📥 텍스트 보고서 (.txt) 다운로드",
            data=st.session_state["last_report"].encode("utf-8"),
            file_name="삼성생명_동향분석보고서.txt",
            mime="text/plain",
            use_container_width=True
        )
