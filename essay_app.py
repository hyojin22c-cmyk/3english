import streamlit as st
import anthropic
import os
import hashlib
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# ── 페이지 설정 ──────────────────────────────────────────
st.set_page_config(
    page_title="에세이 개요 도우미",
    page_icon="✍️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── 스타일 (독작용 — 남색 톤) ────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@400;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap');

:root {
    --bg: #F4F1EE;
    --surface: #FFFFFF;
    --border: #D4CFC6;
    --accent: #1B3A5C;
    --accent-light: #2A5580;
    --accent-pale: #E4EDF5;
    --text: #1A1A1A;
    --text-muted: #5C6470;
    --danger: #C0392B;
}

html, body, .stApp {
    background-color: var(--bg) !important;
    font-family: 'Noto Sans KR', sans-serif;
    color: var(--text);
}

.main-header {
    text-align: center;
    padding: 2.5rem 0 1.5rem;
    border-bottom: 2px solid var(--accent);
    margin-bottom: 2rem;
}
.main-header h1 {
    font-family: 'Noto Serif KR', serif;
    font-size: 2rem;
    font-weight: 700;
    color: var(--accent);
    margin: 0;
    letter-spacing: -0.5px;
}
.main-header p {
    color: var(--text-muted);
    font-size: 0.9rem;
    margin-top: 0.4rem;
    font-weight: 300;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    background: var(--surface);
    border-radius: 8px 8px 0 0;
    border: 1px solid var(--border);
    border-bottom: none;
    padding: 0;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Noto Sans KR', sans-serif;
    font-weight: 500;
    font-size: 0.9rem;
    padding: 0.75rem 1.5rem;
    color: var(--text-muted);
    border-radius: 0;
}
.stTabs [aria-selected="true"] {
    background: var(--accent) !important;
    color: white !important;
}
.stTabs [data-baseweb="tab-panel"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-top: none;
    border-radius: 0 0 8px 8px;
    padding: 1.5rem;
}

.passage-card {
    background: var(--bg);
    border: 1px solid var(--border);
    border-left: 4px solid var(--accent);
    border-radius: 6px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.8rem;
}
.passage-card h4 {
    font-family: 'Noto Serif KR', serif;
    font-size: 1rem;
    color: var(--accent);
    margin: 0 0 0.3rem;
}

.stTextInput input, .stTextArea textarea, .stSelectbox select {
    border: 1px solid var(--border) !important;
    border-radius: 6px !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    background: var(--bg) !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 2px rgba(27, 58, 92, 0.1) !important;
}

.stButton button {
    background: var(--accent) !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    font-weight: 500 !important;
    padding: 0.5rem 1.5rem !important;
    transition: background 0.2s !important;
}
.stButton button:hover {
    background: var(--accent-light) !important;
}

.delete-btn button {
    background: transparent !important;
    color: var(--danger) !important;
    border: 1px solid var(--danger) !important;
    font-size: 0.8rem !important;
    padding: 0.25rem 0.75rem !important;
}
.delete-btn button:hover {
    background: var(--danger) !important;
    color: white !important;
}

hr { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }
.stSuccess, .stError, .stWarning, .stInfo { border-radius: 6px !important; }
.stMultiSelect [data-baseweb="tag"] { background: var(--accent) !important; }
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
</style>
""", unsafe_allow_html=True)

# ── 유틸리티 ─────────────────────────────────────────────
def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.strip().encode()).hexdigest()

def make_cache_key(selected_titles, career, interests, step) -> str:
    raw = f"{sorted(selected_titles)}|{career.strip()}|{interests.strip()}|{step}"
    return hashlib.md5(raw.encode()).hexdigest()

# ── Google Sheets 연동 (시트명 "독작_" 접두사) ────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_gspread_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )
    return gspread.authorize(creds)

def _get_spreadsheet():
    client = get_gspread_client()
    return client.open_by_key(st.secrets["SHEET_ID"])

def _get_or_create_sheet(name, cols, rows=500):
    spreadsheet = _get_spreadsheet()
    try:
        return spreadsheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.add_worksheet(title=name, rows=rows, cols=len(cols))
        sheet.append_row(cols)
        return sheet

def get_passage_sheet():
    return _get_or_create_sheet("독작_지문", ["id", "title", "summary", "keywords"])

def get_auth_sheet():
    return _get_or_create_sheet("독작_학생인증", ["학번", "이름", "비밀번호"])

def get_log_sheet():
    return _get_or_create_sheet("독작_사용기록", ["날짜", "학번", "이름", "선택지문", "진로", "결과"])

def get_cache_sheet():
    return _get_or_create_sheet("독작_캐시", ["key", "result", "created"])

def get_bonus_sheet():
    return _get_or_create_sheet("독작_추가횟수", ["학번", "횟수", "적용월"])

# ── 지문 CRUD ────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_passages():
    try:
        sheet = get_passage_sheet()
        return sheet.get_all_records()
    except Exception as e:
        st.error(f"시트 로드 실패: {e}")
        return []

def save_passage(passage):
    try:
        sheet = get_passage_sheet()
        sheet.append_row([
            passage["id"],
            passage["title"],
            passage["summary"],
            passage.get("keywords", "")
        ])
        load_passages.clear()
        clear_result_cache()
    except Exception as e:
        st.error(f"저장 실패: {e}")

def delete_passage(passage_id):
    try:
        sheet = get_passage_sheet()
        all_values = sheet.get_all_values()
        for i, row in enumerate(all_values):
            if row and str(row[0]) == str(passage_id):
                sheet.delete_rows(i + 1)
                break
        load_passages.clear()
        clear_result_cache()
    except Exception as e:
        st.error(f"삭제 실패: {e}")

# ── 학생 인증 ────────────────────────────────────────────
def find_student(student_id):
    try:
        sheet = get_auth_sheet()
        rows = sheet.get_all_records()
        for row in rows:
            if str(row.get("학번", "")) == str(student_id):
                return row
        return None
    except Exception:
        return None

def register_student(student_id, name, password):
    try:
        sheet = get_auth_sheet()
        sheet.append_row([student_id, name, hash_pw(password)])
        return True
    except Exception:
        return False

def verify_password(stored_hash, input_pw):
    if stored_hash == hash_pw(input_pw):
        return True
    if stored_hash == input_pw:
        return True
    return False

# ── 사용 제한 (월별 자동 초기화) ─────────────────────────
BASE_MONTHLY_LIMIT = 4

def get_student_limit(student_id):
    this_month = datetime.now().strftime("%Y-%m")
    try:
        sheet = get_bonus_sheet()
        rows = sheet.get_all_records()
        extra = 0
        for row in rows:
            if str(row.get("학번", "")) == str(student_id) and str(row.get("적용월", "")) == this_month:
                try:
                    extra += int(row.get("횟수", 0) or 0)
                except (ValueError, TypeError):
                    pass
        return BASE_MONTHLY_LIMIT + extra
    except Exception:
        return BASE_MONTHLY_LIMIT

def grant_extra_usage(student_id, extra_count):
    this_month = datetime.now().strftime("%Y-%m")
    try:
        sheet = get_bonus_sheet()
        sheet.append_row([str(student_id), extra_count, this_month])
        return True, get_student_limit(student_id) - BASE_MONTHLY_LIMIT
    except Exception:
        return False, 0

def reset_extra_usage(student_id):
    this_month = datetime.now().strftime("%Y-%m")
    try:
        sheet = get_bonus_sheet()
        all_values = sheet.get_all_values()
        rows_to_delete = []
        for i, row in enumerate(all_values):
            if i == 0:
                continue
            if row and str(row[0]) == str(student_id) and len(row) > 2 and str(row[2]) == this_month:
                rows_to_delete.append(i + 1)
        for row_idx in reversed(rows_to_delete):
            sheet.delete_rows(row_idx)
        return True
    except Exception:
        return False

@st.cache_data(ttl=120)
def check_monthly_usage(student_id):
    try:
        sheet = get_log_sheet()
        rows = sheet.get_all_records()
        this_month = datetime.now().strftime("%Y-%m")
        return sum(1 for row in rows
                   if str(row.get("학번", "")) == str(student_id)
                   and str(row.get("날짜", ""))[:7] == this_month)
    except Exception:
        return 0

def save_usage_log(student_id, name, selected_titles, career, result_text):
    try:
        sheet = get_log_sheet()
        truncated = result_text[:3000] if result_text else ""
        sheet.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            student_id,
            name,
            " / ".join(selected_titles),
            career,
            truncated
        ])
        check_monthly_usage.clear()
    except Exception as e:
        st.error(f"기록 저장 실패: {e}")

# ── 결과 캐싱 ────────────────────────────────────────────
def get_cached_result(cache_key):
    try:
        sheet = get_cache_sheet()
        rows = sheet.get_all_records()
        for row in rows:
            if str(row.get("key", "")) == cache_key:
                return row.get("result", "")
        return None
    except Exception:
        return None

def save_cached_result(cache_key, result_text):
    try:
        sheet = get_cache_sheet()
        sheet.append_row([cache_key, result_text[:10000], datetime.now().strftime("%Y-%m-%d %H:%M")])
    except Exception:
        pass

def clear_result_cache():
    try:
        sheet = get_cache_sheet()
        all_values = sheet.get_all_values()
        if len(all_values) > 1:
            sheet.delete_rows(2, len(all_values))
    except Exception:
        pass

# ── Claude API ────────────────────────────────────────────
def get_claude_client():
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "") or os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return None
    return anthropic.Anthropic(api_key=api_key)

# ── 프롬프트: 에세이 주제 + 개요 추천 ────────────────────
MAX_SUMMARY_LEN = 100

def build_topic_prompt(selected_passages, career, interests):
    passage_text = ""
    for i, p in enumerate(selected_passages, 1):
        summary = p.get("summary", "")
        keywords = p.get("keywords", "")
        if len(summary) > MAX_SUMMARY_LEN:
            summary = summary[:MAX_SUMMARY_LEN] + "…"
        passage_text += f"{i}. {p['title']}\n"
        passage_text += f"   요약: {summary}\n"
        if keywords:
            passage_text += f"   키워드: {keywords}\n"
        passage_text += "\n"

    return f"""당신은 고등학교 3학년 영어 독해와 작문 수업의 에세이 작성을 도와주는 전문가입니다.

[학생이 선택한 수업 지문]
{passage_text}

[학생 정보]
- 희망 진로/관심 분야: {career if career else '미입력'}
- 추가 관심사: {interests if interests else '미입력'}

위 지문의 주제·내용을 바탕으로, 학생의 진로와 연결할 수 있는 영어 에세이 주제를 정확히 3개 추천하세요.
이 에세이는 수행평가이자 생기부 심화 보고서로도 활용됩니다.

각 주제는 아래 형식으로 작성:

**[주제 번호]. 에세이 주제 (영문 제목)**

📌 **주제 설명**: 이 주제가 무엇을 다루는지, 왜 이 지문과 연결되는지 2~3문장

📝 **에세이 개요**:
- **서론 (Introduction)**: 어떤 문제의식이나 질문으로 시작할지, 독자의 관심을 끄는 방법 제안 (2~3문장)
- **본론 1 (Body 1)**: 첫 번째 핵심 논점과 다룰 내용 (2~3문장)
- **본론 2 (Body 2)**: 두 번째 핵심 논점과 다룰 내용 (2~3문장)
- **결론 (Conclusion)**: 마무리 방향과 주장 요약 방법 (2~3문장)

🔑 **핵심 영어 표현**: 에세이에서 활용하면 좋을 영어 표현 3~4개 (표현 — 한국어 뜻)

[중요 규칙]
- 에세이 본문을 직접 작성하지 마세요. 개요와 방향만 제시합니다.
- 학생이 스스로 작성할 수 있도록 구체적이되 완성된 문장은 주지 않습니다.
- 고등학생 수준에서 실현 가능한 주제로 제안하세요.
- 영어 에세이이므로 영문 제목과 영어 표현을 포함하되, 설명은 한국어로 합니다."""

def build_refine_prompt(chosen_topic, student_memo):
    return f"""학생이 아래 에세이 주제와 개요를 선택했고, 추가 요청사항이 있습니다.

[선택한 주제와 개요]
{chosen_topic}

[학생 요청사항]
{student_memo}

학생의 요청을 반영하여 개요를 수정·보완해주세요.

[중요 규칙]
- 에세이 본문을 직접 작성하지 마세요. 개요와 방향만 수정합니다.
- 수정된 개요를 원래와 같은 형식(서론-본론1-본론2-결론)으로 제시합니다.
- 변경된 부분을 간단히 설명해주세요.
- 핵심 영어 표현도 요청에 맞게 업데이트해주세요."""

# ── 세션 초기화 ───────────────────────────────────────────
if "passages" not in st.session_state:
    st.session_state.passages = load_passages()
if "result" not in st.session_state:
    st.session_state.result = None
if "refine_result" not in st.session_state:
    st.session_state.refine_result = None
if "auth_student" not in st.session_state:
    st.session_state.auth_student = None

# ── 헤더 ─────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>✍️ 에세이 개요 도우미</h1>
    <p>독해와 작문 수업 지문을 바탕으로 나만의 에세이 주제와 개요를 잡아보세요</p>
</div>
""", unsafe_allow_html=True)

# ── 탭 구성 ──────────────────────────────────────────────
tab_student, tab_admin = st.tabs(["✏️ 에세이 주제 잡기", "⚙️ 지문 관리 (선생님)"])

# ════════════════════════════════════════════════════════
# 학생용 탭
# ════════════════════════════════════════════════════════
with tab_student:
    passages = st.session_state.passages

    if not passages:
        st.info("📌 아직 등록된 지문이 없어요. 선생님께 지문을 등록해달라고 요청하세요!")
    else:
        # ── 로그인/등록 ──
        if not st.session_state.auth_student:
            st.markdown("#### 🔐 로그인 / 최초 등록")
            auth_mode = st.radio("", ["로그인", "최초 등록"], horizontal=True, label_visibility="collapsed")

            aid = st.text_input("학번", placeholder="예: 30101")
            apw = st.text_input("비밀번호", type="password", placeholder="본인이 설정한 비밀번호")

            if auth_mode == "최초 등록":
                aname = st.text_input("이름", placeholder="예: 홍길동")
                apw2 = st.text_input("비밀번호 확인", type="password")
                if st.button("등록하기", use_container_width=True):
                    if not aid or not aname or not apw:
                        st.warning("모든 항목을 입력해주세요.")
                    elif apw != apw2:
                        st.error("비밀번호가 일치하지 않습니다.")
                    elif find_student(aid):
                        st.error("이미 등록된 학번입니다. 로그인을 이용해주세요.")
                    else:
                        if register_student(aid, aname, apw):
                            st.session_state.auth_student = {"학번": aid, "이름": aname}
                            st.success(f"✅ {aname}님 등록 완료!")
                            st.rerun()
                        else:
                            st.error("등록 실패. 다시 시도해주세요.")
            else:
                if st.button("로그인", use_container_width=True):
                    if not aid or not apw:
                        st.warning("학번과 비밀번호를 입력해주세요.")
                    else:
                        student = find_student(aid)
                        if not student:
                            st.error("등록되지 않은 학번입니다. 최초 등록을 먼저 해주세요.")
                        elif not verify_password(str(student.get("비밀번호", "")), apw):
                            st.error("비밀번호가 틀렸습니다.")
                        else:
                            st.session_state.auth_student = {"학번": aid, "이름": student.get("이름", "")}
                            st.session_state.result = None
                            st.session_state.refine_result = None
                            st.rerun()

        # ── 로그인 후 메인 화면 ──
        else:
            student = st.session_state.auth_student
            monthly = check_monthly_usage(student["학번"])
            limit = get_student_limit(student["학번"])

            col1, col2 = st.columns([1, 1.2], gap="large")

            with col1:
                st.markdown(f"#### 👋 {student['이름']}님 환영해요")
                st.caption(f"이번 달 {monthly}/{limit}회 사용")

                st.markdown("---")
                st.markdown("#### 📚 지문 선택")
                st.caption("에세이에 활용할 지문을 선택하세요 (1~3개 권장)")

                passage_titles = [p["title"] for p in passages]
                selected_titles = st.multiselect(
                    "지문 선택",
                    passage_titles,
                    max_selections=3,
                    label_visibility="collapsed"
                )

                # 선택된 지문 상세 표시
                if selected_titles:
                    for p in passages:
                        if p["title"] in selected_titles:
                            keywords = p.get("keywords", "")
                            st.markdown(f"""
                            <div class="passage-card">
                                <h4>{p['title']}</h4>
                                <small style="color:#5C6470;">{p.get('summary','')[:80]}...</small>
                                {f"<br><small style='color:#1B3A5C;'>🔑 {keywords}</small>" if keywords else ""}
                            </div>
                            """, unsafe_allow_html=True)

                st.markdown("---")
                career = st.text_input("희망 진로 / 관심 분야", placeholder="예: 국제기구, 환경공학, 심리상담...")
                interests = st.text_input("추가 관심사 (선택)", placeholder="예: 기후변화, 인공지능, 사회적 불평등...")

                if st.button("🚪 로그아웃", use_container_width=True):
                    st.session_state.auth_student = None
                    st.session_state.result = None
                    st.session_state.refine_result = None
                    st.rerun()

            with col2:
                st.markdown("#### ✨ 에세이 주제 & 개요 추천")

                # ── STEP 1: 주제 + 개요 추천 ──
                if st.button("✨ 주제 추천 받기", use_container_width=True):
                    if not selected_titles:
                        st.warning("지문을 1개 이상 선택해주세요!")
                    elif not career:
                        st.warning("희망 진로 / 관심 분야를 입력해주세요!")
                    elif monthly >= limit:
                        st.error(f"⚠️ 이번 달 사용 횟수({monthly}/{limit}회)를 초과했습니다.")
                    else:
                        selected_passages = [p for p in passages if p["title"] in selected_titles]
                        cache_key = make_cache_key(selected_titles, career, interests, "topic")
                        cached = get_cached_result(cache_key)

                        if cached:
                            st.session_state.result = cached
                            st.session_state.refine_result = None
                            save_usage_log(student["학번"], student["이름"], selected_titles, career, cached)
                            st.caption(f"💡 이번 달 {monthly + 1}/{limit}회 사용 (캐시 활용)")
                        else:
                            client = get_claude_client()
                            if not client:
                                st.error("API 키가 설정되지 않았습니다.")
                            else:
                                with st.spinner("에세이 주제를 고민하는 중..."):
                                    try:
                                        prompt = build_topic_prompt(selected_passages, career, interests)
                                        message = client.messages.create(
                                            model="claude-sonnet-4-6",
                                            max_tokens=2048,
                                            messages=[{"role": "user", "content": prompt}]
                                        )
                                        result_text = message.content[0].text
                                        st.session_state.result = result_text
                                        st.session_state.refine_result = None

                                        save_cached_result(cache_key, result_text)
                                        save_usage_log(student["학번"], student["이름"], selected_titles, career, result_text)
                                        st.caption(f"💡 이번 달 {monthly + 1}/{limit}회 사용했습니다.")

                                    except anthropic.RateLimitError:
                                        st.error("⏳ 요청이 많아요. 30초 후 다시 시도해주세요.")
                                    except anthropic.AuthenticationError:
                                        st.error("🔑 API 키에 문제가 있습니다. 선생님께 문의하세요.")
                                    except anthropic.APIError as e:
                                        st.error(f"API 오류가 발생했습니다: {e}")
                                    except Exception as e:
                                        st.error(f"예상치 못한 오류: {e}")

                # ── 주제 추천 결과 표시 ──
                if st.session_state.result:
                    st.markdown(st.session_state.result)

                    st.markdown("---")

                    # ── STEP 2: 개요 다듬기 (1회 피드백) ──
                    st.markdown("#### 🔄 개요 다듬기")
                    st.caption("위 주제 중 하나를 골라 개요를 수정하고 싶다면 아래에 요청하세요. (1회)")

                    chosen = st.text_input("선택한 주제 번호", placeholder="예: 1")
                    memo = st.text_area(
                        "수정 요청사항",
                        placeholder="예: 본론에서 한국 사례를 추가하고 싶어요 / 서론을 질문 형식으로 바꾸고 싶어요 / 환경 문제보다 경제적 관점에서 다루고 싶어요",
                        height=100
                    )

                    if st.button("🔄 개요 다듬기", use_container_width=True):
                        if not chosen or not memo:
                            st.warning("주제 번호와 수정 요청을 모두 입력해주세요.")
                        elif st.session_state.refine_result:
                            st.info("개요 다듬기는 1회만 가능합니다.")
                        elif monthly >= limit:
                            st.error(f"⚠️ 이번 달 사용 횟수를 초과했습니다.")
                        else:
                            client = get_claude_client()
                            if client:
                                with st.spinner("개요를 다듬는 중..."):
                                    try:
                                        # 선택한 주제 부분 추출 (전체 결과에서)
                                        refine_prompt = build_refine_prompt(
                                            f"주제 {chosen}번:\n{st.session_state.result}",
                                            memo
                                        )
                                        message = client.messages.create(
                                            model="claude-sonnet-4-6",
                                            max_tokens=1500,
                                            messages=[{"role": "user", "content": refine_prompt}]
                                        )
                                        st.session_state.refine_result = message.content[0].text
                                        save_usage_log(
                                            student["학번"], student["이름"],
                                            selected_titles, career,
                                            f"[다듬기] {message.content[0].text}"
                                        )
                                    except Exception as e:
                                        st.error(f"오류: {e}")

                    if st.session_state.refine_result:
                        st.markdown("---")
                        st.markdown("#### ✅ 수정된 개요")
                        st.markdown(st.session_state.refine_result)

                    # ── 다운로드 ──
                    st.markdown("---")
                    download_text = st.session_state.result
                    if st.session_state.refine_result:
                        download_text += "\n\n" + "=" * 50 + "\n[수정된 개요]\n" + "=" * 50 + "\n\n"
                        download_text += st.session_state.refine_result

                    st.download_button(
                        label="📥 결과 저장 (txt)",
                        data=download_text,
                        file_name=f"에세이개요_{student['이름']}_{datetime.now().strftime('%Y%m%d')}.txt",
                        mime="text/plain",
                        use_container_width=True
                    )

# ════════════════════════════════════════════════════════
# 관리자 탭
# ════════════════════════════════════════════════════════
with tab_admin:
    if "admin_auth" not in st.session_state:
        st.session_state.admin_auth = False

    if not st.session_state.admin_auth:
        st.markdown("#### 🔐 관리자 로그인")
        pw = st.text_input("비밀번호", type="password", key="admin_pw_input")
        admin_pw = st.secrets.get("ADMIN_PASSWORD", "teacher1234")
        if st.button("로그인", key="admin_login_btn"):
            if pw == admin_pw:
                st.session_state.admin_auth = True
                st.rerun()
            else:
                st.error("비밀번호가 틀렸습니다.")
    else:
        col_form, col_list = st.columns([1, 1.2], gap="large")

        with col_form:
            st.markdown("#### ➕ 지문 추가")

            new_title = st.text_input("지문 제목 *", placeholder="예: The Ethics of Genetic Engineering")
            new_summary = st.text_area(
                "지문 내용 요약 *",
                placeholder="예: 유전자 편집 기술의 발전과 그에 따른 윤리적 쟁점을 다룬 지문. 디자이너 베이비, 유전 질환 치료, 사회적 불평등 심화 가능성 등을 논의함.",
                height=120
            )
            new_keywords = st.text_input(
                "핵심 키워드",
                placeholder="예: genetic engineering, ethics, CRISPR, designer babies"
            )
            st.caption("💡 키워드는 학생이 지문을 선택할 때 참고할 수 있도록 입력해주세요.")

            if st.button("💾 지문 저장", use_container_width=True):
                if not new_title or not new_summary:
                    st.warning("제목과 요약은 필수입니다!")
                else:
                    new_passage = {
                        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                        "title": new_title,
                        "summary": new_summary,
                        "keywords": new_keywords
                    }
                    save_passage(new_passage)
                    st.session_state.passages = load_passages()
                    st.success(f"✅ '{new_title}' 추가 완료!")
                    st.rerun()

            st.markdown("---")
            st.markdown("#### 🎫 학생 추가 횟수 부여")
            st.caption("이번 달에만 적용됩니다. 다음 달에는 자동으로 기본 4회로 돌아갑니다.")

            bonus_id = st.text_input("학번", placeholder="예: 30101", key="bonus_student_id")
            bonus_count = st.number_input("추가할 횟수", min_value=1, max_value=20, value=2, key="bonus_count")

            bc1, bc2 = st.columns(2)
            with bc1:
                if st.button("➕ 횟수 부여", use_container_width=True, key="grant_btn"):
                    if not bonus_id:
                        st.warning("학번을 입력해주세요.")
                    else:
                        target = find_student(bonus_id)
                        if not target:
                            st.error("등록되지 않은 학번입니다.")
                        else:
                            ok, total = grant_extra_usage(bonus_id, bonus_count)
                            if ok:
                                this_month = datetime.now().strftime("%Y년 %m월")
                                new_limit = get_student_limit(bonus_id)
                                st.success(f"✅ {target.get('이름', '')}({bonus_id}) — {this_month} 총 {new_limit}회 사용 가능")
                            else:
                                st.error("부여 실패. 다시 시도해주세요.")
            with bc2:
                if st.button("🔄 추가분 초기화", use_container_width=True, key="reset_btn"):
                    if not bonus_id:
                        st.warning("학번을 입력해주세요.")
                    else:
                        target = find_student(bonus_id)
                        if not target:
                            st.error("등록되지 않은 학번입니다.")
                        else:
                            if reset_extra_usage(bonus_id):
                                st.success(f"✅ {target.get('이름', '')}({bonus_id}) — 이번 달 기본 {BASE_MONTHLY_LIMIT}회로 초기화")
                            else:
                                st.error("초기화 실패.")

            st.markdown("---")
            if st.button("🚪 로그아웃", use_container_width=True, key="admin_logout"):
                st.session_state.admin_auth = False
                st.rerun()

        with col_list:
            st.markdown(f"#### 📋 등록된 지문 ({len(st.session_state.passages)}개)")

            if not st.session_state.passages:
                st.info("등록된 지문이 없습니다.")
            else:
                for i, p in enumerate(st.session_state.passages):
                    with st.container():
                        c1, c2 = st.columns([4, 1])
                        with c1:
                            summary_preview = str(p.get("summary", ""))[:60]
                            keywords = p.get("keywords", "")
                            st.markdown(f"""
                            <div class="passage-card">
                                <h4>{p['title']}</h4>
                                <small style="color:#5C6470;">{summary_preview}...</small>
                                {f"<br><small style='color:#1B3A5C;'>🔑 {keywords}</small>" if keywords else ""}
                            </div>
                            """, unsafe_allow_html=True)
                        with c2:
                            st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
                            if st.button("삭제", key=f"del_{p['id']}"):
                                delete_passage(p["id"])
                                st.session_state.passages = [x for x in st.session_state.passages if x["id"] != p["id"]]
                                st.rerun()
                            st.markdown('</div>', unsafe_allow_html=True)
