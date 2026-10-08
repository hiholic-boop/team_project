"""Streamlit interface for the complete sales diagnosis workflow."""

from __future__ import annotations

import hashlib
import re
from datetime import date
from pathlib import Path

import streamlit as st

from alm_app import (
    MAX_AUDIO_BYTES,
    ApiError,
    configured_api_key,
    find_target,
    normalize_date,
    read_standard_sections,
    review_transcript,
    transcribe,
    validate_upload_filename,
)


BROCHURE_PATH = Path(__file__).with_name("보험_TM_완전판매_점검_한장_설명서_공유용.html")

st.set_page_config(page_title="TM Quality Check | 완전판매 점검", page_icon="●", layout="wide")

st.markdown(
    """
    <style>
    :root { --ink:#0a0a0a; --paper:#f6f6f4; --line:#d8d8d5; --signal:#e60012; }
    html, body, [class*="css"], [data-testid="stAppViewContainer"] {
      font-family:"Helvetica Neue",Arial,"Malgun Gothic","Apple SD Gothic Neo",sans-serif;
      letter-spacing:-.025em;
    }
    .stApp { background:var(--paper); color:var(--ink); }
    [data-testid="stHeader"] { background:rgba(246,246,244,.92); }
    [data-testid="stMainBlockContainer"] { max-width:1120px; padding-top:2.5rem; }
    h1,h2,h3 { color:var(--ink); letter-spacing:-.055em; }
    h1 { font-weight:800; }
    [data-testid="stVerticalBlockBorderWrapper"] > div { border-color:var(--line); border-radius:0; background:#fff; }
    [data-testid="stFileUploader"] { border-radius:0; }
    [data-testid="stFileUploaderDropzone"] { background:var(--paper); border:1px dashed #a6a6a2; border-radius:0; }
    [data-testid="stTextInput"] input, [data-testid="stDateInput"] input,
    [data-testid="stTextArea"] textarea { border-radius:0; background:#fff; color:var(--ink); }
    [data-testid="stTextInput"] input:focus, [data-testid="stDateInput"] input:focus,
    [data-testid="stTextArea"] textarea:focus { border-color:var(--signal); }
    [data-testid="stExpander"] { border:1px solid var(--line); border-radius:0; background:#fff; }
    .stButton button { border-radius:0; font-weight:700; min-height:2.6rem; }
    .stButton button[kind="primary"] {
      background:var(--ink); border-color:var(--ink); color:#fff;
    }
    .stButton button[kind="primary"]:hover {
      background:var(--signal); border-color:var(--signal); color:#fff;
    }
    .stButton button[kind="secondary"] {
      background:#fff; border-color:var(--ink); color:var(--ink);
    }
    .stButton button[kind="secondary"]:hover {
      border-color:var(--signal); color:var(--signal);
    }
    .site-top { border-bottom:1px solid var(--line); padding:0 0 1rem; }
    .site-top .wordmark { font-size:1.15rem; font-weight:800; letter-spacing:-.07em; }
    .site-top .wordmark span { font-weight:400; }
    .site-top .subtitle { color:#767676; font-size:.72rem; font-weight:700; letter-spacing:.05em; margin-top:.25rem; }
    .hero { background:var(--ink); color:#fff; padding:1.6rem 1.8rem; margin:1rem 0 1.3rem; }
    .hero .eyebrow { color:#c7c7c7; font-size:.72rem; font-weight:700; letter-spacing:.07em; }
    .hero h1 { color:#fff; margin:.35rem 0 .7rem; font-size:clamp(1.75rem,4vw,2.65rem); line-height:1.12; }
    .hero p { color:#e2e2e2; margin:0; line-height:1.65; }
    .verdict { padding:1rem 1.2rem; border-left:5px solid; background:#fff;
               font-size:1.25rem; font-weight:750; margin:.5rem 0 1rem; }
    .complete { color:#176845; border-color:#176845; }
    .incomplete { color:#a52c25; border-color:var(--signal); }
    .pending { color:#806009; border-color:#d9ad34; }
    @media(max-width:700px) {
      [data-testid="stMainBlockContainer"] { padding:1.1rem .9rem 2rem; }
      .hero { padding:1.25rem 1.1rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def switch_page(page: str) -> None:
    st.session_state["active_page"] = page
    st.rerun()


st.session_state.setdefault("active_page", "상품소개서")
header_left, header_right = st.columns([4, 1], vertical_alignment="center")
with header_left:
    st.markdown(
        '<div class="site-top"><div class="wordmark">TM <span>Quality Check</span></div>'
        '<div class="subtitle">INSURANCE TM QUALITY CONTROL</div></div>',
        unsafe_allow_html=True,
    )
with header_right:
    if st.session_state["active_page"] == "상품소개서":
        if st.button("녹취 점검 시작", type="primary", use_container_width=True):
            switch_page("녹취 점검")
    elif st.button("상품소개서 보기", type="primary", use_container_width=True):
        switch_page("상품소개서")

if st.session_state["active_page"] == "상품소개서":
    st.markdown(
        '<div class="hero"><div class="eyebrow">ABOUT · PROCESS · CRITERIA · REPORT</div>'
        '<h1>보험 TM 완전판매 점검 솔루션</h1>'
        '<p>상담 녹취를 읽고, 표준 스크립트와 비교하고, 판정 근거를 보고서로 남깁니다.</p>'
        '</div>',
        unsafe_allow_html=True,
    )
    if BROCHURE_PATH.is_file():
        st.caption("상품소개서 · 아래에서 전체 내용을 확인할 수 있습니다.")
        st.iframe(BROCHURE_PATH, height="content")
    else:
        st.error("상품소개서 파일을 찾을 수 없습니다. 앱과 같은 폴더에 소개서 HTML 파일을 넣어주세요.")
    st.stop()

st.markdown(
    '<div class="hero"><div class="eyebrow">SALES REVIEW</div>'
    '<h1>완전판매 점검</h1>'
    '<p>가입 녹취 MP3를 등록하고 표준 스크립트의 필수 안내 항목을 검토합니다.</p></div>',
    unsafe_allow_html=True,
)


def autofill_from_filename(uploaded_file) -> None:
    if uploaded_file is None:
        return
    audio_bytes = uploaded_file.getvalue()
    upload_signature = hashlib.sha256(audio_bytes).hexdigest()
    if st.session_state.get("processed_upload_signature") == upload_signature:
        return
    st.session_state["processed_upload_signature"] = upload_signature

    match = re.fullmatch(r"(.+)_([0-9]{8})\.mp3", uploaded_file.name, flags=re.IGNORECASE)
    if not match:
        st.session_state["filename_notice"] = (
            "파일명 형식이 올바르지 않습니다. 회원번호_YYYYMMDD.mp3 형식으로 선택해주세요."
        )
        return

    member_id, compact_date = match.groups()
    try:
        parsed_date = date(
            int(compact_date[:4]), int(compact_date[4:6]), int(compact_date[6:8])
        )
    except ValueError:
        st.session_state["filename_notice"] = "파일명에 포함된 판매일자가 올바른 날짜가 아닙니다."
        return

    st.session_state["member_id"] = member_id
    st.session_state["sale_date"] = parsed_date
    st.session_state["filename_notice"] = (
        f"파일명에서 회원번호 {member_id}, 판매일자 {parsed_date.isoformat()}를 자동 입력했습니다."
    )


def show_transcript(slot, transcript: str) -> None:
    with slot.container(border=True):
        st.subheader("녹취 텍스트 변환 결과")
        st.text_area("녹취 텍스트", value=transcript, height=320, label_visibility="collapsed", disabled=True)


with st.container(border=True):
    st.subheader("녹취 조회 및 검토")
    uploaded = st.file_uploader(
        "가입 녹취 MP3",
        type=["mp3"],
        help="파일명 형식: 회원번호_YYYYMMDD.mp3 · 최대 25MB",
        key="audio_file",
    )
    autofill_from_filename(uploaded)

    member_col, date_col = st.columns(2)
    with member_col:
        member_id = st.text_input(
            "회원번호", max_chars=80, placeholder="예: M260001", key="member_id"
        ).strip()
    with date_col:
        sale_date_value = st.date_input("판매일자", key="sale_date")

    notice = st.session_state.pop("filename_notice", None)
    if notice:
        if "올바르지" in notice or "형식" in notice:
            st.warning(notice)
        else:
            st.info(notice)

    if uploaded is not None:
        st.audio(uploaded.getvalue(), format="audio/mpeg")

    st.caption("파일명에서 회원번호와 판매일자를 읽어 입력란을 자동으로 채웁니다.")
    submitted = st.button("MP3 업로드 후 검토", type="primary", use_container_width=False)

transcript_slot = st.empty()
visible_transcript = st.session_state.get("transcript_cache")
if uploaded is not None and visible_transcript:
    current_signature = hashlib.sha256(uploaded.getvalue()).hexdigest()
    if visible_transcript["signature"] == current_signature:
        show_transcript(transcript_slot, visible_transcript["text"])

if submitted:
    if uploaded is None:
        st.error("검토할 MP3 파일을 선택해주세요.")
    elif len(uploaded.getvalue()) > MAX_AUDIO_BYTES:
        st.error("MP3 파일은 25MB 이하만 업로드할 수 있습니다.")
    elif not member_id:
        st.error("회원번호를 입력해주세요.")
    elif not configured_api_key():
        st.error("음성 변환과 완전판매 스크립트 검토를 위해 .env에 OPENAI_API_KEY를 설정해야 합니다.")
    else:
        status = None
        try:
            sale_date = normalize_date(sale_date_value)
            target = find_target(member_id, sale_date)
            validate_upload_filename(uploaded.name, member_id, sale_date, target)
            upload_signature = hashlib.sha256(uploaded.getvalue()).hexdigest()
            with st.status("가입 건 확인 완료. 음성 변환을 준비합니다.", expanded=True) as status:
                cached = st.session_state.get("transcript_cache")
                if cached and cached["signature"] == upload_signature:
                    transcript = cached["text"]
                    st.write("이 파일의 변환된 녹취 텍스트를 다시 사용합니다.")
                else:
                    st.write("녹취 음성을 한국어 텍스트로 변환하고 있습니다…")
                    transcript = transcribe(uploaded.getvalue(), uploaded.name)
                if not transcript.strip():
                    raise ApiError(422, "녹취 텍스트가 비어 있습니다.")
                st.session_state["transcript_cache"] = {
                    "signature": upload_signature,
                    "text": transcript,
                }
                if not visible_transcript or visible_transcript["signature"] != upload_signature:
                    show_transcript(transcript_slot, transcript)
                status.update(label="녹취 텍스트 변환 완료. 완전판매 스크립트를 검토합니다.", state="running")
                st.write("완전판매 스크립트 항목과 녹취 내용을 대조하고 있습니다…")
                review = review_transcript(transcript, read_standard_sections())
                status.update(label="녹취 조회와 검토가 완료되었습니다.", state="complete", expanded=False)

            st.divider()
            st.subheader("검토 결과")
            st.caption(
                f"회원번호 {member_id} · 판매일자 {sale_date} · "
                f"상담원 {target.get('상담원 사번') or '정보 없음'} · "
                f"{target.get('상품명') or '상품 정보 없음'} · 녹취 {uploaded.name}"
            )
            judgment = review["judgment"]
            verdict_style = (
                "complete" if judgment.startswith("완전판매")
                else "incomplete" if "불완전판매" in judgment
                else "pending"
            )
            st.markdown(
                f'<div class="verdict {verdict_style}">{judgment}</div>',
                unsafe_allow_html=True,
            )
            st.markdown("**핵심내용**")
            st.write(review.get("summary", ""))

            st.markdown("**완전판매 스크립트 항목별 검토**")
            for item in review.get("sections", []):
                with st.expander(f"{item['id']}. {item['title']} · {item['status']}"):
                    st.write(f"**결과:** {item['status']}")
                    st.write(f"**녹취 근거:** {item.get('evidence') or '확인 근거 없음'}")
                    missing = item.get("missing_points", [])
                    if missing:
                        st.write("**미흡/누락:** " + " · ".join(missing))

            st.caption(
                "AI 검토는 제공된 녹취와 완전판매 스크립트에 근거한 검토 결과입니다. "
                "근거가 없거나 모호한 항목은 최종 확인이 필요합니다."
            )
        except ApiError as exc:
            if status is not None:
                status.update(label="검토를 완료하지 못했습니다.", state="error", expanded=True)
            st.error(exc.message)
        except Exception as exc:
            if status is not None:
                status.update(label="검토를 완료하지 못했습니다.", state="error", expanded=True)
            st.error(f"처리 중 오류가 발생했습니다 ({type(exc).__name__}). 서버 콘솔 로그를 확인해주세요.")
