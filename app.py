import streamlit as st
import yt_dlp
import re
import json
import io
import os
import tempfile
from google import genai
from google.genai import types
from weasyprint import HTML
from pypdf import PdfReader
from docx import Document
from docx.shared import Pt, RGBColor, Cm

# ---------------------------------------------------------
# 1. 페이지 설정 및 비밀번호 인증 (팀원 공유용)
# ---------------------------------------------------------
st.set_page_config(page_title="주일설교 사운드 녹취·편집기", page_icon="📖", layout="wide")

def check_password():
    if "APP_PASSWORD" not in st.secrets:
        return True
    if st.session_state.get("authenticated", False):
        return True
    st.title("🔒 주일설교 사운드 녹취·편집 웹 프로그램")
    pwd = st.text_input("팀원 전용 접속 비밀번호를 입력하세요:", type="password")
    if st.button("접속하기", type="primary"):
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False

if not check_password():
    st.stop()

st.title("📖 주일설교 사운드 직접 녹취 & A4 2장 편집기")
st.caption("유튜브 주소를 입력하면 서버에서 오디오/설명란을 추출하고, 제미나이가 사운드를 직접 녹취·편집하여 A4 2장(상단 0.5장 + 본문 1.5장)으로 완성합니다.")

# API 키 로드 (Secrets 우선, 없으면 사이드바 입력)
api_key = st.secrets.get("GEMINI_API_KEY", "")
with st.sidebar:
    st.header("⚙️ 설정 및 출력 규격")
    if not api_key:
        api_key = st.text_input("Gemini API Key 입력", type="password")
    else:
        st.success("✅ 서버 API 키가 연결되었습니다.")
    st.markdown("---")
    st.markdown(
        "**📋 출력 문서 규격 (A4 2장 고정)**\n"
        "- **상단 (0.5장 이내):** 설교 제목, 설교자/일자, 성경구절 (**푸른색**)\n"
        "- **본문 (1.5장):** 도입(3~4줄) + 1~3대지 요약(목사님 원문 어조 유지, 시간 표시 제거)\n"
        "- **대지 제목:** **<u>진하게 + 밑줄</u>**\n"
        "- **적용질문:** <span style='color:#7B1FA2;'>자주색</span>\n"
        "- **다운로드:** PDF(A4 2장 자동맞춤), 한글(.hwp), 워드(.docx)",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------
# 2. 유튜브 설명란 및 사운드(오디오) 추출 함수
# ---------------------------------------------------------
def extract_youtube_audio_and_meta(url):
    with tempfile.TemporaryDirectory() as tmpdir:
        out_tmpl = os.path.join(tmpdir, "sermon_audio.%(ext)s")
        ydl_opts = {
            "format": "bestaudio[ext=m4a]/bestaudio/best",
            "outtmpl": out_tmpl,
            "noplaylist": True,
            "quiet": True,
            "postprocessors": [{
                "key": "FFmpegExtractAudio",
                "preferredcodec": "mp3",
                "preferredbitrate": "64k",
            }],
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            yt_title = info.get("title", "")
            yt_desc = info.get("description", "")

        mp3_path = os.path.join(tmpdir, "sermon_audio.mp3")
        if not os.path.exists(mp3_path):
            files = [os.path.join(tmpdir, f) for f in os.listdir(tmpdir) if not f.endswith(".part")]
            if not files:
                raise RuntimeError("유튜브에서 사운드 파일을 추출하지 못했습니다.")
            mp3_path = files[0]

        with open(mp3_path, "rb") as f:
            audio_bytes = f.read()

        return yt_title, yt_desc, audio_bytes

# ---------------------------------------------------------
# 3. Gemini 사운드 직접 녹취 및 A4 1.5장 본문 편집 함수
# ---------------------------------------------------------
def transcribe_and_edit_sermon(api_key, yt_title, yt_desc, audio_bytes, p1, p2, p3):
    client = genai.Client(api_key=api_key)

    uploaded_file = None
    tmp_audio_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp_audio:
            tmp_audio.write(audio_bytes)
            tmp_audio_path = tmp_audio.name

        uploaded_file = client.files.upload(file=tmp_audio_path)

        prompt = (
            "당신은 첨부된 설교 사운드(오디오)를 직접 듣고 정확히 녹취한 뒤, 교인들이 읽고 묵상할 수 있도록 "
            "A4 용지 정확히 2장 분량의 요약·편집본으로 만드는 전문 에디터입니다.\n\n"
            "[참고 메타데이터]\n"
            f"- 유튜브 영상 제목: {yt_title}\n"
            f"- 유튜브 조회수 아랫부분(설명란):\n{yt_desc[:2000]}\n\n"
            "[사용자 입력 대지 및 적용질문]\n"
            f"- 1대지와 적용질문:\n{p1}\n"
            f"- 2대지와 적용질문:\n{p2}\n"
            f"- 3대지와 적용질문:\n{p3}\n\n"
            "[편집 필수 지침 - 반드시 준수]\n"
            "1. 상단 정보 (A4 용지 0.5장 이내로 간결하게 구성):\n"
            "   - 'title': 설교 제목\n"
            "   - 'meta': 날짜 및 설교자 (예: 11/22 주일설교_김진규 목사님)\n"
            "   - 'scripture_ref': 성경구절 장절 (예: 고린도전서 12장 1~11절)\n"
            "   - 'scripture_verses': 해당 성경 본문 전체(개역개정 기준, 각 절 앞에 절 번호 포함하여 줄바꿈 '\\n'으로 구분). "
            "상단이 0.5장을 넘지 않도록 깔끔하게 정리하세요.\n\n"
            "2. 사운드 녹취 및 본문 편집 (도입 + 1대지 + 2대지 + 3대지 합계 A4 용지 1.5장 분량 = 공백 포함 약 2,650~2,900자):\n"
            "   - 사운드 직접 녹취: 첨부된 오디오의 실제 음성을 듣고 내용을 파악하세요.\n"
            "   - 목사님 대본 원문 어조 유지: 제3자 시점('목사님은 ~라고 말씀하셨습니다')으로 요약하지 말고, "
            "목사님이 실제로 말씀하신 구어체 어조, 강조 표현, 성경 본문 풀이, 간증과 예화를 최대한 원문 그대로 살려 자연스럽게 다듬어 연결하세요.\n"
            "   - 시간 표시 금지: 00:00 같은 타임스탬프는 절대 넣지 마세요.\n"
            "   - 'intro': 설교 도입 내용을 목사님 말씀 어조를 살려 3~4줄(공백 포함 약 250~320자)로 작성하세요.\n"
            "   - 'p1_heading', 'p2_heading', 'p3_heading': 입력받은 1, 2, 3대지 제목을 그대로 사용하세요.\n"
            "   - 'p1_body', 'p2_body', 'p3_body': 각 대지와 적용질문 사이에 들어갈 본문입니다. "
            "각 대지당 공백 포함 약 820~900자 분량으로 작성하여, 전체 본문(intro + p1_body + p2_body + p3_body)이 정확히 A4 1.5장을 채우도록 하세요. "
            "문단은 2~3개로 나누어 가독성을 높이세요.\n"
            "   - 'p1_question', 'p2_question', 'p3_question': 입력받은 각 대지의 적용질문(Q. ~)을 그대로 넣으세요.\n\n"
            '반드시 아래 JSON 형식으로만 응답하세요:\n'
            '{"title": "", "meta": "", "scripture_ref": "", "scripture_verses": "", '
            '"intro": "", "p1_heading": "", "p1_body": "", "p1_question": "", '
            '"p2_heading": "", "p2_body": "", "p2_question": "", '
            '"p3_heading": "", "p3_body": "", "p3_question": ""}'
        )

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[uploaded_file, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.2,
            ),
        )
        return json.loads(response.text)
    finally:
        if tmp_audio_path and os.path.exists(tmp_audio_path):
            try:
                os.remove(tmp_audio_path)
            except Exception:
                pass
        if uploaded_file is not None:
            try:
                client.files.delete(name=uploaded_file.name)
            except Exception:
                pass

# ---------------------------------------------------------
# 4. 문서 생성 함수 (PDF A4 2장 자동 보정 / HWP / DOCX)
# ---------------------------------------------------------
def render_pdf_html(data, font_size=9.6, line_height=1.52, margin_v=14):
    verses_html = "<br>".join(data["scripture_verses"].split("\n"))
    points_html = ""
    for idx in [1, 2, 3]:
        h = data[f"p{idx}_heading"]
        b = data[f"p{idx}_body"]
        q = data[f"p{idx}_question"]
        paragraphs = "".join([
            f"<p class='body-text'>{para.strip()}</p>"
            for para in b.split("\n\n") if para.strip()
        ])
        if not paragraphs:
            paragraphs = f"<p class='body-text'>{b.strip()}</p>"
        points_html += (
            f'<div class="point-block">'
            f'<div class="point-heading">{h}</div>'
            f'{paragraphs}'
            f'<div class="point-question">{q}</div>'
            f'</div>'
        )
    return (
        f"<!DOCTYPE html><html lang='ko'><head><meta charset='UTF-8'>"
        f"<style>@page {{ size: A4; margin: {margin_v}mm 16mm {margin_v}mm 16mm; }}"
        f"body {{ font-family: 'NanumGothic', 'Malgun Gothic', sans-serif; font-size: {font_size}pt; line-height: {line_height}; color: #111111; margin: 0; padding: 0; word-break: keep-all; }}"
        f".header-title {{ font-size: {font_size + 3.4}pt; font-weight: bold; margin-bottom: 2px; }}"
        f".header-meta {{ font-size: {font_size + 0.4}pt; font-weight: bold; margin-bottom: 6px; }}"
        f".scripture-box {{ color: #1B4F9C; font-size: {font_size - 0.4}pt; line-height: {line_height - 0.08}; margin-bottom: 9px; }}"
        f".scripture-ref {{ font-weight: bold; margin-bottom: 2px; }}"
        f".intro-box {{ margin-bottom: 9px; text-align: justify; }}"
        f".point-block {{ margin-bottom: 8px; }}"
        f".point-heading {{ font-size: {font_size + 0.6}pt; font-weight: bold; text-decoration: underline; text-underline-offset: 3px; margin-top: 5px; margin-bottom: 4px; }}"
        f".body-text {{ margin: 0 0 4px 0; text-align: justify; }}"
        f".point-question {{ color: #7B1FA2; font-weight: 600; margin-top: 3px; margin-bottom: 5px; }}"
        f"</style></head><body>"
        f"<div class='header-title'>{data['title']}</div>"
        f"<div class='header-meta'>{data['meta']}</div>"
        f"<div class='scripture-box'><div class='scripture-ref'>{data['scripture_ref']}</div><div>{verses_html}</div></div>"
        f"<div class='intro-box'>{data['intro']}</div>"
        f"{points_html}</body></html>"
    )

def create_exact_2page_pdf(data):
    font_size = 9.6
    line_height = 1.52
    margin_v = 14
    pdf_bytes = HTML(string=render_pdf_html(data, font_size, line_height, margin_v)).write_pdf()

    for _ in range(8):
        page_count = len(PdfReader(io.BytesIO(pdf_bytes)).pages)
        if page_count == 2:
            break
        elif page_count > 2:
            font_size -= 0.25
            line_height -= 0.04
            margin_v = max(10, margin_v - 1)
        else:
            font_size += 0.3
            line_height += 0.05
        pdf_bytes = HTML(string=render_pdf_html(data, font_size, line_height, margin_v)).write_pdf()
    return pdf_bytes

def create_hwp_file(data):
    verses_html = "<br>".join(data["scripture_verses"].split("\n"))
    points_html = ""
    for idx in [1, 2, 3]:
        h = data[f"p{idx}_heading"]
        b = data[f"p{idx}_body"].replace("\n", "<br>")
        q = data[f"p{idx}_question"]
        points_html += (
            f"<p style='margin:7pt 0 3pt 0; font-size:10.2pt; font-weight:bold; text-decoration:underline;'>{h}</p>"
            f"<p style='margin:0 0 4pt 0; line-height:150%; text-align:justify;'>{b}</p>"
            f"<p style='margin:3pt 0 7pt 0; color:#7B1FA2; line-height:148%;'>{q}</p>"
        )
    hwp_html = (
        "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'>"
        "<head><meta charset='utf-8'><style>@page { size: A4; margin: 14mm 16mm 14mm 16mm; }"
        "body { font-family: '맑은 고딕', '함초롬바탕', sans-serif; font-size: 9.5pt; color: #111111; }</style></head><body>"
        f"<p style='font-size:13pt; font-weight:bold; margin:0 0 2pt 0;'>{data['title']}</p>"
        f"<p style='font-size:10pt; font-weight:bold; margin:0 0 6pt 0;'>{data['meta']}</p>"
        f"<p style='color:#1B4F9C; font-weight:bold; margin:0 0 2pt 0;'>{data['scripture_ref']}</p>"
        f"<p style='color:#1B4F9C; font-size:9pt; line-height:142%; margin:0 0 9pt 0;'>{verses_html}</p>"
        f"<p style='line-height:150%; text-align:justify; margin:0 0 9pt 0;'>{data['intro']}</p>"
        f"{points_html}</body></html>"
    )
    return hwp_html.encode("utf-8")

def create_docx_file(data):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Cm(1.4)
        section.bottom_margin = Cm(1.4)
        section.left_margin = Cm(1.6)
        section.right_margin = Cm(1.6)

    style = doc.styles["Normal"]
    style.font.name = "맑은 고딕"
    style.font.size = Pt(9.5)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run(data["title"])
    r_title.bold = True
    r_title.font.size = Pt(13)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(6)
    r_meta = p_meta.add_run(data["meta"])
    r_meta.bold = True
    r_meta.font.size = Pt(10)

    blue_color = RGBColor(27, 79, 156)
    p_ref = doc.add_paragraph()
    p_ref.paragraph_format.space_after = Pt(2)
    r_ref = p_ref.add_run(data["scripture_ref"])
    r_ref.bold = True
    r_ref.font.color.rgb = blue_color

    for v in data["scripture_verses"].split("\n"):
        if v.strip():
            p_v = doc.add_paragraph()
            p_v.paragraph_format.space_after = Pt(1)
            p_v.paragraph_format.line_spacing = 1.25
            r_v = p_v.add_run(v.strip())
            r_v.font.color.rgb = blue_color
            r_v.font.size = Pt(9)

    p_intro = doc.add_paragraph(data["intro"])
    p_intro.paragraph_format.space_before = Pt(6)
    p_intro.paragraph_format.space_after = Pt(7)
    p_intro.paragraph_format.line_spacing = 1.35

    purple_color = RGBColor(123, 31, 162)
    for idx in [1, 2, 3]:
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(6)
        p_head.paragraph_format.space_after = Pt(3)
        r_head = p_head.add_run(data[f"p{idx}_heading"])
        r_head.bold = True
        r_head.underline = True
        r_head.font.size = Pt(10.2)

        for para in data[f"p{idx}_body"].split("\n\n"):
            if para.strip():
                p_body = doc.add_paragraph(para.strip())
                p_body.paragraph_format.space_after = Pt(4)
                p_body.paragraph_format.line_spacing = 1.35

        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_after = Pt(6)
        r_q = p_q.add_run(data[f"p{idx}_question"])
        r_q.font.color.rgb = purple_color

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ---------------------------------------------------------
# 5. 메인 화면 입력 폼 및 결과 출력
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("1️⃣ 설교 정보 입력")
    youtube_url = st.text_input("1) 설교 유튜브 주소", placeholder="https://www.youtube.com/watch?v=...")
    point1 = st.text_area("2-1) 1대지와 적용질문", height=110, placeholder="1. 신령한 것을 알아야 합니다.\nQ. 여러분은 하나님이 은혜로 주신 것을 가지고 나를 높이고 있지는 않습니까?")
    point2 = st.text_area("2-2) 2대지와 적용질문", height=110, placeholder="2. 섬기는 자가 되어야 합니다.\nQ. 여러분은 맡겨진 자리에서 섬기고 있습니까?")
    point3 = st.text_area("2-3) 3대지와 적용질문", height=110, placeholder="3. 서로의 유익을 구해야 합니다.\nQ. 여러분은 지금 누구를 믿음이 부족하다고 판단하고 있습니까?")

with col2:
    st.subheader("2️⃣ 사운드 녹취·편집 및 다운로드")
    if st.button("🎧 유튜브 사운드 녹취 & A4 2장 편집하기", type="primary", use_container_width=True):
        if not api_key:
            st.error("Gemini API Key를 설정해주세요.")
        elif not youtube_url or not point1 or not point2 or not point3:
            st.warning("유튜브 주소와 1~3대지 및 적용질문을 모두 입력해주세요.")
        else:
            try:
                with st.spinner("1/3 유튜브 영상에서 사운드와 설명란 정보를 추출하는 중..."):
                    yt_title, yt_desc, audio_bytes = extract_youtube_audio_and_meta(youtube_url)
                with st.spinner("2/3 제미나이가 사운드를 직접 녹취하고 목사님 대본 어조 그대로 A4 1.5장 분량으로 편집하는 중..."):
                    summary_data = transcribe_and_edit_sermon(
                        api_key, yt_title, yt_desc, audio_bytes, point1, point2, point3
                    )
                with st.spinner("3/3 PDF, 한글(.hwp), 워드(.docx) 파일을 생성하는 중..."):
                    pdf_bytes = create_exact_2page_pdf(summary_data)
                    hwp_bytes = create_hwp_file(summary_data)
                    docx_bytes = create_docx_file(summary_data)

                safe_title = re.sub(r'[\\/*?:"<>|]', '', summary_data["title"]).strip() or "주일설교"
                st.success("✅ 사운드 녹취 및 A4 2장 편집이 완료되었습니다!")

                b1, b2, b3 = st.columns(3)
                with b1:
                    st.download_button("📄 PDF (A4 2장)", data=pdf_bytes, file_name=f"{safe_title}_요약.pdf", mime="application/pdf", use_container_width=True)
                with b2:
                    st.download_button("🇰🇷 한글 (.hwp)", data=hwp_bytes, file_name=f"{safe_title}_요약.hwp", mime="application/x-hwp", use_container_width=True)
                with b3:
                    st.download_button("📝 워드 (.docx)", data=docx_bytes, file_name=f"{safe_title}_요약.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

                total_body_chars = len(
                    summary_data["intro"] + summary_data["p1_body"] + summary_data["p2_body"] + summary_data["p3_body"]
                )
                st.caption(f"본문 분량: {total_body_chars:,}자 (도입 + 1~3대지 합계 A4 1.5장 규격)")

                st.markdown("---")
                st.markdown(f"### {summary_data['title']}")
                st.markdown(f"**{summary_data['meta']}**")
                st.markdown(f"<span style='color:#1B4F9C; font-weight:bold;'>{summary_data['scripture_ref']}</span>", unsafe_allow_html=True)
                st.markdown(f"<div style='color:#1B4F9C; font-size:0.9em; margin-bottom:10px;'>{'<br>'.join(summary_data['scripture_verses'].split('\\n'))}</div>", unsafe_allow_html=True)
                st.write(summary_data["intro"])
                for idx in [1, 2, 3]:
                    st.markdown(f"**<u>{summary_data[f'p{idx}_heading']}</u>**", unsafe_allow_html=True)
                    st.write(summary_data[f"p{idx}_body"])
                    st.markdown(f"<span style='color:#7B1FA2;'>{summary_data[f'p{idx}_question']}</span>", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
