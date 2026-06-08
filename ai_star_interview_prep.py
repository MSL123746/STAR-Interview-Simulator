import io
import os

import requests
import streamlit as st
from docx import Document
from docx.shared import Pt


st.set_page_config(page_title="AI-STAR Interview Prep", layout="wide")

st.markdown(
    """
    <style>
    /* Global bold text styling */
    .stApp,
    .stApp p,
    .stApp span,
    .stApp label,
    .stApp div,
    .stApp li,
    .stApp h1,
    .stApp h2,
    .stApp h3,
    .stApp h4,
    .stApp h5,
    .stApp h6 {
        font-weight: 700 !important;
    }

    /* Keep main title/subheaders as-is, but make all other UI text 16px */
    .stApp p,
    .stApp span,
    .stApp label,
    .stApp li,
    .stApp small,
    .stButton > button,
    .stDownloadButton > button,
    .stCheckbox label {
        font-size: 16px !important;
    }

    div[data-testid="stTextInput"] input,
    div[data-testid="stTextArea"] textarea {
        font-size: 16px !important;
    }

    div[data-testid="stTextInput"] input::placeholder,
    div[data-testid="stTextArea"] textarea::placeholder {
        font-size: 16px !important;
    }

    /* Professional neutral styling for top behavioral question input */
    div[data-testid="stTextInput"] input {
        width: 100% !important;
        background: #f8fafc !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }

    div[data-testid="stTextInput"] input::placeholder {
        color: #64748b !important;
    }

    div[data-testid="stTextInput"] input:focus {
        border: 1px solid #2563eb !important;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2) !important;
    }

    /* Professional neutral styling for STAR narrative text areas */
    div[data-testid="stTextArea"] textarea {
        width: 100% !important;
        background: #f8fafc !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }

    div[data-testid="stTextArea"] textarea::placeholder {
        color: #64748b !important;
    }

    div[data-testid="stTextArea"] textarea:focus {
        border: 1px solid #2563eb !important;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.2) !important;
    }

    div[data-testid="stTextArea"] label {
        color: #111827 !important;
        font-weight: 600 !important;
    }

    /* Green full-width download button */
    div.stDownloadButton > button {
        width: 100% !important;
        background: #16a34a !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        padding: 0.65rem 1rem !important;
    }

    div.stDownloadButton > button:hover {
        background: #15803d !important;
    }

    .answer-shell {
        min-height: 360px;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        background: #ffffff;
        padding: 1rem;
    }

    .brand-row {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 0.5rem;
    }

    .star-logo {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: #fbbf24;
        border: 1px solid #f59e0b;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 24px;
        font-weight: 700;
        line-height: 1;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25);
    }

    .brand-title {
        margin: 0;
        color: #0f172a;
        font-size: 2rem;
        line-height: 1.15;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


HF_MODEL = "microsoft/Phi-3-mini-4k-instruct"


def get_hf_token() -> str:
    token = st.secrets.get("HF_API_TOKEN", "") if hasattr(st, "secrets") else ""
    if not token:
        token = os.getenv("HF_API_TOKEN", "")
    if not token:
        token = os.getenv("HUGGINGFACEHUB_API_TOKEN", "")
    return token


def build_prompt(question: str, situation: str, task: str, action: str, result: str, add_followups: bool) -> str:
    followup_instruction = (
        "Then add exactly 2 tailored follow-up interview questions after the final answer under a header 'Follow-up Questions'."
        if add_followups
        else "Do not add follow-up questions."
    )

    return f"""You are an expert interview coach.
Create a polished, professional STAR interview response in first person.

Behavioral Question:
{question}

Candidate STAR Notes:
Situation: {situation}
Task: {task}
Action: {action}
Result: {result}

Requirements:
1) Produce a cohesive final answer that sounds natural, specific, and impactful.
2) Keep the response concise but substantive.
3) Use a clear structure with these bold section headers:
- Situation
- Task
- Action
- Result
4) {followup_instruction}

Return only the final response text.
"""


def call_hf_inference(prompt: str, model: str, token: str) -> str:
    url = f"https://api-inference.huggingface.co/models/{model}"
    headers = {"Authorization": f"Bearer {token}"}
    payload = {
        "inputs": prompt,
        "parameters": {
            "max_new_tokens": 450,
            "temperature": 0.4,
            "top_p": 0.9,
            "return_full_text": False,
        },
        "options": {"wait_for_model": True},
    }

    response = requests.post(url, headers=headers, json=payload, timeout=90)
    response.raise_for_status()
    data = response.json()

    if isinstance(data, list) and data and "generated_text" in data[0]:
        return data[0]["generated_text"].strip()
    if isinstance(data, dict) and "generated_text" in data:
        return data["generated_text"].strip()

    raise ValueError("Unexpected response format from Hugging Face Inference API.")


def _set_times_new_roman(paragraph, size=12, bold=False):
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(size)
        run.bold = bold


def build_docx_bytes(question: str, situation: str, task: str, action: str, result: str, final_answer: str) -> bytes:
    doc = Document()

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)

    title = doc.add_paragraph("AI-STAR Interview Prep Response")
    _set_times_new_roman(title, size=14, bold=True)
    title.paragraph_format.space_after = Pt(10)

    q_header = doc.add_paragraph("Behavioral Question")
    _set_times_new_roman(q_header, size=12, bold=True)
    q_header.paragraph_format.space_after = Pt(2)

    q_body = doc.add_paragraph(question)
    _set_times_new_roman(q_body, size=12, bold=False)
    q_body.paragraph_format.space_after = Pt(10)

    for header, body in [
        ("Situation", situation),
        ("Task", task),
        ("Action", action),
        ("Result", result),
    ]:
        h = doc.add_paragraph(header)
        _set_times_new_roman(h, size=12, bold=True)
        h.paragraph_format.space_after = Pt(2)

        p = doc.add_paragraph(body)
        _set_times_new_roman(p, size=12, bold=False)
        p.paragraph_format.space_after = Pt(8)

    final_header = doc.add_paragraph("Final Answer")
    _set_times_new_roman(final_header, size=12, bold=True)
    final_header.paragraph_format.space_after = Pt(2)

    final_body = doc.add_paragraph(final_answer)
    _set_times_new_roman(final_body, size=12, bold=False)
    final_body.paragraph_format.space_after = Pt(8)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def reset_all_fields():
    st.session_state.behavioral_question = ""
    st.session_state.situation_text = ""
    st.session_state.task_text = ""
    st.session_state.action_text = ""
    st.session_state.result_text = ""
    st.session_state.ai_followups = False
    st.session_state.final_answer = ""


if "final_answer" not in st.session_state:
    st.session_state.final_answer = ""
if "behavioral_question" not in st.session_state:
    st.session_state.behavioral_question = ""
if "situation_text" not in st.session_state:
    st.session_state.situation_text = ""
if "task_text" not in st.session_state:
    st.session_state.task_text = ""
if "action_text" not in st.session_state:
    st.session_state.action_text = ""
if "result_text" not in st.session_state:
    st.session_state.result_text = ""
if "ai_followups" not in st.session_state:
    st.session_state.ai_followups = False

st.markdown(
    """
    <div class="brand-row">
        <div class="star-logo" aria-label="Star logo">&#9733;</div>
        <h1 class="brand-title">AI-STAR Interview Behavioral Question &amp; Narrative Builder</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

question_left, question_right = st.columns([1, 1])
with question_left:
    behavioral_question = st.text_input(
        "ENTER BEHAVIORAL QUESTION",
        key="behavioral_question",
        placeholder="Enter details...",
    )
with question_right:
    st.write("")

left_col, right_col = st.columns([1, 1])

with left_col:
    st.subheader("CRAFT YOUR STAR NARRATIVE")

    situation_text = st.text_area(
        "(S) SITUATION (Background & Context)",
        key="situation_text",
        placeholder="Enter situation details...",
        height=80,
    )
    task_text = st.text_area(
        "(T) TASK (The Responsibility & Goal)",
        key="task_text",
        placeholder="Enter task details...",
        height=80,
    )
    action_text = st.text_area(
        "(A) ACTION (Detailed Steps taken)",
        key="action_text",
        placeholder="Enter action details...",
        height=80,
    )
    result_text = st.text_area(
        "(R) RESULT (The Outcome & Impact)",
        key="result_text",
        placeholder="Enter result details...",
        height=80,
    )

    ai_followups = st.checkbox(
        "AI ENHANCEMENT: Get 2 customized AI follow-up questions based on my STAR narrative.",
        key="ai_followups",
    )

    st.button("Reset Fields", on_click=reset_all_fields)

    all_star_complete = all(
        [
            situation_text.strip(),
            task_text.strip(),
            action_text.strip(),
            result_text.strip(),
        ]
    )
    question_complete = bool(behavioral_question.strip())

    if not all_star_complete or not question_complete:
        st.warning("Complete all 4 STAR fields and enter the behavioral question to enable AI generation.")

    generate_clicked = st.button(
        "Generate Final Answer with AI",
        disabled=not (all_star_complete and question_complete),
    )

    if generate_clicked:
        if not all_star_complete:
            st.warning("Please fill in all STAR fields before generating.")
            st.stop()

        if not behavioral_question.strip():
            st.warning("Please enter the behavioral question before generating.")
            st.stop()

        hf_token = get_hf_token()
        if not hf_token:
            st.error("Missing Hugging Face API token. Add HF_API_TOKEN to Streamlit secrets or environment variables.")
            st.stop()

        prompt = build_prompt(
            question=behavioral_question,
            situation=situation_text,
            task=task_text,
            action=action_text,
            result=result_text,
            add_followups=ai_followups,
        )

        with st.spinner("Generating your final STAR answer..."):
            try:
                generated_text = call_hf_inference(prompt, HF_MODEL, hf_token)
                st.session_state.final_answer = generated_text
            except Exception as gen_error:
                st.error(f"Generation failed: {gen_error}")

with right_col:
    st.subheader("YOUR COMPLETED STAR NARRATIVE")

    st.markdown('<div class="answer-shell">', unsafe_allow_html=True)
    if st.session_state.final_answer:
        st.markdown(st.session_state.final_answer)
    else:
        st.write("")
    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.final_answer:
        docx_bytes = build_docx_bytes(
            question=behavioral_question,
            situation=situation_text,
            task=task_text,
            action=action_text,
            result=result_text,
            final_answer=st.session_state.final_answer,
        )

        st.download_button(
            "Download Answer as MS Word Doc (.docx)",
            data=docx_bytes,
            file_name="ai_star_answer.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        st.download_button(
            "Download Answer as MS Word Doc (.docx)",
            data=b"",
            file_name="ai_star_answer.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=True,
            use_container_width=True,
        )
