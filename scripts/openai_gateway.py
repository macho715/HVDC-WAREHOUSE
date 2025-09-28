"""OpenAI 연동 FastAPI 백엔드. | FastAPI backend wired with OpenAI."""

from __future__ import annotations

import base64
import io
import json
import logging
import os
from typing import Iterable, List, Literal, Sequence

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from openai import OpenAI
from openai.types import Response
from pydantic import BaseModel, Field
from PyPDF2 import PdfReader


LOGGER = logging.getLogger(__name__)


DEFAULT_MODEL = "gpt-4o-mini"


class ChatMessage(BaseModel):
    """사용자/AI 메시지 구조. | Chat message schema."""

    role: Literal["user", "assistant", "system"]
    content: str = Field(..., max_length=6000)


class BriefingRequest(BaseModel):
    """일일 브리핑 요청 본문. | Daily briefing payload."""

    current_time: str
    vessel_name: str
    vessel_status: str
    current_voyage: str | None = None
    schedule: List[dict] = Field(default_factory=list)
    weather_windows: List[dict] = Field(default_factory=list)


class BriefingResponse(BaseModel):
    """일일 브리핑 응답. | Daily briefing response."""

    briefing: str


class AssistantResponse(BaseModel):
    """AI 어시스턴트 응답. | AI assistant response."""

    answer: str


def _require_client() -> OpenAI:
    """OpenAI 클라이언트를 생성. | Build an OpenAI client."""

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY is not configured")
    return OpenAI(api_key=api_key)


def _pdf_to_text(payload: bytes) -> str:
    """PDF를 텍스트로 추출. | Extract text from PDF."""

    reader = PdfReader(io.BytesIO(payload))
    text_chunks: List[str] = []
    for page in reader.pages:
        snippet = page.extract_text() or ""
        text_chunks.append(snippet)
    return "\n".join(text_chunks)


def _build_history(messages: Sequence[ChatMessage]) -> List[dict]:
    """응답 API용 메시지 배열 구성. | Build response API messages."""

    history: List[dict] = []
    for item in messages:
        if item.role == "assistant":
            history.append(
                {
                    "role": "assistant",
                    "content": [{"type": "output_text", "text": item.content}],
                }
            )
        else:
            history.append(
                {
                    "role": item.role,
                    "content": [{"type": "input_text", "text": item.content}],
                }
            )
    return history


def _image_to_base64(file: UploadFile, payload: bytes) -> dict:
    """이미지를 base64로 변환. | Convert image to base64."""

    mime = file.content_type or "image/png"
    encoded = base64.b64encode(payload).decode("utf-8")
    data_url = f"data:{mime};base64,{encoded}"
    return {"type": "input_image", "image_url": data_url}


def _build_user_content(prompt: str, files: Iterable[UploadFile], raw_payloads: List[bytes]) -> List[dict]:
    """사용자 메시지 콘텐츠 구성. | Compose user content payload."""

    content: List[dict] = [{"type": "input_text", "text": prompt}]
    for idx, file in enumerate(files):
        data = raw_payloads[idx]
        filename = file.filename or ""
        display_name = file.filename or "unnamed attachment"
        lower_name = filename.lower()

        if file.content_type and file.content_type.startswith("image/"):
            content.append(_image_to_base64(file, data))
        elif (file.content_type == "application/pdf") or lower_name.endswith(".pdf"):
            pdf_text = _pdf_to_text(data)[:8000]
            descriptor = f"\n[첨부 PDF: {display_name}]\n"
            content.append({"type": "input_text", "text": descriptor + pdf_text})
        else:
            try:
                decoded = data.decode("utf-8")
            except UnicodeDecodeError:
                decoded = base64.b64encode(data).decode("utf-8")
                decoded = f"[base64-encoded attachment]\n{decoded[:6000]}"
            content.append({"type": "input_text", "text": f"\n[첨부 파일: {display_name}]\n{decoded[:8000]}"})
    return content


def _call_openai(messages: List[dict], *, model: str) -> Response:
    """OpenAI Responses API 호출. | Invoke OpenAI Responses API."""

    client = _require_client()
    return client.responses.create(model=model, input=messages)


app = FastAPI(title="HVDC Logistics AI Gateway", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/health")
async def healthcheck() -> dict:
    """헬스체크. | Service health check."""

    return {"status": "ok"}


@app.post("/api/assistant", response_model=AssistantResponse)
async def run_assistant(
    prompt: str = Form(..., max_length=4000),
    history: str = Form("[]"),
    files: List[UploadFile] | None = File(default=None),
    model: str = Form(DEFAULT_MODEL),
) -> AssistantResponse:
    """어시스턴트 호출. | Execute assistant call."""

    try:
        raw_history = json.loads(history)
        history_messages = [ChatMessage.model_validate(item) for item in raw_history]
    except (json.JSONDecodeError, TypeError, ValueError) as exc:  # pragma: no cover - validation
        raise HTTPException(status_code=400, detail=f"Invalid history payload: {exc}") from exc

    attachments = list(files or [])
    payloads: List[bytes] = []
    for file in attachments:
        payloads.append(await file.read())

    messages = _build_history(history_messages)
    messages.append({"role": "user", "content": _build_user_content(prompt, attachments, payloads)})

    try:
        response = _call_openai(messages, model=model)
    except Exception as exc:  # pragma: no cover - network failure
        LOGGER.exception("OpenAI call failed")
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return AssistantResponse(answer=response.output_text)


@app.post("/api/briefing", response_model=BriefingResponse)
async def generate_briefing(payload: BriefingRequest) -> BriefingResponse:
    """일일 브리핑 생성. | Create a daily briefing."""

    schedule_summary = json.dumps(payload.schedule, ensure_ascii=False, indent=2)
    weather_summary = json.dumps(payload.weather_windows, ensure_ascii=False, indent=2)
    prompt = (
        "당신은 해상 물류 관제 전문가입니다. 아래 데이터를 참고하여 200자 내외의 한국어 일일 브리핑을 작성하세요."
        "\n- 현재 시각: {time}\n- 선박명: {vessel}\n- 현재 항차: {voyage}\n- 선박 상태: {status}\n"
        "- 전체 일정: {schedule}\n- 기상 윈도우: {weather}\n"
        "브리핑은 핵심 일정, 위험, 권고사항을 bullet로 정리하세요."
    ).format(
        time=payload.current_time,
        vessel=payload.vessel_name,
        voyage=payload.current_voyage or "N/A",
        status=payload.vessel_status,
        schedule=schedule_summary,
        weather=weather_summary,
    )

    try:
        response = _call_openai(
            [
                {
                    "role": "system",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "Respond in Korean and include concise actionable items.",
                        }
                    ],
                },
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": prompt}],
                },
            ],
            model=DEFAULT_MODEL,
        )
    except Exception as exc:  # pragma: no cover - network failure
        LOGGER.exception("OpenAI call failed")
        raise HTTPException(status_code=502, detail=str(exc)) from exc

    return BriefingResponse(briefing=response.output_text)

