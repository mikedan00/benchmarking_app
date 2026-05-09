#!/usr/bin/env python3
"""
AI 벤치마킹 보고서 자동 작성 시스템
HuggingFace LLM (google/gemma-4-26B-A4B-it) 기반 경쟁사 분석 통합 플랫폼
"""

import streamlit as st
import os
import json
import re
import time
import requests
import pandas as pd
from pathlib import Path
from io import BytesIO
from datetime import datetime
import PyPDF2
from docx import Document as DocxDocument
from bs4 import BeautifulSoup
from huggingface_hub import InferenceClient
from duckduckgo_search import DDGS

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PAGE CONFIGURATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.set_page_config(
    page_title="AI 벤치마킹 보고서 시스템",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "AI 기반 벤치마킹 보고서 자동 생성 시스템 | Powered by HuggingFace LLM"
    }
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CUSTOM CSS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("""
<style>
    .stApp { background: #0f172a; }
    .main .block-container { padding-top: 1rem; max-width: 1400px; }

    .hero-header {
        background: linear-gradient(135deg, #1e40af 0%, #7c3aed 50%, #0891b2 100%);
        border-radius: 16px; padding: 36px 32px; margin-bottom: 24px;
        text-align: center; box-shadow: 0 20px 60px rgba(124,58,237,0.4);
    }
    .hero-title { font-size: 2.4rem; font-weight: 800; color: #ffffff; margin: 0 0 8px 0; }
    .hero-subtitle { font-size: 1.05rem; color: rgba(255,255,255,0.85); margin: 0; }
    .hero-badge {
        display: inline-block; background: rgba(255,255,255,0.15);
        border: 1px solid rgba(255,255,255,0.3); border-radius: 999px;
        padding: 4px 14px; font-size: 0.78rem; color: rgba(255,255,255,0.9); margin-top: 14px;
    }
    .step-card {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 16px;
    }
    .step-card h4 { color: #e2e8f0; margin: 0 0 8px 0; font-size: 1rem; }
    .step-card p  { color: #94a3b8; margin: 0; font-size: 0.88rem; }

    .badge-ok   { background:#064e3b;color:#6ee7b7;border:1px solid #065f46;border-radius:999px;padding:3px 10px;font-size:0.75rem; }
    .badge-warn { background:#422006;color:#fcd34d;border:1px solid #92400e;border-radius:999px;padding:3px 10px;font-size:0.75rem; }

    .voc-card { background:#1e293b;border:1px solid #334155;border-left:4px solid #7c3aed;border-radius:8px;padding:14px;margin-bottom:10px; }
    .voc-card .voc-title { color:#c4b5fd;font-weight:600;font-size:0.9rem;margin-bottom:6px; }
    .voc-card .voc-body  { color:#94a3b8;font-size:0.82rem;line-height:1.5; }

    section[data-testid="stSidebar"] { background: linear-gradient(180deg,#0f172a 0%,#1e1b4b 100%) !important; border-right:1px solid #334155; }

    .stTabs [data-baseweb="tab-list"] { background:#1e293b;border-radius:10px;padding:4px;gap:4px; }
    .stTabs [data-baseweb="tab"]      { border-radius:8px;color:#64748b;font-weight:500; }
    .stTabs [aria-selected="true"]    { background:linear-gradient(135deg,#1e40af,#7c3aed) !important;color:white !important; }

    .stButton>button { background:linear-gradient(135deg,#1e40af,#7c3aed);color:white;border:none;border-radius:8px;font-weight:600; }
    .stProgress>div>div { background:linear-gradient(90deg,#1e40af,#7c3aed);border-radius:999px; }
</style>
""", unsafe_allow_html=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SESSION STATE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DEFAULTS = {
    # st.secrets (Streamlit Cloud) → os.getenv (.env 로컬) → 빈 문자열 순으로 탐색
    "hf_token": (
        st.secrets.get("HF_TOKEN", "")
        or os.getenv("HF_TOKEN", "")
    ),
    "model_id": "google/gemma-4-26B-A4B-it",
    "benchmark_guide": None,
    "guide_extracted": False,
    "voc_my": {}, "voc_comp": {},
    "voc_analyzed_my": {}, "voc_analyzed_comp": {},
    "spec_comparison": {},
    "spec_my_raw": "", "spec_comp_raw": "",
    "report_md": "",
    "my_company": "", "my_product": "",
    "competitor": "", "comp_product": "",
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# LLM CLIENT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def get_client():
    if not st.session_state.hf_token:
        return None
    return InferenceClient(
        model=st.session_state.model_id,
        token=st.session_state.hf_token,
    )


def call_llm(prompt: str, max_tokens: int = 2048, system: str = "") -> str:
    """
    HuggingFace Inference API – chat.completions 전용 호출.
    google/gemma-4-26B-A4B-it 및 Zyphra/ZAYA1-8B 는 'conversational' 태스크만
    지원하므로 text_generation fallback을 제거하고 chat completions만 사용.
    """
    client = get_client()
    if client is None:
        return "⚠️ HuggingFace API 토큰을 먼저 사이드바에 입력해주세요."

    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})

    try:
        resp = client.chat.completions.create(
            messages=messages,
            max_tokens=max_tokens,
            temperature=0.25,
        )
        result = resp.choices[0].message.content
        return result.strip() if result else ""
    except Exception as e:
        err_msg = str(e)
        # 사용자에게 친절한 오류 메시지 반환
        if "not supported for task" in err_msg:
            return (
                f"❌ 모델 태스크 오류\n\n"
                f"선택한 모델이 Chat Completions(conversational) 태스크를 "
                f"지원하지 않습니다.\n\n"
                f"원본 오류: {err_msg}\n\n"
                f"💡 해결 방법: 사이드바에서 지원되는 모델로 변경하거나, "
                f"HuggingFace 모델 페이지에서 'Inference API' 지원 여부를 확인하세요."
            )
        if "401" in err_msg or "Unauthorized" in err_msg:
            return "❌ 인증 오류: HuggingFace API 토큰이 올바르지 않습니다. 사이드바에서 토큰을 확인해주세요."
        if "429" in err_msg or "rate limit" in err_msg.lower():
            return "❌ API 호출 한도 초과: 잠시 후 다시 시도해주세요. (HuggingFace 무료 티어 제한)"
        if "timeout" in err_msg.lower():
            return "❌ 응답 시간 초과: 모델 서버가 바쁩니다. 잠시 후 재시도해주세요."
        return f"❌ LLM API 오류: {err_msg}"


def safe_json(text: str):
    try:
        cleaned = re.sub(r"```(?:json)?", "", text).strip().rstrip("`").strip()
        start = min(
            (cleaned.find("{") if "{" in cleaned else len(cleaned)),
            (cleaned.find("[") if "[" in cleaned else len(cleaned)),
        )
        if start < len(cleaned):
            return json.loads(cleaned[start:])
    except Exception:
        pass
    return None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FILE TEXT EXTRACTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def extract_pdf(data: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(BytesIO(data))
        return "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception as e:
        return f"[PDF 추출 오류: {e}]"

def extract_docx(data: bytes) -> str:
    try:
        doc = DocxDocument(BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception as e:
        return f"[DOCX 추출 오류: {e}]"

def extract_file(uploaded) -> str:
    raw = uploaded.read()
    name = uploaded.name.lower()
    if name.endswith(".pdf"):
        return extract_pdf(raw)
    elif name.endswith((".docx", ".doc")):
        return extract_docx(raw)
    else:
        return raw.decode("utf-8", errors="ignore")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DEFAULT GUIDE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DEFAULT_GUIDE = {
    "title": "벤치마킹 보고서 표준 작성 가이드 v1.0",
    "summary": "경쟁사 비교 분석을 위한 체계적 벤치마킹 보고서 작성 기준",
    "main_sections": [
        {"id":"1","name":"경영진 요약 (Executive Summary)",
         "items":["핵심 발견사항 Top 5","전략적 시사점","즉각적 액션 아이템"],
         "criteria":["명확성","실행가능성","비즈니스 임팩트"]},
        {"id":"2","name":"분석 범위 및 방법론",
         "items":["분석 목적","대상 기업/제품","데이터 수집 방법","분석 기간"],
         "criteria":["객관성","데이터 신뢰성"]},
        {"id":"3","name":"기업 및 제품 개요 비교",
         "items":["회사 규모/역사","제품 포트폴리오","시장 포지셔닝","가격 전략"],
         "criteria":["시장점유율","성장률","브랜드 인지도"]},
        {"id":"4","name":"고객 VOC 분석",
         "items":["고객 만족도 CSAT","NPS 비교","주요 불만 사항","칭찬 포인트","개선 요청"],
         "criteria":["감성 분석","빈도 분석","영향도 분석"]},
        {"id":"5","name":"제품/서비스 사양 및 성능 비교",
         "items":["핵심 기능 매트릭스","성능 벤치마크","기술 사양","가격 대비 성능"],
         "criteria":["기능 완성도","성능 우위","혁신성"]},
        {"id":"6","name":"차별화 강점·약점·갭 분석",
         "items":["고유 기능 분석","SWOT 매트릭스","기능 갭","기회 영역"],
         "criteria":["차별화 지속 가능성","모방 난이도"]},
        {"id":"7","name":"전략적 권고사항",
         "items":["단기 액션 0~6M","중기 전략 6~18M","장기 로드맵 18M+"],
         "criteria":["ROI 예상","실행 난이도","우선순위"]},
    ],
    "scoring_dimensions": ["기능성","성능","가격경쟁력","고객만족도","혁신성","지원/서비스"],
    "scoring_scale": "1~10점 척도 (10점: 압도적 우위)",
    "data_sources": ["웹 리뷰 (블로그/커뮤니티)","공식 제품 스펙 페이지","유통 채널 리뷰","SNS VOC","IT 전문 미디어"],
    "analysis_methods": ["감성 분석","갭 분석","SWOT 분석","기능 매트릭스","가치 맵"],
}


def build_guide_from_files(texts: list) -> dict:
    combined = "\n\n━━ 문서 구분 ━━\n\n".join(texts[:15])
    if len(combined) > 10_000:
        combined = combined[:10_000] + "\n...[토큰 제한으로 생략됨]"

    system = "당신은 벤치마킹 보고서 전문 컨설턴트입니다. 여러 보고서를 분석하여 표준 가이드를 JSON으로 추출합니다. 한국어로 답하고, 반드시 순수 JSON만 출력하세요."
    prompt = f"""아래는 여러 벤치마킹 보고서의 내용입니다. 공통 패턴을 분석해 표준 작성 가이드를 추출하세요.

[보고서 내용]
{combined}

다음 JSON 스키마로 정확히 출력하세요 (주석·마크다운 없이 JSON만):
{{
  "title": "벤치마킹 보고서 표준 작성 가이드 v1.0",
  "summary": "한 문장 요약",
  "main_sections": [
    {{"id":"1","name":"섹션명","items":["핵심항목1","핵심항목2"],"criteria":["평가기준1"]}}
  ],
  "scoring_dimensions": ["차원1","차원2"],
  "scoring_scale": "평가 척도 설명",
  "data_sources": ["출처1","출처2"],
  "analysis_methods": ["방법1","방법2"]
}}"""

    raw = call_llm(prompt, max_tokens=3000, system=system)
    parsed = safe_json(raw)
    if parsed and isinstance(parsed, dict) and "main_sections" in parsed:
        return parsed
    guide = DEFAULT_GUIDE.copy()
    guide["llm_notes"] = raw
    return guide


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# WEB SEARCH & VOC
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def ddg_search(query: str, n: int = 8, region: str = "kr-kr") -> list:
    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=n, region=region))
        time.sleep(0.8)
        return results
    except Exception as e:
        st.warning(f"검색 실패 ({query[:40]}…): {e}")
        return []

def collect_voc(company: str, product: str, n_per_query: int = 8) -> list:
    queries = [
        f"{company} {product} 사용후기 리뷰 장단점",
        f"{company} {product} VOC 불만 개선사항",
        f'"{company}" "{product}" review pros cons',
        f"{company} {product} 고객만족 평가",
    ]
    seen, results = set(), []
    for q in queries:
        for r in ddg_search(q, n_per_query):
            url = r.get("href", "")
            if url not in seen:
                seen.add(url)
                results.append(r)
    return results

def collect_specs(company: str, product: str) -> list:
    queries = [
        f"{company} {product} 제품 사양 스펙 specification",
        f"{company} {product} 기능 비교 성능",
        f"{company} {product} official features datasheet",
    ]
    results = []
    for q in queries:
        results.extend(ddg_search(q, 6))
    return results

def analyze_voc(company: str, product: str, items: list) -> dict:
    snippets = "\n".join(
        f"[{i+1}] 제목: {r.get('title','')}\n내용: {(r.get('body') or r.get('snippet',''))[:400]}"
        for i, r in enumerate(items[:20])
    )
    if not snippets.strip():
        return {"error": "VOC 데이터 없음"}

    system = "당신은 고객 경험(CX) 분석 전문가입니다. VOC 데이터를 분석해 순수 JSON으로만 답변하세요."
    prompt = f"""다음은 '{company} {product}'에 대한 실제 사용자 VOC/리뷰입니다.

[VOC 원문]
{snippets}

다음 JSON 스키마로 분석 결과를 출력하세요:
{{
  "company": "{company}",
  "product": "{product}",
  "voc_count": 0,
  "overall_sentiment": "긍정/부정/보통",
  "satisfaction_score": 7.5,
  "top_positives": ["강점1","강점2","강점3"],
  "top_negatives": ["약점1","약점2","약점3"],
  "key_features_praised": ["칭찬기능1","기능2"],
  "key_features_criticized": ["비판기능1","기능2"],
  "improvement_requests": ["개선요청1","요청2"],
  "trending_topics": ["이슈1","이슈2"],
  "notable_snippets": ["인용1","인용2"]
}}"""

    raw = call_llm(prompt, max_tokens=1500, system=system)
    result = safe_json(raw)
    return result if isinstance(result, dict) else {"raw": raw, "company": company}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SPEC COMPARISON
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def compare_specs(my_co, my_prod, comp_co, comp_prod, my_text, comp_text) -> dict:
    system = "당신은 IT 제품 분석 전문가입니다. 두 제품을 객관적으로 비교해 순수 JSON으로만 출력하세요."
    prompt = f"""다음 두 제품의 정보를 비교 분석해주세요.

[자사: {my_co} – {my_prod}]
{my_text[:2500] if my_text else '직접 입력된 사양 없음 (웹 수집 데이터 활용)'}

[경쟁사: {comp_co} – {comp_prod}]
{comp_text[:2500] if comp_text else '직접 입력된 사양 없음 (웹 수집 데이터 활용)'}

다음 JSON으로 출력하세요:
{{
  "overall_summary": "한 문장 종합 비교",
  "overall_winner": "{my_co} 우위 / {comp_co} 우위 / 동등",
  "spec_matrix": [
    {{"category":"카테고리","my_value":"자사 사양","comp_value":"경쟁사 사양","winner":"{my_co}/{comp_co}/동등","importance":"상/중/하","note":"비고"}}
  ],
  "my_strengths": ["자사강점1","강점2","강점3"],
  "comp_strengths": ["경쟁사강점1","강점2","강점3"],
  "my_unique_only": ["자사고유기능1","기능2"],
  "comp_unique_only": ["경쟁사고유기능1","기능2"],
  "my_gaps": ["자사부족영역1","영역2"],
  "comp_gaps": ["경쟁사부족영역1","영역2"],
  "differentiation_score": {{
    "functionality":   {{"my":7,"comp":8}},
    "performance":     {{"my":8,"comp":7}},
    "price_value":     {{"my":7,"comp":6}},
    "innovation":      {{"my":8,"comp":9}},
    "customer_support":{{"my":7,"comp":7}}
  }},
  "strategic_recommendation": "전략적 권고사항"
}}"""

    raw = call_llm(prompt, max_tokens=3000, system=system)
    result = safe_json(raw)
    return result if isinstance(result, dict) else {"raw": raw}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# REPORT GENERATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def generate_report(guide, my_co, my_prod, comp_co, comp_prod, my_voc, comp_voc, spec_comp) -> str:
    guide_str  = json.dumps(guide,     ensure_ascii=False)[:1800]
    my_voc_s   = json.dumps(my_voc,    ensure_ascii=False)[:1000]
    comp_voc_s = json.dumps(comp_voc,  ensure_ascii=False)[:1000]
    spec_s     = json.dumps(spec_comp, ensure_ascii=False)[:2000]
    today      = datetime.now().strftime("%Y년 %m월 %d일")

    system = (
        "당신은 McKinsey 수준의 전략 컨설턴트입니다. "
        "데이터 기반의 통찰력 있는 벤치마킹 보고서를 한국어 마크다운으로 작성합니다. "
        "구체적 수치, 비교표, 실행 가능한 권고사항을 반드시 포함하세요."
    )
    prompt = f"""아래 데이터를 바탕으로 완전한 벤치마킹 보고서를 작성하세요.

## 메타 정보
- 자사: {my_co} ({my_prod})
- 경쟁사: {comp_co} ({comp_prod})
- 보고서 작성일: {today}

## 작성 가이드
{guide_str}

## VOC 분석 데이터
- 자사 VOC: {my_voc_s}
- 경쟁사 VOC: {comp_voc_s}

## 제품 사양 비교 데이터
{spec_s}

---

아래 구조로 전문 보고서를 작성하세요 (마크다운 형식):

# 📊 벤치마킹 보고서: {my_co} vs {comp_co}
**작성일:** {today} | **대상 제품:** {my_prod} vs {comp_prod}

---

## 1. 경영진 요약 (Executive Summary)
> 핵심 발견사항 5가지와 전략적 시사점

## 2. 분석 개요
## 3. 기업/제품 개요 비교 (마크다운 테이블)
## 4. 고객 VOC 심층 분석
### 4.1 {my_co} 고객 반응
### 4.2 {comp_co} 고객 반응
### 4.3 VOC 비교 시사점
## 5. 제품 사양 및 성능 상세 비교 (매트릭스 테이블)
## 6. 차별화 강점·약점·갭 분석
### 6.1 {my_co} 고유 강점
### 6.2 {my_co} 취약 영역 및 갭
### 6.3 {comp_co} 고유 강점
### 6.4 기능 갭 요약표
## 7. SWOT 분석 (2x2 마크다운 테이블)
## 8. 전략적 권고사항
### 8.1 단기 액션 플랜 (0~6개월)
### 8.2 중기 전략 (6~18개월)
### 8.3 장기 로드맵 (18개월+)
## 9. 결론

---
*본 보고서는 AI 벤치마킹 시스템으로 자동 생성되었습니다.*
*생성 모델: {st.session_state.model_id}*"""

    return call_llm(prompt, max_tokens=4096, system=system)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SIDEBAR
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with st.sidebar:
    st.markdown("## ⚙️ 시스템 설정")
    st.markdown("---")

    st.markdown("### 🔑 HuggingFace 인증")
    hf_token = st.text_input(
        "API Token (HF_TOKEN)", value=st.session_state.hf_token,
        type="password", placeholder="hf_xxxxxxxxxxxxxxxxxxxx",
        help="huggingface.co → Settings → Access Tokens 에서 발급",
    )
    if hf_token:
        st.session_state.hf_token = hf_token

    # ── 모델 선택 ──────────────────────────────────────────────────────────
    st.markdown("### 🤖 LLM 모델 선택")

    # ── 지원 모델 목록 ───────────────────────────────────────────────────────
    # HuggingFace Inference API 'conversational' 태스크 지원 모델만 등록
    PRESET_MODELS = {
        "🌟 Google Gemma 4 26B Instruct":   "google/gemma-4-26B-A4B-it",
        "⚡ Zyphra ZAYA1 8B (경량·빠름)":    "Zyphra/ZAYA1-8B",
        "🔵 Mistral 7B Instruct v0.3":      "mistralai/Mistral-7B-Instruct-v0.3",
        "🟢 Meta Llama 3.1 8B Instruct":    "meta-llama/Llama-3.1-8B-Instruct",
        "🟠 Qwen2.5 72B Instruct":          "Qwen/Qwen2.5-72B-Instruct",
        "✏️ 직접 입력":                      "__custom__",
    }

    # 현재 model_id 가 프리셋에 있으면 해당 항목, 없으면 직접 입력으로
    _reverse = {v: k for k, v in PRESET_MODELS.items() if v != "__custom__"}
    _default_label = _reverse.get(st.session_state.model_id, "✏️ 직접 입력")

    selected_label = st.selectbox(
        "모델 선택",
        options=list(PRESET_MODELS.keys()),
        index=list(PRESET_MODELS.keys()).index(_default_label),
        help="HuggingFace Inference API가 지원하는 모델만 동작합니다.",
    )

    selected_model_id = PRESET_MODELS[selected_label]

    if selected_model_id == "__custom__":
        custom_id = st.text_input(
            "모델 ID 직접 입력",
            value=(st.session_state.model_id
                   if st.session_state.model_id not in _reverse.values()
                   else ""),
            placeholder="예: meta-llama/Llama-3.1-8B-Instruct",
        )
        if custom_id.strip():
            st.session_state.model_id = custom_id.strip()
    else:
        st.session_state.model_id = selected_model_id

    # 현재 선택된 모델 표시 + 태스크 안내
    _short = st.session_state.model_id.split('/')[-1]
    st.markdown(
        f'<div style="background:#0f172a;border:1px solid #334155;border-radius:8px;'
        f'padding:8px 12px;margin-top:4px;font-size:0.78rem;color:#94a3b8;">'
        f'적용 모델: <code style="color:#c4b5fd">{_short}</code></div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="font-size:0.72rem;color:#475569;margin-top:5px;padding:6px 10px;'
        'background:#0f172a;border-radius:6px;border:1px solid #1e293b;">'
        '✅ Chat Completions (conversational) 태스크 사용<br>'
        '⚠️ text-generation 전용 모델은 지원 안됨</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.hf_token:
        st.markdown('<span class="badge-ok">✅ API 토큰 설정됨</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="badge-warn">⚠️ 토큰 미설정</span>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📊 작업 현황")
    for label, done in [
        ("가이드 추출", st.session_state.guide_extracted),
        ("자사 VOC",    bool(st.session_state.voc_analyzed_my)),
        ("경쟁사 VOC",  bool(st.session_state.voc_analyzed_comp)),
        ("사양 비교",   bool(st.session_state.spec_comparison)),
        ("보고서",      bool(st.session_state.report_md)),
    ]:
        color = "#6ee7b7" if done else "#64748b"
        icon  = "✅" if done else "⬜"
        st.markdown(f'<div style="color:{color};font-size:0.9rem;padding:3px 0">{icon} {label}</div>', unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🔄 전체 초기화", use_container_width=True):
        for k in list(DEFAULTS.keys()):
            if k not in ("hf_token","model_id"):
                st.session_state[k] = DEFAULTS[k]
        st.rerun()

    st.markdown('<div style="color:#475569;font-size:0.75rem;text-align:center">AI 벤치마킹 시스템 v2.0<br>Powered by HuggingFace</div>', unsafe_allow_html=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HERO HEADER
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown("""
<div class="hero-header">
  <div class="hero-title">📊 AI 벤치마킹 보고서 자동 작성 시스템</div>
  <div class="hero-subtitle">벤치마킹 가이드 추출 → 실시간 VOC 수집 → 제품 사양 비교 → 전문 보고서 자동 생성</div>
  <div class="hero-badge">🤖 Powered by HuggingFace LLM</div>
</div>
""", unsafe_allow_html=True)

# 현재 선택된 모델을 히어로 아래 실시간 표시
st.markdown(
    f'<div style="text-align:center;margin:-12px 0 18px 0;font-size:0.8rem;color:#64748b;">' +
    f'적용 모델: <code style="color:#a78bfa;background:#1e1b4b;' +
    f'padding:2px 8px;border-radius:4px;">{st.session_state.model_id}</code></div>',
    unsafe_allow_html=True,
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MAIN TABS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Step 1 · 가이드 추출",
    "🗣️ Step 2 · VOC 수집",
    "🔬 Step 3 · 사양 비교",
    "📝 Step 4 · 보고서 생성",
    "📖 도움말",
])

# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 – GUIDE EXTRACTION
# ──────────────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown("## 📋 벤치마킹 보고서 작성 기준 가이드 추출")
    st.markdown("기존 벤치마킹 보고서 파일을 업로드하면 AI가 공통 구조·핵심 항목·평가 기준을 자동으로 분석하여 표준 가이드를 생성합니다.")

    with st.expander("📁 보고서 파일 업로드 (PDF · DOCX · TXT · MD)", expanded=True):
        uploaded = st.file_uploader(
            "파일을 드래그하거나 클릭하여 선택 (다중 선택 가능)",
            accept_multiple_files=True,
            type=["pdf","docx","doc","txt","md"],
        )
        if uploaded:
            st.success(f"✅ {len(uploaded)}개 파일 업로드 완료")
            cols = st.columns(3)
            for i, f in enumerate(uploaded):
                cols[i%3].markdown(
                    f'<div class="step-card"><h4>📄 {f.name}</h4><p>{f.size:,} bytes</p></div>',
                    unsafe_allow_html=True,
                )

    c1, c2, c3 = st.columns([2,2,1])
    with c1:
        if st.button("🚀 AI로 가이드 자동 추출", type="primary", use_container_width=True, disabled=not uploaded):
            if not st.session_state.hf_token:
                st.error("❌ HuggingFace API 토큰을 먼저 입력해주세요.")
            else:
                prog = st.progress(0, "파일 텍스트 추출 중…")
                texts = []
                for i, f in enumerate(uploaded):
                    prog.progress((i+1)/len(uploaded), f"추출 중: {f.name}")
                    t = extract_file(f)
                    if t.strip():
                        texts.append(f"[파일: {f.name}]\n{t}")
                if texts:
                    prog.progress(1.0, "LLM으로 가이드 생성 중…")
                    with st.spinner("🤖 AI가 보고서 패턴을 분석하고 있습니다…"):
                        guide = build_guide_from_files(texts)
                    st.session_state.benchmark_guide = guide
                    st.session_state.guide_extracted = True
                    prog.empty()
                    st.success(f"✅ {len(texts)}개 파일 분석 완료! 가이드가 생성되었습니다.")
                else:
                    st.error("파일에서 텍스트를 추출할 수 없었습니다.")

    with c2:
        if st.button("📝 기본 표준 가이드 사용", use_container_width=True):
            st.session_state.benchmark_guide = DEFAULT_GUIDE
            st.session_state.guide_extracted = True
            st.success("✅ 표준 가이드가 적용되었습니다.")

    with c3:
        if st.session_state.benchmark_guide:
            st.download_button(
                "💾 JSON 저장",
                json.dumps(st.session_state.benchmark_guide, ensure_ascii=False, indent=2),
                "benchmark_guide.json", "application/json", use_container_width=True,
            )

    if st.session_state.benchmark_guide:
        st.markdown("---")
        g = st.session_state.benchmark_guide
        st.markdown(f"### 📌 {g.get('title','가이드')}")
        if g.get("summary"):
            st.info(g["summary"])
        if g.get("llm_notes"):
            st.text_area("LLM 분석 원문", g["llm_notes"], height=200)

        sections = g.get("main_sections", [])
        if sections:
            st.markdown("#### 📂 주요 섹션 구성")
            cols_g = st.columns(2)
            for idx, sec in enumerate(sections):
                with cols_g[idx%2]:
                    with st.expander(f"{sec.get('id','')+'.' if 'id' in sec else ''} {sec.get('name','섹션')}"):
                        if sec.get("items"):
                            st.markdown("**핵심 항목:**")
                            for it in sec["items"]:
                                st.markdown(f"  - {it}")
                        if sec.get("criteria"):
                            st.markdown("**평가 기준:**")
                            for cr in sec["criteria"]:
                                st.markdown(f"  - {cr}")

        r1, r2 = st.columns(2)
        with r1:
            if g.get("scoring_dimensions"):
                st.markdown("#### 🎯 평가 차원")
                for d in g["scoring_dimensions"]: st.markdown(f"  - {d}")
        with r2:
            if g.get("analysis_methods"):
                st.markdown("#### 🔬 분석 방법론")
                for m in g["analysis_methods"]: st.markdown(f"  - {m}")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 – VOC COLLECTION
# ──────────────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown("## 🗣️ 실시간 VOC 수집 및 AI 감성 분석")
    st.markdown("DuckDuckGo 검색으로 인터넷의 최신 사용후기·리뷰·커뮤니티 글을 수집하고, LLM이 자동으로 감성 분석합니다.")

    v1, v2 = st.columns(2)
    with v1:
        st.markdown("### 🏢 자사")
        my_co_v = st.text_input("회사명", key="v_my_co", placeholder="예: 삼성전자")
        my_pr_v = st.text_input("제품/서비스명", key="v_my_pr", placeholder="예: Galaxy S25 Ultra")
    with v2:
        st.markdown("### 🏭 경쟁사")
        cp_co_v = st.text_input("회사명", key="v_cp_co", placeholder="예: Apple")
        cp_pr_v = st.text_input("제품/서비스명", key="v_cp_pr", placeholder="예: iPhone 16 Pro Max")

    n_voc = st.slider("쿼리당 수집 건수", 4, 15, 8)

    if st.button("🔍 VOC 수집 및 분석 시작", type="primary", use_container_width=True, disabled=not (my_co_v and cp_co_v)):
        if not st.session_state.hf_token:
            st.error("❌ HuggingFace API 토큰을 먼저 입력해주세요.")
        else:
            st.session_state.update(my_company=my_co_v, my_product=my_pr_v, competitor=cp_co_v, comp_product=cp_pr_v)
            p1, p2 = st.columns(2)
            with p1:
                with st.spinner(f"🔍 {my_co_v} VOC 수집 중…"):
                    raw_my = collect_voc(my_co_v, my_pr_v, n_voc)
                    st.session_state.voc_my = raw_my
                with st.spinner(f"🤖 {my_co_v} VOC 분석 중…"):
                    st.session_state.voc_analyzed_my = analyze_voc(my_co_v, my_pr_v, raw_my)
                st.success(f"✅ {my_co_v}: {len(raw_my)}건 수집 완료")
            with p2:
                with st.spinner(f"🔍 {cp_co_v} VOC 수집 중…"):
                    raw_cp = collect_voc(cp_co_v, cp_pr_v, n_voc)
                    st.session_state.voc_comp = raw_cp
                with st.spinner(f"🤖 {cp_co_v} VOC 분석 중…"):
                    st.session_state.voc_analyzed_comp = analyze_voc(cp_co_v, cp_pr_v, raw_cp)
                st.success(f"✅ {cp_co_v}: {len(raw_cp)}건 수집 완료")

    def render_voc_panel(col, data, raw, label):
        with col:
            score     = data.get("satisfaction_score","N/A")
            sentiment = data.get("overall_sentiment","N/A")
            count     = data.get("voc_count", len(raw))
            st.markdown(f"#### 🏢 {label}")
            m1,m2,m3 = st.columns(3)
            m1.metric("만족도",   f"{score}/10" if score!="N/A" else "N/A")
            m2.metric("감성",     sentiment)
            m3.metric("수집 VOC", f"{count}건")
            if data.get("raw"):
                st.text_area("LLM 원문", data["raw"][:1000], height=120); return
            for title, key, icon in [
                ("주요 긍정 포인트","top_positives","✅"),
                ("주요 불만 사항","top_negatives","❌"),
                ("개선 요청","improvement_requests","💡"),
                ("트렌딩 이슈","trending_topics","🔥"),
                ("칭찬받는 기능","key_features_praised","⭐"),
                ("비판받는 기능","key_features_criticized","⚠️"),
            ]:
                if data.get(key):
                    with st.expander(f"{icon} {title}"):
                        for item in data[key]: st.markdown(f"- {item}")
            with st.expander(f"🔎 원시 VOC 미리보기 ({len(raw)}건 중 5건)"):
                for i, item in enumerate(raw[:5]):
                    st.markdown(
                        f'<div class="voc-card"><div class="voc-title">{i+1}. {item.get("title","N/A")}</div>'
                        f'<div class="voc-body">{(item.get("body") or item.get("snippet",""))[:300]}</div></div>',
                        unsafe_allow_html=True,
                    )

    if st.session_state.voc_analyzed_my or st.session_state.voc_analyzed_comp:
        st.markdown("---"); st.markdown("### 📊 VOC 분석 결과")
        d1, d2 = st.columns(2)
        render_voc_panel(d1, st.session_state.voc_analyzed_my,  st.session_state.voc_my,  f"{st.session_state.my_company} {st.session_state.my_product}")
        render_voc_panel(d2, st.session_state.voc_analyzed_comp,st.session_state.voc_comp, f"{st.session_state.competitor} {st.session_state.comp_product}")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 – SPEC COMPARISON
# ──────────────────────────────────────────────────────────────────────────────
with tab3:
    st.markdown("## 🔬 제품 사양 및 성능 비교 분석")
    s1, s2 = st.columns(2)
    with s1:
        st.markdown("### 🏢 자사 제품")
        sp_my_co = st.text_input("회사명", value=st.session_state.my_company,  key="sp_my_co")
        sp_my_pr = st.text_input("제품명", value=st.session_state.my_product,  key="sp_my_pr")
        my_spec_txt = st.text_area("제품 사양/특징 (직접 입력)", value=st.session_state.spec_my_raw, height=220,
            placeholder="예:\n- AP: Snapdragon 8 Elite\n- 배터리: 5000mAh\n- 카메라: 200MP", key="sp_my_spec")
        st.session_state.spec_my_raw = my_spec_txt
    with s2:
        st.markdown("### 🏭 경쟁사 제품")
        sp_cp_co = st.text_input("회사명",   value=st.session_state.competitor,   key="sp_cp_co")
        sp_cp_pr = st.text_input("제품명",   value=st.session_state.comp_product, key="sp_cp_pr")
        cp_spec_txt = st.text_area("제품 사양/특징 (직접 입력)", value=st.session_state.spec_comp_raw, height=220,
            placeholder="예:\n- AP: A18 Pro\n- 배터리: 4422mAh\n- 카메라: 48MP 트리플", key="sp_cp_spec")
        st.session_state.spec_comp_raw = cp_spec_txt

    b1, b2 = st.columns(2)
    with b1:
        if st.button("🌐 웹에서 사양 자동 수집", use_container_width=True, disabled=not (sp_my_co and sp_cp_co)):
            with st.spinner("웹 검색으로 사양 수집 중…"):
                my_res  = collect_specs(sp_my_co, sp_my_pr)
                cp_res  = collect_specs(sp_cp_co, sp_cp_pr)
                st.session_state.spec_my_raw   = "\n".join(f"[{r.get('title','')}] {r.get('body',r.get('snippet',''))}" for r in my_res[:6])[:3000]
                st.session_state.spec_comp_raw = "\n".join(f"[{r.get('title','')}] {r.get('body',r.get('snippet',''))}" for r in cp_res[:6])[:3000]
                st.success("✅ 웹 사양 수집 완료!"); st.rerun()
    with b2:
        if st.button("⚡ AI 사양 비교 분석 실행", type="primary", use_container_width=True):
            if not st.session_state.hf_token:
                st.error("❌ HuggingFace API 토큰을 먼저 입력해주세요.")
            elif not (sp_my_co and sp_cp_co):
                st.error("회사명을 모두 입력해주세요.")
            else:
                st.session_state.update(my_company=sp_my_co, my_product=sp_my_pr, competitor=sp_cp_co, comp_product=sp_cp_pr)
                with st.spinner("🤖 AI가 사양을 비교 분석 중…"):
                    result = compare_specs(sp_my_co, sp_my_pr, sp_cp_co, sp_cp_pr,
                                           st.session_state.spec_my_raw, st.session_state.spec_comp_raw)
                    st.session_state.spec_comparison = result
                st.success("✅ 사양 비교 완료!")

    if st.session_state.spec_comparison:
        st.markdown("---")
        comp = st.session_state.spec_comparison
        if comp.get("raw"):
            st.text_area("LLM 분석 원문", comp["raw"][:2000], height=200)
        else:
            st.markdown("### 📊 비교 분석 결과")
            if comp.get("overall_summary"): st.info(f"**종합 평가:** {comp['overall_summary']}")
            if comp.get("overall_winner"):  st.metric("🏆 종합 우위", comp["overall_winner"])

            if comp.get("spec_matrix"):
                st.markdown("#### 📋 상세 사양 비교 매트릭스")
                rows = []
                for row in comp["spec_matrix"]:
                    w = row.get("winner","")
                    rows.append({
                        "카테고리": row.get("category",""),
                        "중요도":   row.get("importance",""),
                        f"{sp_my_co}": row.get("my_value",""),
                        f"{sp_cp_co}": row.get("comp_value",""),
                        "우위": f"✅ {w}" if sp_my_co in w else (f"❌ {w}" if w!="동등" else "🟡 동등"),
                        "비고": row.get("note",""),
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            sg1, sg2 = st.columns(2)
            with sg1:
                for title, key in [("강점","my_strengths"),("고유 기능","my_unique_only"),("취약 영역","my_gaps")]:
                    if comp.get(key):
                        st.markdown(f"#### {sp_my_co} {title}")
                        for item in comp[key]: st.markdown(f"- {item}")
            with sg2:
                for title, key in [("강점","comp_strengths"),("고유 기능","comp_unique_only"),("취약 영역","comp_gaps")]:
                    if comp.get(key):
                        st.markdown(f"#### {sp_cp_co} {title}")
                        for item in comp[key]: st.markdown(f"- {item}")

            if comp.get("differentiation_score"):
                st.markdown("#### 🎯 차별화 점수 비교")
                sdf = pd.DataFrame([
                    {"평가 차원": d, f"{sp_my_co} 점수": v.get("my",0), f"{sp_cp_co} 점수": v.get("comp",0)}
                    for d, v in comp["differentiation_score"].items()
                ])
                st.dataframe(sdf, use_container_width=True, hide_index=True)

            if comp.get("strategic_recommendation"):
                st.success(f"💡 **전략적 권고:** {comp['strategic_recommendation']}")


# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 – REPORT GENERATION
# ──────────────────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("## 📝 벤치마킹 보고서 자동 생성")
    r1c,r2c,r3c,r4c = st.columns(4)
    r1c.metric("가이드",   "✅" if st.session_state.guide_extracted       else "⚠️")
    r2c.metric("자사 VOC", "✅" if st.session_state.voc_analyzed_my      else "⚠️")
    r3c.metric("경쟁사 VOC","✅" if st.session_state.voc_analyzed_comp   else "⚠️")
    r4c.metric("사양 비교","✅" if st.session_state.spec_comparison       else "⚠️")

    with st.expander("⚙️ 보고서 기본 정보"):
        re1, re2 = st.columns(2)
        with re1:
            r_my_co = st.text_input("자사명",   value=st.session_state.my_company,  key="r_my_co")
            r_my_pr = st.text_input("자사 제품", value=st.session_state.my_product,  key="r_my_pr")
        with re2:
            r_cp_co = st.text_input("경쟁사명",   value=st.session_state.competitor,   key="r_cp_co")
            r_cp_pr = st.text_input("경쟁사 제품", value=st.session_state.comp_product, key="r_cp_pr")

    if st.button("🚀 벤치마킹 보고서 생성 (LLM 기반)", type="primary", use_container_width=True, disabled=not st.session_state.hf_token):
        if not (r_my_co and r_cp_co):
            st.error("자사명과 경쟁사명을 입력해주세요.")
        else:
            if not st.session_state.benchmark_guide:
                st.session_state.benchmark_guide = DEFAULT_GUIDE
                st.session_state.guide_extracted = True
            with st.spinner("🤖 AI가 벤치마킹 보고서를 작성 중입니다… (2~5분 소요)"):
                report = generate_report(
                    st.session_state.benchmark_guide,
                    r_my_co, r_my_pr, r_cp_co, r_cp_pr,
                    st.session_state.voc_analyzed_my, st.session_state.voc_analyzed_comp,
                    st.session_state.spec_comparison,
                )
                st.session_state.report_md = report
            st.success("✅ 보고서 생성 완료!")

    if st.session_state.report_md:
        st.markdown("---")
        st.markdown("### 📄 생성된 보고서")
        st.markdown(st.session_state.report_md)
        st.markdown("---")

        today_str = datetime.now().strftime("%Y%m%d_%H%M")
        fname = f"benchmarking_{st.session_state.my_company}_vs_{st.session_state.competitor}_{today_str}"
        dl1, dl2, dl3 = st.columns(3)

        with dl1:
            st.download_button("📥 Markdown 다운로드", st.session_state.report_md,
                               f"{fname}.md", "text/markdown", use_container_width=True)
        with dl2:
            html_doc = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<title>벤치마킹 보고서</title>
<style>body{{font-family:'Malgun Gothic',sans-serif;max-width:1100px;margin:40px auto;padding:0 24px;color:#1e293b;line-height:1.8}}
h1{{color:#1e40af;border-bottom:3px solid #1e40af;padding-bottom:12px}}h2{{color:#2a5298}}
table{{border-collapse:collapse;width:100%;margin:16px 0}}th{{background:#1e40af;color:white;padding:10px 14px}}
td{{padding:9px 14px;border-bottom:1px solid #e2e8f0}}tr:nth-child(even){{background:#f8fafc}}</style>
</head><body>{st.session_state.report_md.replace(chr(10),'<br>')}</body></html>"""
            st.download_button("📥 HTML 다운로드", html_doc,
                               f"{fname}.html", "text/html", use_container_width=True)
        with dl3:
            full_json = json.dumps({
                "metadata": {"generated_at": datetime.now().isoformat(), "model": st.session_state.model_id,
                             "my_company": st.session_state.my_company, "competitor": st.session_state.competitor},
                "guide": st.session_state.benchmark_guide,
                "voc_my": st.session_state.voc_analyzed_my,
                "voc_comp": st.session_state.voc_analyzed_comp,
                "spec_comparison": st.session_state.spec_comparison,
                "report_md": st.session_state.report_md,
            }, ensure_ascii=False, indent=2)
            st.download_button("📥 전체 JSON", full_json,
                               f"{fname}_full.json", "application/json", use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 5 – HELP
# ──────────────────────────────────────────────────────────────────────────────
with tab5:
    st.markdown("## 📖 사용 가이드 및 시스템 정보")
    with st.expander("🚀 빠른 시작 가이드", expanded=True):
        st.markdown("""
### 사전 준비
1. **HuggingFace 계정 생성** → [huggingface.co](https://huggingface.co)
2. **API Token 발급** → Settings → Access Tokens → New Token (Read 권한)
3. **모델 접근 동의** → `google/gemma-4-26B-A4B-it` 모델 페이지에서 라이선스 동의
4. 사이드바에 토큰 입력 후 사용 시작

### 4단계 워크플로우

| 단계 | 작업 | 소요 시간 |
|------|------|-----------|
| Step 1 | 보고서 파일 업로드 → AI 가이드 추출 | 1~3분 |
| Step 2 | 자사·경쟁사 입력 → VOC 자동 수집·분석 | 2~5분 |
| Step 3 | 사양 입력/수집 → AI 비교 분석 | 1~2분 |
| Step 4 | 버튼 클릭 → 전문 보고서 자동 생성 | 2~5분 |
""")
    with st.expander("⚙️ 기술 스택"):
        st.markdown("""
| 구성 요소 | 기술 |
|-----------|------|
| UI | Streamlit |
| LLM | google/gemma-4-26B-A4B-it (HuggingFace Inference API) |
| 웹 검색 | DuckDuckGo Search (무료) |
| 파일 파싱 | PyPDF2 · python-docx |
| 웹 스크래핑 | requests · BeautifulSoup4 |
| 데이터 처리 | pandas |
""")
    with st.expander("❓ 트러블슈팅"):
        st.markdown("""
| 증상 | 해결 방법 |
|------|-----------|
| 401 Unauthorized | HF 토큰 재확인·재발급 |
| Model not found | HF 모델 페이지에서 라이선스 동의 |
| VOC 수집 0건 | 검색어 수정 또는 잠시 후 재시도 |
| PDF 텍스트 없음 | OCR 전처리 후 재업로드 |
""")
    st.markdown(f"""
---
<div style="text-align:center;color:#475569;font-size:0.85rem">
AI 벤치마킹 보고서 자동 작성 시스템 v2.0<br>
HuggingFace LLM · DuckDuckGo Search · Streamlit<br>
현재 모델: <code>{st.session_state.model_id}</code>
</div>
""", unsafe_allow_html=True)
