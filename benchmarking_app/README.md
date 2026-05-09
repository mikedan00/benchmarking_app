# 📊 AI 벤치마킹 보고서 자동 작성 시스템

> HuggingFace LLM (`google/gemma-4-26B-A4B-it`) × DuckDuckGo Search × Streamlit

---

## 🗂️ 프로젝트 구조

```
benchmarking_app/
├── app.py                  ← 메인 Streamlit 애플리케이션
├── requirements.txt        ← Python 패키지 의존성
├── .env.example            ← 환경변수 템플릿
├── .env                    ← 실제 환경변수 (Git 제외)
├── Dockerfile              ← Docker 이미지 빌드
├── docker-compose.yml      ← 컨테이너 오케스트레이션
├── run.sh                  ← 로컬/Docker 통합 실행 스크립트
├── .gitignore
└── .streamlit/
    └── config.toml         ← Streamlit 테마 & 서버 설정
```

---

## ⚙️ 사전 준비

### 1. HuggingFace API 토큰 발급

1. [huggingface.co](https://huggingface.co) 회원가입
2. `Settings → Access Tokens → New Token` (Read 권한)
3. 아래 모델 페이지에서 라이선스 동의 **(필수)**:
   → [google/gemma-4-26B-A4B-it](https://huggingface.co/google/gemma-4-26B-A4B-it)

### 2. 환경변수 설정

```bash
cp .env.example .env
# .env 파일을 열어 HF_TOKEN 값 입력
```

---

## 🚀 실행 방법

### A. 로컬 직접 실행 (개발/테스트)

```bash
chmod +x run.sh
./run.sh
# → http://localhost:8501
```

수동 설치:
```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

### B. Docker (서버 배포)

```bash
./run.sh docker        # 빌드 & 백그라운드 실행
./run.sh stop          # 중지
docker compose logs -f # 로그 확인
```

### C. Streamlit Community Cloud (무료 배포)

1. GitHub 레포에 Push
2. [share.streamlit.io](https://share.streamlit.io) → `New app` → 레포 선택
3. `Advanced settings → Secrets`:
   ```toml
   HF_TOKEN = "hf_실제토큰값"
   ```
4. `Deploy!`

---

## 📋 기능 설명 (4단계 워크플로우)

| 단계 | 탭 | 기능 | 소요 시간 |
|------|-----|------|-----------|
| Step 1 | 가이드 추출 | 기존 보고서 PDF/DOCX 업로드 → LLM이 공통 구조·항목 추출 | 1~3분 |
| Step 2 | VOC 수집 | 자사·경쟁사 이름 입력 → DuckDuckGo 실시간 수집 → 감성 분석 | 2~5분 |
| Step 3 | 사양 비교 | 사양 직접 입력 또는 웹 수집 → AI 비교 매트릭스·갭 분석 | 1~2분 |
| Step 4 | 보고서 생성 | 1~3단계 데이터 통합 → McKinsey급 보고서 자동 작성 | 2~5분 |

---

## 🔧 모델 변경

사이드바 또는 `.env`에서:

```env
# 대안 모델 예시
HF_MODEL_ID=mistralai/Mixtral-8x7B-Instruct-v0.1
HF_MODEL_ID=meta-llama/Llama-3.1-70B-Instruct
HF_MODEL_ID=Qwen/Qwen2.5-72B-Instruct
```

---

## ❓ 트러블슈팅

| 증상 | 원인 | 해결 |
|------|------|------|
| `401 Unauthorized` | 토큰 오류 | HF 토큰 재확인 |
| `Model not found` | 접근 권한 없음 | 모델 페이지에서 라이선스 동의 |
| `Too many requests` | API 한도 초과 | 잠시 후 재시도 |
| VOC 수집 0건 | DuckDuckGo 차단 | VPN 변경 또는 검색어 수정 |
| PDF 텍스트 없음 | 스캔 이미지 PDF | OCR 전처리 후 재업로드 |

---

## 🏗️ 아키텍처

```
사용자 브라우저 (Streamlit UI)
        │
  ┌─────┴──────────────────────────┐
  │         app.py (메인 엔진)      │
  │  ┌──────────┐ ┌─────────────┐  │
  │  │File Parser│ │Web Crawler  │  │
  │  │PDF·DOCX  │ │DuckDuckGo  │  │
  │  └──────┬───┘ └──────┬──────┘  │
  │         └──────┬──────┘         │
  │    ┌───────────▼──────────┐     │
  │    │  LLM Orchestrator    │     │
  │    │  Prompt Engineering  │     │
  │    └───────────┬──────────┘     │
  └────────────────┼────────────────┘
                   ▼
      HuggingFace Inference API
      google/gemma-4-26B-A4B-it
                   │
          ┌────────┴────────┐
    JSON Guide  VOC분석  사양비교  보고서MD
```

---

*AI 벤치마킹 보고서 시스템 v2.0 | Streamlit + HuggingFace LLM*
