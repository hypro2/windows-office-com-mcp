# Windows Office COM MCP (`windows-office-com-mcp`)

Claude Desktop, Cursor, Antigravity IDE 같은 AI 에이전트가 **내 Windows PC에 열려 있는 MS Word, Excel, PowerPoint, 한컴오피스(HWP) 문서를 직접 읽고 편집**할 수 있게 해주는 통합 **MCP(Model Context Protocol) 서버**입니다.

---

## 💡 무엇을 할 수 있나요?

- **열려 있는 문서 실시간 읽기·편집**: 파일을 닫거나 따로 내보낼 필요 없이, 지금 화면에 띄워 둔 문서 본문과 마우스로 선택한 영역(`Selection`)을 AI가 바로 읽고 수정합니다.
- **깔끔한 보고서·공문서 서식 자동화**: 마크다운 글을 서식 있는 Word 보고서로 바꾸고, 표 스타일링('겉바속촉'), 한컴 HWP 공문서 양식, Excel 재무 모델 시트, PowerPoint 슬라이드까지 자연어 지시로 작성합니다.
- **내 PC 안에서 로컬 직접 실행**: Windows 데스크톱 오피스(Win32 COM 및 한컴 ROT `!HwpObject`)에 직접 연결되어 동작합니다.

---

## 📊 지원 프로그램 및 도구 구성

총 **87개 Tools · 3개 Resources · 3개 Prompts**를 제공합니다 (`전용 도구 52개 + 공용 도구 35개`).

| 프로그램 | 제공 도구 수 | AI에게 시킬 수 있는 대표적인 작업 |
| :--- | :---: | :--- |
| **MS PowerPoint** (`.pptx`) | **53개** (`ppt_*` 19 + 공용 34) | 메모·기획안 텍스트로 슬라이드 자동 생성, 113종 레이아웃 슬라이스 삽입, Morph 전환 애니메이션, Pitch Deck 구성 |
| **MS Word** (`.docx`) | **50개** (`word_*` 15 + 공용 35) | 마크다운 보고서 서식 변환, 표 서식 다듬기, 7단계 개조식·줄간격 교정, 체크박스/드롭다운 실제 입력 양식(SDT) 삽입, 학술 논문 규격(APA/IEEE/KCI 등) 적용 |
| **MS Excel** (`.xlsx`) | **47개** (`excel_*` 12 + 공용 35) | 선택 영역 읽기/쓰기, 표 자동 서식, 수식 실시간 평가, 공백·주민번호 마스킹 등 데이터 정제, 3-Zone 재무 모델(DCF/LBO) 생성 및 CFO 색상 감사 |
| **한컴오피스 HWP** (`.hwp`) | **37개** (`hwp_*` 6 + 공용 31) | 열려 있는 한글 문서 본문·선택 블록 읽기, 표 채우기, 행안부/KISA 1310 공문서 마크업(`제목:`, `네모:`, `표:` 등) 렌더링 |

### 4대 오피스 공용 기능 (`office_*` 35개)
- **비즈니스 도형 & 차트**: 조직도, 프로세스 흐름도, 타임라인, 비교표, KPI 카드, 결재란(Word·Excel), 코호트 리텐션 및 7종 네이티브 차트 삽입
- **표준 양식 검색 & 조립**: 283종 표준 비즈니스 참조 양식을 자연어로 검색해 삽입하고, 여러 블록을 조합한 커스텀 양식 생성
- **내 문서 스타일 복제 & 자동 문서 기획**: 기존 문서나 폴더의 폰트·컬러·여백 스타일을 학습해 새 문서에 그대로 적용하고, 거친 원시 메모만으로 알맞은 포맷을 골라 보고서 자동 생성
- **치환·리뷰·납품 전 품질 점검**: `{{key}}` 토큰 일괄 치환(메일 머지), 수정 제안 코멘트 마킹, 납품 전 품질 감사 게이트 및 PNG/PDF 미리보기 추출

> [!NOTE]
> 공용 도구 35개 중 **31개는 4대 포맷 모두 지원**, 3개(`office_harmonize_document`, `office_mark_review_item`, `office_export_visual_snapshot`)는 **Word·Excel·PPT 3종 지원**, 1개(`office_insert_approval_block`)는 **Word·Excel 2종 지원**입니다.

---

## 🚀 3분 빠른 시작

### 1. 사전 준비 및 패키지 설치
Windows 환경과 Python, 그리고 제어하려는 오피스 프로그램(MS Office 및/또는 한컴오피스 HWP)이 설치되어 있어야 합니다.

```powershell
git clone https://github.com/hypro2/windows-office-mcp.git windows-office-com-mcp
cd windows-office-com-mcp
pip install -r requirements.txt
```

### 2. 정상 동작 확인 (CLI)
```powershell
# 등록된 전체 도구(87개)·리소스(3개)·프롬프트(3개) 목록 확인
python -m mcp_server.client_demo --list-tools

# 현재 PC에 실행 중인 오피스 창과 활성 문서 연결 상태 확인
python -m mcp_server.client_demo --status

# 데모 실행 (all | live-com | components | dynamic)
python -m mcp_server.client_demo --demo all
```

### 3. AI 앱에 MCP 서버 연결하기
사용 중인 AI 데스크톱 앱 설정 파일에 아래 내용을 추가한 뒤 앱을 재시작하세요 (`cwd`는 실제 저장소 경로로 변경).
- **Claude Desktop**: `%APPDATA%\Claude\claude_desktop_config.json`
- **Cursor**: `.cursor/mcp.json`
- **Antigravity IDE**: MCP 설정 파일

```json
{
  "mcpServers": {
    "windows-office-com-mcp": {
      "command": "python",
      "args": [
        "-m",
        "mcp_server.server",
        "--transport",
        "stdio"
      ],
      "cwd": "C:\\path\\to\\windows-office-com-mcp"
    }
  }
}
```

<details>
<summary>웹/원격 에이전트용 SSE 모드로 직접 실행하기</summary>

```powershell
python -m mcp_server.server --transport sse --host 127.0.0.1 --port 8765
```

</details>

---

## 💬 자연어 활용 예시

1. **열려 있는 Word 문서 다듬기**
   > *"지금 열려 있는 워드 문서에서 내가 선택한 문단을 보고서 톤으로 다듬고, 문서 전체 글꼴과 표 서식을 깔끔하게 통일해 줘."*
2. **Excel 표 정리 & 재무 모델 생성**
   > *"현재 선택한 시트 표에 헤더 강조와 천 단위 콤마 서식을 넣고, 새 시트에 DCF 재무 모델 뼈대도 만들어 줘."*
3. **메모만으로 PPT 발표자료 만들기**
   > *"아래 회의 메모를 바탕으로 6장짜리 요약 프레젠테이션을 만들어 줘. 핵심 KPI 카드와 분기별 로드맵 슬라이드도 포함해 줘."*
4. **한컴오피스(HWP) 공문서 작성**
   > *"지금 띄워 둔 한글 문서 커서 위치에 추진 배경, 주요 일정표, 결재란까지 공문서 표준 서식으로 넣어 줘."*
5. **완성된 문서 최종 검수**
   > *"작성한 문서에 `{{치환 누락}}`이나 깨진 표·수식이 없는지 검사하고, 페이지 미리보기 PNG로 뽑아 줘."*

---

## 📂 프로젝트 구조

```text
windows-office-com-mcp/
├── mcp_server/                  # 실시간 데스크톱 COM/ROT 제어 통합 MCP 서버
│   ├── server.py                # FastMCP 진입점 (stdio / sse, 87 Tools · 3 Resources · 3 Prompts)
│   ├── com_manager.py           # Win32 COM 및 HWP ROT(!HwpObject) 실시간 창 연결 매니저
│   ├── client_demo.py           # MCP 클라이언트 검증 및 데모 CLI
│   ├── tools/                   # Word/Excel/PPT/HWP/공용 도구 모듈 & 참조·슬라이스 카탈로그 JSON
│   ├── assets/                  # 공용 템플릿/렌더링 보조 리소스
│   └── README.md                # 87개 전체 도구 목록 및 세부 사용 가이드
│
├── data/                        # PPT 기법 청사진, 컬러 팔레트, 카탈로그 JSON 데이터셋
└── requirements.txt             # Python 의존성 패키지 목록 (mcp, python-docx, python-pptx, openpyxl, pywin32 등)
```

> [!TIP]
> 87개 전체 도구의 정확한 이름과 역할은 [`mcp_server/README.md`](mcp_server/README.md)에서 확인할 수 있습니다. 사용자 템플릿 학습(`blueprints/`)과 폴더 서식 프로필(`user_profiles/`) 디렉터리는 해당 도구 실행 시 자동으로 생성됩니다.

---

## 🔗 참조 및 오픈소스 레퍼런스

본 프로젝트의 COM 자동화 아키텍처, 문서 구성 지능, 품질 감사 게이트는 아래 오픈소스 프로젝트 및 공개 표준 리서치를 참조하여 설계되었습니다.

| 구분 / 프로젝트 | 원출처 (Repository) | 주요 참조 기술 및 반영 영역 |
| :--- | :--- | :--- |
| **iOfficeAI Suite** (`AionUi`, `AionCore`, `OfficeCLI`) | [github.com/iOfficeAI/AionUi](https://github.com/iOfficeAI/AionUi)<br>[github.com/iOfficeAI/AionCore](https://github.com/iOfficeAI/AionCore)<br>[github.com/iOfficeAI/OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) | Morph 3-Prefix 안무(`!!scene-*`/`!!actor-*`/`#sN-*`), Word SDT·학술 조판, Excel 3-Zone·CFO 코딩, 5-Gate 납품 감사, 스냅샷·리뷰 마킹 |
| **Meta Astryx** | [github.com/meta/astryx](https://github.com/meta/astryx) | Anti-Pattern Guardrails 기반 사전 차단 규칙, 시맨틱 디자인 토큰 체계 |
| **Genspark GenOffice** | [github.com/genspark-ai/genoffice](https://github.com/genspark-ai/genoffice) | 바이트 보존 부분 패치(`SurgicalPatcher`), 단계별 파이프라인 및 공간 감사(`SpatialAuditor`) |
| **Presenton** | [github.com/presenton/presenton](https://github.com/presenton/presenton) | FastMCP 비동기 작업 큐(`AsyncJobManager`), 완전 편집 가능한 Native PPTX 지향 및 로컬 실행 구조 |
| **Hugo He ppt-master** | [github.com/hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) | SVG 중간 표현(IR) → 네이티브 도형 컴파일(`SvgPptCompiler`), 텍스트 메트릭 계산 |
| **HKUDS Paper2Slides** | [github.com/HKUDS/Paper2Slides](https://github.com/HKUDS/Paper2Slides) | 장문 문서의 표/그림 원본 무손실 추출 및 슬라이드 슬롯 바인딩(`DocumentElementExtractor`) |
| **KISA 1310 / 공문서 서식 표준** | KISA 공문서·비즈니스 서식 자동화 규격 | 한컴 HWP ROT(`!HwpObject`) 제어, 겉바속촉 표, 7단계 개조식 정제, 공문서 마크업 엔진 |
