# Windows Office COM MCP (`windows-office-com-mcp`) — 전체 도구 가이드

코딩 어시스턴트 데스크톱 앱(**Claude Desktop**, **Antigravity IDE**, **Cursor** 등)에서 자연어로 질의하면, **현재 화면에 열려 있는 오피스 문서(MS Word, Excel, PowerPoint, 한컴 HWP)의 본문과 마우스 선택 영역(Selection)을 읽고, 문서를 단계별로 작성하며, 표준 비즈니스 서식·KISA 1310 공문서 서식·283종 RAG 양식·자율 문서 구성(Composition Engine)·6-Gate 납품 품질 감사까지 일괄 수행**하는 MCP 서버입니다.

---

## 1. 제공 구성 요약 & 포맷별 사용 가능 도구 수

- **MCP Tools**: 총 **87개** (`전용 도구 52개 + 공용 도구 35개`)
- **MCP Resources**: 총 **3개** (`office://status`, `office://guides/kisa1310-hwp`, `office://guides/aionui-domain-playbooks`)
- **MCP Prompts**: 총 **3개** (`polish_active_selection`, `draft_official_hwp_report`, `create_dynamic_business_form`)

| 오피스 포맷 | 전용 도구 | 공용 지원 도구 (`office_*`) | 사용 가능 총 도구 수 | 미지원 공용 도구 (제외 항목) |
| :--- | :---: | :---: | :---: | :--- |
| **MS PowerPoint** (`.pptx`) | `19개` (`ppt_*`) | `34개` | **53개** | `office_insert_approval_block` (1개) |
| **MS Word** (`.docx`) | `15개` (`word_*`) | `35개` | **50개** | 없음 (`office_*` 35개 전수 지원) |
| **MS Excel** (`.xlsx`) | `12개` (`excel_*`) | `35개` | **47개** | 없음 (`office_*` 35개 전수 지원) |
| **한컴오피스 HWP** (`.hwp`) | `6개` (`hwp_*`) | `31개` | **37개** | `office_insert_approval_block`, `office_harmonize_document`, `office_mark_review_item`, `office_export_visual_snapshot` (4개) |

> **공용 도구(`office_*`) 35개 지원 범위 분류**
> - **4대 포맷 공통 (31개)**: 시스템 제어(3) · 비즈니스 컴포넌트(7) · 네이티브 차트/시각화(3) · RAG 양식/빌더(3) · 사용자 템플릿 클론(5) · 폴더 벌크 학습/프로필 관리(4) · 문서 구성 지능 Composition Engine(4) · 5-Gate 납품 감사 & 토큰 병합(2)
> - **Word · Excel · PPT 3종 (3개)**: `office_harmonize_document`, `office_mark_review_item`, `office_export_visual_snapshot`
> - **Word · Excel 2종 (1개)**: `office_insert_approval_block`

---

## 2. 카테고리별 전체 도구 목록 (87개 Tools)

### A. 공통 시스템 제어 (`system_tools.py` — 3개 / 4대 포맷 공통)
- `office_get_status`: Word, Excel, PowerPoint, HWP 설치/실행 여부, 활성 문서명, 마우스 선택 영역(Selection) 미리보기 조회
- `office_launch_or_open`: 오피스 앱 실행 또는 지정 문서 파일을 데스크톱 창에 열기
- `office_run_com_script`: 세부 Python COM/VBA 로직 직접 실행 (50줄 초과 또는 루프 내 `if`/`elif` 3개 이상 시 자동 차단)

### B. MS Word 전용 (`word_*` — 15개)
- **실시간 편집 · 마크다운 · 표 · 개조식 정제 (`word_tools.py`, `component_tools.py`)**
  - `word_read_document`: 전체 본문(`all`), 선택 영역(`selection`), 표 목록(`tables`) 읽기
  - `word_write_text`: 커서 삽입 / 선택 교체 / 문서 끝 추가 및 프리셋 서식(`navy_title`, `steel_title`, `section`, `heading1`, `heading2`, `body`) 적용
  - `word_format_selection`: 선택 문단 또는 `paragraph_pattern` 정규식 대상에 글꼴/크기/색상/간격/음영/강조바 즉시 적용
  - `word_format_table`: 현재 커서 또는 지정 번호 표를 '겉바속촉' 서식(`navy`, `ib_steel_blue`, `sage_grid`, `slate_contrast`, `gray_form`, `gray`)으로 다듬기
  - `word_insert_table`: 커서 위치에 2차원 데이터 표 삽입 및 비즈니스 테마 적용
  - `word_render_markdown_report`: 마크다운 보고서를 네이비(`navy`) 또는 스틸블루(`steel_blue`) 서식으로 일괄 렌더링
  - `word_find_and_replace`: 본문 텍스트 일괄 찾기 및 바꾸기
  - `word_save_document`: 활성 Word 문서 저장 / 다른 이름으로 저장
  - `word_clean_and_structure`: 107개 기관 서식 기반 7단계 개조식 위계(`1. -> 가. -> 1) -> 가) -> (1) -> (가) -> o`) 및 자간/줄간격 일괄 교정
- **리얼 양식(SDT) · 학술 논문 · Gate 6 진정성 비교 (`word_form_and_academic_tools.py`, `word_authenticity_gate.py`)**
  - `word_insert_sdt_control`: 밑줄(`___`) 가짜 양식 대신 실제 채움형 콘텐츠 컨트롤(`text`, `dropdown`, `date`, `checkbox` 등) 삽입
  - `word_set_form_protection`: 정적 레이아웃 보호(`wdAllowOnlyFormFields=2`) 켜기/끄기
  - `word_audit_form_fields`: SDT/FormField 누락, 빈 드롭다운, 밑줄 안티패턴 및 보호 상태 감사
  - `word_apply_academic_venue_style`: 학술 논문 투고 규격(`APA 7`, `Chicago 17`, `IEEE`, `MLA 9`, `KCI`) 여백·서체·줄간격 적용
  - `word_insert_academic_caption_or_ref`: 동적 필드(`SEQ Figure/Table/Equation`, `PAGEREF`) 및 내어쓰기 참고문헌 삽입
  - `word_compare_authenticity_gate`: **Gate 6** — 기업 실무 양식/KISA 아키타입 및 원본 레퍼런스 대비 구조 유사도(%)·AI 합성 티(AI-Tell)·진위(Real vs AI) 판별

### C. MS Excel 전용 (`excel_*` — 12개)
- **실시간 읽기 · 쓰기 · 수식 · 매크로 · 데이터 정제 (`excel_tools.py`, `component_tools.py`)**
  - `excel_read_data`: 마우스 선택 영역(`use_selection=True`) 또는 지정 범위(`range_address`)를 2D 배열·Markdown 표로 읽기
  - `excel_write_data`: 지정 셀부터 2차원 배열 고속 작성 및 표 테마(`navy`, `modern_teal`, `ib_steel_blue`, `겉바속촉`) 적용
  - `excel_format_range`: 선택/지정 범위에 배경, 글꼴, 숫자 서식, 테두리 프리셋(`겉바속촉`, `겉바속점`, `ib_hair`, `all_thin`) 적용
  - `excel_auto_format_table`: Excel 시트 겉바속촉 자동 서식화(외곽 굵은선, 네이비 헤더, 천단위 콤마, AutoFit)
  - `excel_clean_selection`: 선택 영역 데이터 정제(`trim_whitespace`, `mask_rrn` 주민번호 마스킹, `calc_korean_age` 만나이 계산)
  - `excel_manage_sheets`: 시트 목록 조회, 추가, 삭제, 이름 변경, 복제, 활성화
  - `excel_evaluate_formula`: 수식 실시간 평가(`Evaluate`) 및 전체 통합문서 재계산
  - `excel_run_macro`: 활성 워크북 VBA 매크로 실행
  - `excel_save_workbook`: 활성 워크북 저장 / 다른 이름으로 저장
- **3-Zone 재무 모델 · CFO 코딩 · 적응형 대시보드 (`excel_modeling_tools.py`)**
  - `excel_scaffold_financial_model`: 3-Zone 아키텍처(Inputs 노랑 / Calc 파랑 / Outputs 초록) 기반 재무 모델(`3-statement`, `dcf`, `lbo`) 및 살아있는 수식 구축
  - `excel_apply_cfo_color_coding`: CFO 표준 5색 코딩 적용 및 Calc Zone 하드코딩 숫자 감사
  - `excel_plan_adaptive_dashboard`: 데이터 크기(`row_count`)별 대시보드 복잡도 산정 및 `###` 잘림 방지 동적 열 너비 계산·적용

### D. MS PowerPoint 전용 (`ppt_*` — 19개)
- **실시간 제어 · 분할 생성 · 도형 정밀 배치 (`ppt_tools.py`)**
  - `ppt_read_presentation`: 전체 슬라이드 목차/텍스트/표 또는 특정/선택 슬라이드 읽기
  - `ppt_create_report_deck_shell`: 단계별 분할 생성 1단계 — 기본 틀(Shell, 기본 `business_navy`) 생성 및 데스크톱 팝업
  - `ppt_update_shape_text`: 지정 슬라이드 내 도형/텍스트박스 텍스트·서식 실시간 수정
  - `ppt_find_and_replace_text`: 전체 또는 지정 슬라이드의 도형/표/그룹 내 텍스트 일괄 치환
  - `ppt_add_shapes_batch`: 단일 호출로 다수 도형·텍스트박스 좌표(pt) 기반 일괄 생성
  - `ppt_transform_shapes`: 기존 도형 위치(`Left`, `Top`)·크기·Z-Order 일괄 변경
  - `ppt_align_shapes`: 지정 도형 일괄 정렬(`Align`) 및 균등 분배(`Distribute`)
  - `ppt_save_presentation`: 활성 프레젠테이션 저장 / 다른 이름으로 저장
- **113종 슬라이스 · 레시피 · 로드맵 · 원샷 구성 (`ppt_tools.py`)**
  - `ppt_list_slice_templates`: 113종 템플릿 분석 기반 17대 역할(Role), 15대 세분 영역 슬롯, 7대 테마(`business_navy`, `steel_blue`, `modern_teal`, `clean_charcoal`, `emerald_green`, `warm_beige`, `indigo_violet`) 목록 반환
  - `ppt_insert_slice`: 외부 템플릿 없이 지정 영역에 비즈니스 슬라이스 실시간 정밀 주입
  - `ppt_insert_recipe_slice`: 카탈로그 등록 113종 디자이너 템플릿 레시피(`recipe_001~recipe_113`) 정밀 재현 삽입
  - `ppt_split_slide_layout`: 황금 분할 레이아웃(1:1, 1:2, 2:1, 4분면 등) 구성 및 각 슬롯에 슬라이스 적재
  - `ppt_compose_slide`: 상단/중단/하단 다중 슬라이스를 결합해 완성형 복합 슬라이드 1장 합성
  - `ppt_insert_column_roadmap`: 113종 분석 황금 마진 및 7대 테마 컬러 기반 모던 컬럼 로드맵 삽입
  - `ppt_auto_compose_from_text`: 비정형 원시 텍스트 기반 서사 구조·17대 슬라이스 자동 기획 및 완결형 PPT 덱 생성
- **Morph 애니메이션 · Pitch Deck (`ppt_morph_engine.py`)**
  - `ppt_clone_and_morph_slide`: 슬라이드 복제 후 Morph 전환(`ppEffectMorph=3954`) 적용 및 이전 슬라이드 `#sN-*` 도형 자동 고스팅
  - `ppt_upsert_morph_actor`: `!!scene-*`, `!!actor-*`, `#sN-*` 도형 이름 기반 위치·크기·회전·투명도·텍스트 생성/갱신
  - `ppt_verify_morph_deck`: Morph 전환 누락, 미고스팅 잔존, 고스트 과누적, 인접 슬라이드 중복, 최소 변위(5cm) 전수 검증
  - `ppt_plan_pitch_deck`: 투자 단계(`Seed`~`Series C`, `Bridge`) × 5대 버티컬 맞춤형 Pitch Deck 아크 및 레이아웃 기획

### E. 한컴오피스(HWP) 전용 (`hwp_*` — 6개)
- `hwp_read_document`: ROT(`!HwpObject`) 바인딩으로 전체 본문 또는 선택 블록(`only_selection=True`) 읽기
- `hwp_write_styled_text`: 커서 위치에 행안부/KISA 1310 공기업/비즈니스 서식 적용 텍스트 입력
- `hwp_format_selection`: 블록 선택 영역 또는 현재 줄에 표준 서식 즉시 적용
- `hwp_insert_table`: 커서 위치에 표 생성 및 2차원 배열 데이터 채우기
- `hwp_render_kisa1310_report`: KISA 1310 공문서 마크업(`제목:`, `부서:`, `네모:`, `원:`, `바:`, `별:`, `표:`) 일괄 렌더링
- `hwp_save_document`: 활성 한글 문서 저장 / 다른 이름으로 저장

### F. 오피스 공용 도구 (`office_*` — 32개, 시스템 제어 3개 제외)
- **8대 비즈니스 컴포넌트 & 일괄 조화 (`component_tools.py`)**
  - `office_insert_hierarchy` *(4대)*: 3단 이상 중첩 지원 조직도/계층 구조도
  - `office_insert_process_flow` *(4대)*: N단계 카드 + `➜` 화살표 업무 흐름도
  - `office_insert_timeline` *(4대)*: 간트형 마일스톤·기간·상태 배지 추진일정 타임라인
  - `office_insert_smart_table` *(4대)*: 상하 굵은선·헤더 이중선·오픈 테이블 겉바속촉 표
  - `office_insert_comparison` *(4대)*: AS-IS vs TO-BE 대칭 비교 카드
  - `office_insert_callout_box` *(4대)*: 좌측 2.5pt 강조선 + 파스텔 배경 콜아웃 박스(`info`/`tip`/`warning`/`quote`)
  - `office_insert_kpi_cards` *(4대)*: 수치·단위·증감율 배지(`▲ +15.4%`) KPI 카드
  - `office_insert_approval_block` *(Word·Excel 2종)*: KISA 공문서·기업 표준 결재란 생성
  - `office_harmonize_document` *(Word·Excel·PPT 3종)*: 글꼴·제목 스케일·테마 컬러·표 열 너비(`content`/`equal`/`proportional`)·역할별 정렬·얼룩무늬 일괄 통일
- **고급 네이티브 차트 & 데이터 시각화 (`chart_tools.py` — 4대 포맷 공통)**
  - `office_insert_cohort_retention`: 오피스 자체 표 API + 잔존율 비례 히트맵 색상 코호트 리텐션 표
  - `office_insert_small_multiples`: Excel 셀 내장 스파크라인 / Word·PPT·HWP 12세그먼트 비교 매트릭스
  - `office_insert_native_chart`: 오피스 자체 `AddChart2` 내장 차트 7종(`control_chart`, `contribution`, `lift_chart`, `survival_curve`, `forest_plot`, `fan_chart`, `roc_curve`)
- **283종 표준 참조 서식 RAG & 동적 빌더 (`document_forms.py`, `dynamic_form_tools.py` — 4대 포맷 공통)**
  - `office_list_forms`: Okapi BM25 기반 283종 표준 참조 서식 자연어 검색
  - `office_insert_form`: 검색된 참조 서식의 여백·테두리·병합 구조를 100% 보존한 채 데이터 주입·문구 치환
  - `office_build_custom_form`: 참조 양식의 AST 스타일을 상속받아 11종 블록(`reference_form`, `title_header`, `approval_line`, `key_value_grid`, `callout_box`, `table`, `comparison`, `checklist`, `process_flow`, `kpi_cards`, `signature`)을 조합하는 복합 서식 빌더
- **사용자 특화 템플릿 역공학 클론 (`user_template_tools.py` — 4대 포맷 공통)**
  - `office_learn_user_template`: 1건 이상의 문서를 역공학해 고정 외피·가변 슬롯 청사진으로 영구 적재
  - `office_capture_active_template`: 현재 데스크톱 활성 창 문서를 즉시 캡처 학습
  - `office_list_user_templates`: 등록된 사용자 템플릿 목록·슬롯 스키마 조회
  - `office_inspect_user_template`: 특정 템플릿 전체 구조·필수 슬롯·샘플 페이로드 조회
  - `office_generate_from_user_template`: 학습된 청사진에 새 payload를 주입해 원본과 동일한 비주얼로 실시간 COM 클론 생성
- **실무 폴더 벌크 학습 & 서식 프로필 (`folder_tools.py` — 4대 포맷 공통)**
  - `office_learn_folder`: 지정 폴더의 PPTX·XLSX·DOCX·HWP 문서를 순차 COM 스캔해 브랜드 컬러·폰트·여백 추출 및 고유 서식 신규 레시피(`custom_slice_*`) 자동 합성
  - `office_list_profiles`: 학습된 서식 프로필 목록·활성 상태·합성 도구 목록 조회
  - `office_set_active_profile`: 활성 서식 프로필 전환 및 기존 도구 기본 스타일 동기화
  - `office_show_tuning_status`: 현재 도구들에 미세조정 적용 중인 사용자 폰트·팔레트·여백 규격 확인
- **문서 구성 지능 Composition Engine (`composition_tools.py` — 4대 포맷 공통)**
  - `office_plan_composition`: 거친 원시 내용과 목적 기반 정보 아키텍처·공간 예산·린터가 반영된 구성 청사진(JSON) 기획
  - `office_render_composition`: 기획된 구성 청사진을 데스크톱 활성/신규 창에 시각 아키타입으로 실시간 렌더링
  - `office_audit_composition_budget`: HWP(A4 28줄 상한), PPT(카드 45자·Action Title), Excel(가로 12열·수식), DOCX(4대 블록) 공간 예산 및 품질 감사
  - `office_autocompose_raw_dump`: 원시 텍스트 덤프만으로 시맨틱 성분 분해 → 최적 오피스 포맷 자율 결정 → 기획·렌더링 원샷 실행
- **납품 품질 5-Gate & 토큰 병합 · 리뷰 · 시각 스냅샷 (`delivery_gate_and_merge_tools.py`)**
  - `office_run_delivery_gate` *(4대)*: 5단계 납품 감사(Gate 1 구조 / Gate 2 토큰·마크다운·수식 에러 / Gate 3 필수 요소 / Gate 4 저대비·재무 불균형·하드코딩 / Gate 5 Morph 고스트·Word 밑줄 안티패턴)
  - `office_merge_template_tokens` *(4대)*: 기존 폰트·색상·레이아웃을 보존한 `{{key}}` 무손실 일괄 치환(Mail-Merge)
  - `office_mark_review_item` *(Word·Excel·PPT 3종)*: 덮어쓰기 전 수정 제안 대상에 리뷰 메모(Comment) 및 강조 색상 비파괴 마킹
  - `office_export_visual_snapshot` *(Word·Excel·PPT 3종)*: 활성 문서 실제 렌더링 화면을 PNG/PDF 스냅샷으로 비파괴 Export

---

## 3. MCP Resources & Prompts

### Resources (3개)
- `office://status`: 4대 오피스 실행 상태 및 활성 문서 정보 JSON
- `office://guides/kisa1310-hwp`: KISA 1310 기반 한컴오피스(HWP) 공문서 작성 마크업 가이드
- `office://guides/aionui-domain-playbooks`: iOfficeAI(AionUi · AionCore · OfficeCLI) 도메인 플레이북(Morph 3-Prefix, Word SDT/학술, Excel 3-Zone/CFO 코딩, 5-Gate 납품 감사)

### Prompts (3개)
- `polish_active_selection`: 활성 창 선택 영역을 읽어 비즈니스 문체로 다듬고 지정 서식 적용
- `draft_official_hwp_report`: 한컴오피스(HWP)에 행안부/공기업 표준 공문서 서식과 표가 포함된 보고서 작성
- `create_dynamic_business_form`: 자연어 업무 서식을 283종 표준 참조 양식 기반으로 검색·커스텀 생성 및 톤앤매너 조화

---

## 4. 코딩 어시스턴트 데스크톱 앱 연동 설정 (JSON)

사용하는 AI 데스크톱 앱의 MCP 설정 파일에 아래 구성을 추가하세요. (`command`와 `cwd`는 로컬 환경 경로에 맞게 조정하세요.)
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

SSE 모드로 별도 서버를 띄워 연동할 경우:
```powershell
python -m mcp_server.server --transport sse --host 127.0.0.1 --port 8765
```

---

## 5. 자연어 질의 활용 예시

1. **열려 있는 문서 선택 영역 교정 및 톤앤매너 일괄 조화**
   > *"지금 열려 있는 워드 문서에서 내가 선택한 부분 다듬어서 네이비 섹션 제목 서식으로 바꾸고, 전체 표와 글꼴을 `office_harmonize_document`로 깔끔하게 맞춰줘."*
2. **표준 참조 양식 RAG 검색 기반 커스텀 양식 생성**
   > *"회의록에 들어갈 참석자 정보 표와 결재란, 액션 아이템 점검표가 필요해. 283종 참조 서식에서 어울리는 양식을 찾아 커스텀 양식으로 작성해줘."*
3. **PPT Morph 슬라이드 전환 & Pitch Deck 기획**
   > *"Series A B2B SaaS 투자유치 피치덱을 기획하고, 핵심 지표 슬라이드는 Morph 애니메이션(`!!actor-*`, `#sN-*` 고스팅)으로 자연스럽게 이어지게 만든 뒤 검증해줘."*
4. **Word 리얼 양식(SDT) · 학술 논문 규격 · Gate 6 진정성 검증**
   > *"계약 신청서의 밑줄(`____`) 입력란을 실제 SDT 콘텐츠 컨트롤로 바꾸고 양식 보호를 걸어줘. 그리고 원본 레퍼런스 대비 Gate 6 진정성 점수도 확인해줘."*
5. **Excel 3-Zone 재무 모델 & CFO 5색 코딩**
   > *"3-Statement 재무 모델 스켈레톤을 만들고 CFO 5색 코딩으로 감사해줘."*
6. **원시 메모 기반 자율 구성(Composition) & 납품 전 5-Gate 품질 감사**
   > *"이 회의 녹취록을 분석해서 가장 알맞은 오피스 포맷으로 보고서를 자동 구성하고, 끝나면 `office_run_delivery_gate`와 시각 스냅샷으로 깨진 곳이 없는지 검사해줘."*

---

## 6. CLI 테스트 및 데모 실행 방법

```powershell
# 1. 등록된 전체 도구(87개)/리소스(3개)/프롬프트(3개) 확인
python -m mcp_server.client_demo --list-tools

# 2. 현재 데스크톱 4대 오피스 상태 확인
python -m mcp_server.client_demo --status

# 3. 7대 비즈니스 컴포넌트 & 스마트 정제 도구 데모 실행
python -m mcp_server.client_demo --demo components

# 4. 참조 양식 기반 동적 서식 빌더 데모 실행
python -m mcp_server.client_demo --demo dynamic

# 5. 표준 오피스(PPT/Excel/Word) + 실시간 COM 읽기/쓰기 통합 데모 실행
python -m mcp_server.client_demo --demo all

# (옵션) 데스크톱 창을 띄우지 않고 파일 생성만 검증할 때
python -m mcp_server.client_demo --demo live-com --no-open-desktop
```
