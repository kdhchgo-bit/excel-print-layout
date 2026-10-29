# Excel Print Layout Skill

엑셀 인쇄 레이아웃을 정리하고, 원본과 숨김 상태를 보존하면서 실제 출력 결과까지 검증하는 AI 에이전트용 스킬입니다. A skill for preparing readable Excel print layouts and verifying the result without losing workbook content or hidden state.

**[한국어](#한국어) · [English](#english)**

## 한국어

### 무엇을 하나요?

- 요청한 파일과 제외 대상을 구분하고, 원본을 보존한 수정본을 만듭니다.
- A4 가로·세로 방향, 인쇄영역, 여백, 배율, 반복 머리글과 페이지 구분을 시트별로 결정합니다.
- 숨김 행·열·시트, 병합 셀, 수식·계산 캐시, 외부 연결을 유지하고 원본과 대조합니다.
- 표가 작으면 읽기 좋게 확대하고, 지나치게 넓거나 긴 표는 머리글·식별 열·테두리를 유지하는 논리적 단위로 나눕니다.
- 실제 Excel 저장본을 다시 열고 PDF를 출력해 `###`, 잘린 문자, 표 경계, 빈 쪽과 여백을 확인합니다.
- 작업 계획과 원본·결과 해시, 파일별·시트별 검증 관문을 기록해 같은 작업을 반복할 때 변경이 누적되지 않게 돕습니다.

이번 방식은 여러 종류의 산출표와 도면이 섞인 워크북 정리 경험을 바탕으로 구성했지만, 특정 사업의 값·크기·파일 경로는 기본값으로 사용하지 않습니다. 상세한 **19개 작업 계층, 작업별 관문, 진척 상태, 실패 복구 및 멱등성 규칙**은 [한국어 상세 명세](references/architecture.ko.md)를 참고하세요.

### 설치와 사용

Codex 스킬을 자동 검색하는 Windows 환경에서 다음과 같이 설치할 수 있습니다. 이미 같은 이름의 스킬이 있으면 먼저 기존 내용을 확인하세요.

```powershell
git clone https://github.com/kdhchgo-bit/excel-print-layout.git "$env:USERPROFILE\.codex\skills\excel-print-layout"
```

새 대화에서 아래처럼 요청하거나 `$excel-print-layout`을 직접 지정합니다.

> `$excel-print-layout`을 사용해 지정한 엑셀 파일을 A4로 인쇄하기 좋게 정리해 줘. 숨김은 유지하고 원본은 보존해 줘.

다른 AI 에이전트에서도 `SKILL.md`와 참조 파일을 읽을 수 있다면 동일한 절차를 사용할 수 있습니다. 에이전트마다 스킬 설치 위치와 자동 검색 방식은 다릅니다.

### 실행 환경과 도구

실제 Excel 인쇄 설정 수정·저장·재열기에는 Windows와 Microsoft Excel이 필요합니다. 보조 검증 도구는 Python 3.10 이상을 사용하며 필요한 패키지는 기능별로 다릅니다.

```powershell
python -m pip install pywin32 xlrd openpyxl PyMuPDF Pillow
```

| 도구 | 역할 | 확인하지 못하는 것 |
|---|---|---|
| `scripts/compare_workbooks.py` | `.xls`/`.xlsx` 등의 숨김 범위·병합·값 및 지원되는 수식/캐시 비교 | `.xls` 수식 원문, 외부 연결 대상, 출력상 잘림 |
| `scripts/compare_excel_com.py` | 별도 Excel 인스턴스에서 수식·값·외부 링크·시트 상태 비교 | 숨김 범위·병합의 독립 비교와 시각적 품질 |
| `scripts/audit_pdf.py` | A4 크기·빈 페이지·`###`·문자 크기·대략적 점유율 검사 및 이미지 생성 | 셀 안의 반쪽 글자, 모든 그림과 테두리의 정확성 |
| `scripts/progress_ledger.py` | 파일/시트별 관문과 원본·계획·결과 해시 관리 | 실제 검토의 진실성이나 엑셀 파일 편집 |

```powershell
python scripts/compare_workbooks.py original.xlsx revised.xlsx --output work/structure.json
python scripts/compare_excel_com.py original.xls revised.xls --output work/excel.json
python scripts/audit_pdf.py work/pdfs --output work/pdf-audit.json --render-dir work/page-images
```

PDF 자동 검사와 실제 이미지 검토를 함께 수행해야 합니다. 보고서나 렌더 이미지가 이미 존재하고 내용이 다르면 기본적으로 덮어쓰지 않으며, 의도적으로 교체할 때만 `--overwrite`를 사용합니다. 도구의 출력 경로로 입력 엑셀이나 PDF를 지정할 수 없습니다. 수식 연결 경로도 결과 파일을 옮긴 환경에서 다시 확인해야 하며, PDF 확인은 종이 실인쇄 확인과 다릅니다.

### 진척과 재실행

파일별 장부는 `init → add-sheet → record → finalize → status` 순서로 사용합니다. `record`에는 실제 증거 파일을 연결하고, `finalize`는 원본·요청 지문이 일치하는 인쇄 계획 JSON과 필수 관문을 확인합니다. 출력 관련 증거는 기록 당시의 결과 파일 해시에 연결됩니다. 결과나 증거를 나중에 바꾸면 기존 검증은 무효가 됩니다. 장부는 통합문서 하나당 하나로 만들고 단일 조정자가 기록합니다. [실행 예시](references/ledger-usage.md)를 참고하세요.

요청·원본·계획·결과가 모두 같은 `verified` 상태라면 다시 Excel에 저장하지 않고 결과를 재사용합니다. 이 방식은 **누적 변경을 방지하는 절차**이며, AI의 판단이나 시각 검토가 자동으로 완벽해진다는 보장은 아닙니다.

### 결과 해석과 제한

- 구조 비교 성공은 검사한 필드의 일치만 의미하며, 사진·도면·실제 종이 출력까지 보증하지 않습니다.
- 그림에 포함된 작은 치수는 셀 글꼴을 키워도 커지지 않을 수 있습니다. 읽기 어려운 글자가 남으면 구체적으로 보고합니다.
- 사용자의 Excel이나 프린터 상태에 따라 실제 페이지 경계가 달라질 수 있습니다.
- 인쇄 작업 중 발견한 원본의 수량·규격 불일치는 임의 수정하지 않고 별도 확인사항으로 알립니다.
- Excel COM 오류나 누락된 의존성이 있으면 작업을 중단하고 확인하지 못한 범위를 명시합니다.

## English

### What it does

- Establishes the requested scope and exclusions, then works on named copies of the original workbooks.
- Selects paper orientation, print areas, margins, zoom, repeat titles and page breaks per sheet.
- Preserves and compares hidden rows, columns and sheets, merged cells, formulas, cached values and external-link targets.
- Enlarges small tables when practical; splits very wide or long tables only at coherent boundaries with repeated identifiers, headers and complete borders.
- Reopens saved workbooks and exports PDF evidence to inspect clipped text, `###`, borders, blank pages and whitespace.
- Tracks a source-bound plan and workbook/sheet verification gates so repeated runs do not accumulate width changes, breaks or repair lines.

The skill is based on practical mixed-table and drawing workflows but does not hard-code any project dimensions, values or local paths. See the [detailed English architecture](references/architecture.en.md) for the **19 layers, tasks, evidence gates, progress, recovery and idempotence mechanisms**.

### Installation and use

For Codex on Windows, clone into your personal skills directory. Inspect an existing directory of the same name before installing.

```powershell
git clone https://github.com/kdhchgo-bit/excel-print-layout.git "$env:USERPROFILE\.codex\skills\excel-print-layout"
```

In a new conversation, ask the agent to use the skill explicitly:

> Use `$excel-print-layout` to prepare the specified workbooks for readable A4 printing. Preserve hidden items and the originals.

Other AI agents can follow `SKILL.md` and its referenced files if their environment supports local skill instructions. Discovery and installation paths vary by agent.

### Requirements and helpers

Editing, saving and reopening in native Excel requires Windows and Microsoft Excel. The helpers use Python 3.10+; install the dependencies needed for your chosen checks:

```powershell
python -m pip install pywin32 xlrd openpyxl PyMuPDF Pillow
```

| Helper | Purpose | Important coverage limit |
|---|---|---|
| `scripts/compare_workbooks.py` | Read-only hidden-state, merge, value and supported formula/cache comparison | BIFF formula expressions, actual link targets and visual clipping |
| `scripts/compare_excel_com.py` | Compare formulas, values, links and sheet state in a private Excel instance | Independent hidden/merge comparison and visual quality |
| `scripts/audit_pdf.py` | Check A4 size, blank pages, `###`, apparent font sizes and ink coverage; render pages | Complete detection of clipped glyphs, drawings and border defects |
| `scripts/progress_ledger.py` | Track workbook/sheet gates and source/plan/output hashes | Proving a reviewer actually inspected a page or editing a workbook |

```powershell
python scripts/compare_workbooks.py original.xlsx revised.xlsx --output work/structure.json
python scripts/compare_excel_com.py original.xls revised.xls --output work/excel.json
python scripts/audit_pdf.py work/pdfs --output work/pdf-audit.json --render-dir work/page-images
```

Automated PDF checks are triage aids; inspect rendered pages. Existing reports and renders with different contents are not overwritten unless `--overwrite` is explicit, and helper outputs cannot replace an input workbook or PDF. Verify external links after moving a workbook. A PDF export is not a physical print test.

### Progress and repeatability

Use one single-writer ledger per workbook: `init → add-sheet → record → finalize → status`. Each passed gate needs a real evidence file. `finalize` checks the source/request-bound JSON plan and required gates; output-dependent evidence is bound to the output hash when recorded. A later output or evidence change invalidates prior evidence. See the [command example](references/ledger-usage.md).

When the same request, source, plan and output are still `verified`, reuse the existing result without resaving it. This is a **process for preventing cumulative changes**, not a guarantee that an AI's judgment or visual review is flawless.

### Limits

- A structural match means only that the fields actually checked match; it does not guarantee the appearance of images or a physical print.
- Text embedded in a drawing may remain small despite cell font changes. Report concrete limitations rather than forcing an unsafe enlargement.
- Actual printer margins and drivers can change pagination.
- Report pre-existing quantity/specification discrepancies separately and do not silently alter engineering data during print formatting.
- If Excel COM or a required dependency is unavailable, report which checks could not be performed.

## Repository layout

```text
SKILL.md                         Agent entrypoint
agents/openai.yaml               Codex display metadata
references/architecture.ko.md    Detailed Korean workflow and gates
references/architecture.en.md    Detailed English workflow and gates
assets/print-plan.example.json   Plan shape; values are examples only
scripts/                         Read-only verification and progress helpers
tests/                           Synthetic regression checks
```

Run the regression checks with `python -m unittest discover -s tests -v`. No source workbooks, private paths or project data are included in this repository.
