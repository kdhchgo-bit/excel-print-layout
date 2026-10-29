# Progress ledger usage / 진척 장부 사용법

The ledger is **one JSON file per workbook**, with one writer. It records claims and hashes; it does not edit a workbook or independently validate visual evidence. All paths and evidence remain local unless explicitly shared. 장부는 워크북별 1개이며 단일 조정자가 작성한다. 장부는 검증 기록을 관리하지만 엑셀을 수정하거나 육안 검사를 대신하지 않는다.

## Example / 예시

The following is a command sequence, not a substitute for the actual checks. Use real workbook paths, complete sheet inventory, a truthful normalized request and actual evidence. `init` reads the sheet catalog from OOXML or, for `.xls`, from `xlrd`; every cataloged sheet must be classified before finalization. Native Excel must also check any legacy chart or macro sheets that `xlrd` may omit. For encrypted or unsupported formats, supply a native-Excel-generated `--catalog` JSON shaped as `{"sheets": [{"name": "Sheet1", "visibility": "visible"}]}`. Classifications and the catalog still require genuine inventory evidence. 아래 예시의 근거를 그대로 통과 처리하지 말고 실제 경로와 검토 내용을 기록한다. 원본 시트 목록은 OOXML 또는 `.xls`의 `xlrd`에서 읽으며, 장부에 모든 시트를 분류해야 한다. `.xls` 차트/매크로 시트는 Excel로 추가 확인한다. 자동 조사가 불가능하면 Excel에서 만든 JSON을 `--catalog`로 제공한다.

```powershell
# Fingerprint the current scope/constraints. Do not store the request text in the ledger.
$request = 'A4; selected files only; keep hidden; preserve originals'
$hasher = [System.Security.Cryptography.SHA256]::Create()
$requestHash = -join ($hasher.ComputeHash([Text.Encoding]::UTF8.GetBytes($request)) | ForEach-Object { $_.ToString('x2') })

python scripts/progress_ledger.py init --ledger work/one-ledger.json --source original.xlsx --output revised.xlsx --request-fingerprint $requestHash
python scripts/progress_ledger.py add-sheet --ledger work/one-ledger.json --sheet 'Sheet1' --scope printable
python scripts/progress_ledger.py add-sheet --ledger work/one-ledger.json --sheet 'HiddenCalc' --scope hidden

python scripts/progress_ledger.py record --ledger work/one-ledger.json --gate inventory --state passed --evidence 'Complete sheet inventory' --evidence-file work/inventory.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --gate baseline --state passed --evidence 'Source baseline' --evidence-file work/baseline.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --sheet 'Sheet1' --gate plan --state passed --evidence 'Final target layout' --evidence-file work/plan.json
```

Create `work/plan.json` using the shape in [print-plan.example.json](../assets/print-plan.example.json). Replace the example with **the current source file's SHA-256**, `$requestHash`, every printable sheet's actual name and the exact planned print properties. The example's placeholder digest and page settings are not defaults. `plan.json`에 원본 SHA-256, 요청 지문, 출력 대상 시트 전부의 실제 목표 설정을 기입한다.

After saving the **final** revised workbook, perform the checks and record the output-bound gates:

```powershell
python scripts/progress_ledger.py record --ledger work/one-ledger.json --gate reopen --state passed --evidence 'Saved workbook reopened' --evidence-file work/reopen.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --gate content --state passed --evidence 'Structure and formula comparison' --evidence-file work/structure.json --evidence-file work/excel.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --gate links --state passed --evidence 'External link targets checked' --evidence-file work/excel.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --sheet 'Sheet1' --gate pdf --state passed --evidence 'Sheet PDF exported and checked' --evidence-file work/pdfs/Sheet1.pdf --evidence-file work/pdf-audit.json
python scripts/progress_ledger.py record --ledger work/one-ledger.json --sheet 'Sheet1' --gate visual --state passed --evidence 'Rendered page reviewed' --evidence-file work/page-images/Sheet1-p001.png
python scripts/progress_ledger.py finalize --ledger work/one-ledger.json --plan work/plan.json
python scripts/progress_ledger.py status --ledger work/one-ledger.json --request-fingerprint $requestHash
```

Every `passed` gate requires at least one existing `--evidence-file`; its SHA-256 is recorded. The `plan` gate for every printable sheet must include the exact final plan file. If there are no external links, `links --state not_applicable` is allowed **only after links were checked**; give the reason in `--evidence`. A failed gate can be recorded as `--state failed`. `status` requires the current request fingerprint; if it differs, `verified` is not returned. A source, output, plan or evidence-file change also makes a finalized ledger `stale`. 각 `passed` 관문은 존재하는 `--evidence-file`이 하나 이상 필요하며 해시가 기록된다. 외부 링크가 없음을 확인한 경우에만 `links`에 `not_applicable`을 사용할 수 있다. 현재 요청 지문이 다르거나 원본·결과·계획·증거 파일이 변경되면 `verified`가 아니다.

Record `reopen`, `content`, `links`, `pdf` and `visual` only after the workbook state supporting the evidence has been saved. These gates store the current **output SHA-256**. If the workbook changes after one is recorded, `finalize` rejects that gate until its evidence is reviewed against the new output and recorded again. `plan`, `inventory` and `baseline` are source/planning gates. Adding a sheet invalidates the inventory gate. After finalization, running the same `init` or `finalize` again is a no-op if the bound state is identical.

Before deleting a temporary evidence folder, retain the ledger, final plan and evidence at a user-approved durable location if future no-op reuse is desired. If evidence is removed, `status` correctly becomes stale and a later agent must revalidate rather than claim a verified no-op. 재실행 시 검증된 결과를 그대로 재사용하려면 장부·최종 계획·증거를 사용자가 승인한 위치에 보관한다. 증거를 삭제했다면 재검증해야 한다.

`status` also reports passed/required gate counts and their percentage after all source-catalog sheets are classified. This is **gate coverage**, not a time estimate or a measure of quality. For this tool, exit code `0` means the command succeeded (or `status` is verified), `1` means `status` is incomplete/stale, and `2` means invalid input or an error. A `passed` claim is only as good as the evidence supplied. 상태의 퍼센트는 전체 시트 분류 뒤의 관문 통과 비율이며 소요 시간이나 품질 예측이 아니다. 검사 도구의 성공, PDF 렌더, 셀 내부의 글자와 선에 대한 시각 확인은 서로 다른 관문이다.
