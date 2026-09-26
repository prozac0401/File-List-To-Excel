# File Collect 구현·납품 보고 — v1.2 평가 후보

기록일: 2026-09-24. [작업지시서](../FileListToExcel_FileCollect_WorkOrder.md)에 따라 기존 File List to Excel에 **Excel 목록에서 선택한 실제 파일의 복사본을 새 폴더에 모으는 기능**을 통합했다. Excel 통합문서의 Ctrl+S를 대신하는 기능이 아니다.

**구현·평가 MSI는 완료했고 실제 Excel 자동 로드·우클릭 복사·결과 보기와 결과창 닫기는 통과했지만, 전체 출시 게이트는 실패했다.** M6의 OFF·ON·실제 1.1→1.2 업그레이드를 모두 실행했으며 세 경로의 전체 종료 코드는 모두 1이다. 일반 파일 메뉴 등록 누락과 ON 제거 후 시험 값 잔류는 실패로, 비활성 상태 보존은 미확정으로 남았다. 원인은 확정하지 않았다. 최종 시험 Excel은 Quit 반환·창 닫힘 후 상당한 지연 뒤 최종 부재를 확인했으며 지연 원인은 미확정이다.

## 1. 구현과 재사용

| 영역 | 수정/추가 위치 | 동작 및 기존 코드 재사용 |
| --- | --- | --- |
| Excel 추가 기능 | [src/FileListToExcel.ExcelAddIn](../src/FileListToExcel.ExcelAddIn/) | 네이티브 x86/x64 IDTExtensibility2 COM 추가 기능. 현재 Excel 인스턴스의 메뉴·이벤트·메모리 선택을 사용하고 컴파일된 콜백으로 helper를 호출 |
| 목록/중복 workbook | [WorkbookWriter](../src/FileListToExcel.Core/WorkbookWriter.cs), [DuplicateWorkbookWriter](../src/FileListToExcel.Core/DuplicateWorkbookWriter.cs), [FileCollectContract](../src/FileListToExcel.Core/FileCollectContract.cs) | 기존 직접 ZIP/XML 생성, 셀 형식·문자열 escaping·하이퍼링크·Table·시트 분할을 재사용. 같은 Table 안에 숨김 행 정보를 추가 |
| 계획/복사/보고서 | [CopyPlanner](../src/FileListToExcel.Core/CopyPlanner.cs), [CopyNative](../src/FileListToExcel.Core/CopyNative.cs), [CopyService](../src/FileListToExcel.Core/CopyService.cs), [CollectModels](../src/FileListToExcel.Core/CollectModels.cs) | 기존 cloud/reparse 정책 재사용. Windows 핸들 검증, CopyFileExW, no-overwrite, 충돌 이름, 취소와 행별 결과 |
| 기존 helper 연결 | [CollectRequestReader](../src/FileListToExcel.App/CollectRequestReader.cs), [CollectFolderPicker](../src/FileListToExcel.App/CollectFolderPicker.cs), [CollectWindow](../src/FileListToExcel.App/CollectWindow.cs), 기존 CommandLine/Program | 기존 helper에 별도 --collect-request 동작 추가. 소유 GUID 요청을 한 번 소비하고 목적지·진행·완료 화면 제공 |
| 설치/개발 검증 | [Product.wxs](../installer/Product.wxs), [설치 안내](../installer/README.md), [Build](../scripts/Build.ps1), [Test-Installer](../scripts/Test-Installer.ps1), [Test-Upgrade](../scripts/Test-Upgrade.ps1), [진단](../scripts/Get-ExcelIntegrationStatus.ps1) | 기존 UpgradeCode·설치 경로·탐색기 등록 유지. 선택형 Excel 구성 요소와 x86/x64 COM 뷰, 기본 OFF·복구·제거 회귀 추가 |
| 시험 | [계약 시험](../tests/FileListToExcel.Core.Tests/FileCollectContractTests.cs), [복사 시험](../tests/FileListToExcel.Core.Tests/CollectCopyTests.cs), [요청 시험](../tests/FileListToExcel.App.Tests/CollectRequestTests.cs), [실제 Excel 시험 도구](../tests/FileListToExcel.ExcelAcceptance/) | 기존 목록/중복 회귀에 추가. 테스트 도구는 개발용이며 최종 사용자 배포 절차가 아님 |

기존 visible 열과 FileListTableN 이름을 유지했다. Files 계열은 기술 열 2개를 추가하여 14열, Duplicates는 13열, Matches는 10열 Table이 되며 새 열은 숨겨진다. Summary/Errors/Skipped는 복사 가능한 표로 등록하지 않는다. 기존 결과 파일은 계속 조회할 수 있지만 새 기능에는 새로 생성한 스키마 1 결과가 필요하다.

일반 목록/중복 찾기 본체는 추가 기능 없이 동작한다. 명단대조기와 Workspace의 선택범위 내보내기에는 의존하지 않는다. 기존의 “Excel Add-in 설치 금지” 조건을 이번 선택 기능에 한해 변경한 결정은 [ADR 0003](adr/0003-excel-file-collect.md)에 기록했다.

## 2. 자동 로드와 설치 소유권

설치 화면에서 **Excel에서 선택한 파일 복사**를 한 번 선택하면 MSI가 제품 COM 등록과 최초 정상 시작 로드 설정을 담당한다. 구성 요소는 기본 OFF이며 기존 사용자의 무인 업그레이드에서 새로 켜지지 않도록 설계했다. 이미 열린 Excel은 사용자가 작업을 저장하고 정상적으로 닫은 뒤 다시 시작한다.

이후 사용자가 하는 일은 결과표의 행을 선택하고 **파일목록 → 선택한 파일 복사…**를 누른 다음 목적지 화면에서 **여기에 복사**를 선택하는 것이다. 자동 동작은 메뉴·이벤트 준비다. 통합문서 열기·선택 변경만으로 파일 검증이나 복사를 시작하지 않는다.

추가 기능 ProgID/CLSID와 메뉴 Tag는 탐색기 확장 및 다른 제품과 분리한다. 종료 시 소유 컨트롤과 이벤트만 해제하며 CommandBars.Reset을 호출하지 않는다. VBA/VBS·매크로 허용·추가 기능 수동 등록, 광범위한 신뢰 위치, 정책 완화, 임의 인증서 신뢰 등록을 요구하지 않는다. 사용자/Office가 비활성화한 추가 기능을 감시하며 강제로 켜는 프로세스도 없다.

본체는 자체 포함 Windows x64 helper이고 추가 기능은 정적 C++ 런타임을 사용한다. 최종 사용자에게 Python, 개발 SDK, VSTO 런타임 설치나 스크립트 실행 정책 변경을 요구하지 않는다. 조직의 추가 기능·서명 정책은 적용되며 MSI 성공만으로 어떤 환경에서나 Excel 자동 로드가 보장되지는 않는다.

## 3. 표와 행의 대응

통합문서에 FLT.Product=FileListToExcel, FLT.ActionSchemaVersion=1, FLT.WorkbookId를 쓰고, FLT.Table.<기존 표 이름> 속성으로 복사 가능한 표를 등록한다. 같은 Table의 __FLT_ItemId와 __FLT_SourceRecord에 행 UUID 및 kind/정확한 경로/크기/UTC 수정 ticks를 기록한다. 큰 정수는 JSON 문자열이다. JSON이 셀 한도를 넘으면 자르지 않고 unavailable과 오류 기록으로 해당 행의 복사 불가를 명시한다.

추가 기능은 현재 Excel 메모리에서 선택과 DataBodyRange의 교차를 구하고, 숨김·필터 행을 제외하며 같은 행을 여러 셀에서 선택해도 한 번만 요청한다. 필드 이름으로 열을 찾고 같은 Table의 행 정보를 함께 읽는다. 정상 Table 정렬·열 이동·시트 이름 변경·Save As 후에도 이 대응을 유지한다. 마지막 저장된 .xlsx를 다시 읽어 현재 상태를 대신하지 않는다.

**표식은 출처 인증이나 파일 접근 권한이 아니다.** helper는 UUID·버전·JSON·표시 이름/경로를 확인하고, 복사 엔진은 실제 파일/경로/메타정보를 다시 검증한다. 일관되게 위조된 기록도 이 파일시스템 검사를 통과해야 한다. 크기와 수정시간 일치는 생성 당시 바이트의 증명이 아니며 과거 파일 버전 복원 기능도 아니다. 자세한 계약은 [스키마 문서](FILE_COLLECT_SCHEMA.md)에 있다.

## 4. 원본 보호와 개인정보

- 원본은 읽기/복사만 한다. 새 작업 하위 폴더에 평평하게 모으고 동일 이름은 복사본 이름만 구분한다. OS 복사 호출도 덮어쓰기를 금지한다.
- source 및 조상 경로의 reparse/recall 상태, 현재 크기·UTC 수정시간·가능한 파일 식별자를 확인한다. 삭제·변경·접근 거부는 행별 결과에 남긴다. cloud-only를 일부러 내려받거나 EFS/DRM을 우회하지 않는다.
- 기본 복사는 직렬이며 취소 후 새 파일을 시작하지 않는다. 완료본은 유지하고 자신의 미완료 출력만 안전하게 정리한다. 네트워크/OS 호출의 취소 지연은 안내한다.
- 성공·실패·제외·취소·중복 경로 결과와 원본/목적지 대응은 제품 LocalAppData의 Reports에 기록한다. 전달 폴더에 내부 원본 경로 보고서를 자동으로 넣지 않는다. 외부 전송은 없다.
- 결과 폴더 열기는 사용자의 명시적 동작이며 복사된 문서/실행파일을 자동으로 열지 않는다.

원본 hardlink는 일반 파일로 허용한다. 요청 파일 hardlink는 거절한다. 모든 reparse 경유 경로를 보수적으로 제외하므로 로컬에 있는 cloud 파일도 제외될 수 있다. 10,000 보이는 행/32 MiB 한도는 “고유 파일 10,000개”보다 엄격하며, 극단적으로 긴 확장자의 충돌 이름은 제외할 수 있다.

## 5. 실행한 검증과 현재 판단

마지막 관리 코드 전체 실행은 **Core 138 + App 47 = 185개 통과, 실패/건너뜀 0, 종료 코드 0**이다. 결과 grid 문자열 형식 수정과 실제 숨김 STA DataGridView의 FormattedValue/DataError 회귀를 포함하여 전체 솔루션을 Release로 시험했다. [최종 요약](../artifacts/validation/collect/final-185-summary.md), [명령 출력](../artifacts/validation/collect/final-185-output.txt), [TRX](../artifacts/validation/collect/final-185/), [grid 수정 전후 증거](../artifacts/validation/collect/report-grid-regression-summary.md)에 기록했다. 이 결과는 최종 MSI의 실제 UI 수용을 대신하지 않는다.

이전 실행 이력은 기준 Core 90/App 31, 중간 전체 Core 135/App 44, 후속 복사 집중 35개(실제 ACL 시험 3개 포함), App 집중 46개, grid 수정 전 전체 184개다. 이들을 최종 185개에 더하지 않는다. 개발 네이티브 셸은 RequestSink dispatch 포함 414개 검사, 설치된 실제 DLL의 표준 시험은 253개를 통과했다. RequestSink가 없는 installed 실제 helper에 개발 전용 --dispatch를 적용한 실패 실행은 올바른 installed 시험이 아니며 두 개수를 혼합하지 않는다. 추가 기능 x86/x64는 각각 COM/요청 시험을 통과했다.

설치된 M5 시험 MSI와 실제 Excel x64 16.0.20326.20158에서 다음을 확인했다.

- VBA action 없이 MSI 등록 후 새 Excel 프로세스의 자동 연결·메뉴 준비, 이전 시험 Excel 종료 후 다른 프로세스의 재연결.
- 100행에서 필터로 보이는 7개만 실제 복사, 각 SHA-256 동일, 전달 폴더 내부 보고서 없음.
- 한 행 수동 숨김과 겹치는 두 열 선택에서 6개로 계산, 목적지 취소 기록.
- 내림차순 정렬·경로 열 이동·시트 이름 변경·Save As 후 전체 행 2:3 선택으로 f099/f100만 복사.
- 관계없는 셀이 섞인 선택의 실제 거절과 helper 미실행. 기술 열 삭제 후 명령 클릭도 실제 거절.
- 정상적인 시험 통합문서 Close/Quit 후 해당 Excel 프로세스 및 helper가 남지 않음. M5 시험 설치도 [제거 로그](../artifacts/file-collect-live/m5-uninstall.log)에서 종료 코드 0으로 정상 제거됐으며 최종 후보 M6 시험과 구분함.
- 세 작업의 스키마 1 내부 보고서에서 Copied 7 / Cancelled 6 / Copied 2를 확인했고, 모두 전달 폴더 밖에 보관됨.

SDI 수정 MSI(ProductCode {ABF0839A-8787-4D9A-9912-A4BD8E41CE50}, SHA-256 c89582f4837dbf297a7649289943f83d42dba4c128bac4a1822b910f28486801)의 후속 실제 시험에서는 다음을 확인했다.

- 서로 다른 Excel 인스턴스 2개 모두 자동 연결, EnableEvents=true, Cell/List Range Popup의 제품 메뉴 각각 1개.
- 일반 workbook, 미래 버전 999, 기술 열 삭제, Summary에서 비활성. 지원 workbook로 돌아가거나 스키마 1로 복구하면 활성.
- 미저장 필터 7개 실제 복사와 fixture 일치, 새 창 Duplicates 3개 복사, 다른 인스턴스 Matches 2개 복사 및 보고서 저장.

따라서 이전 SDI 메뉴 상태 문제는 실제 Excel 재시험에서 해결을 확인했다. [SDI-A](../artifacts/file-collect-live/sdi-a/), [SDI-B](../artifacts/file-collect-live/sdi-b/), [7개 복사 검증](../artifacts/file-collect-live/sdi-a/final-copy-verification.json)에 증거를 남겼다.

그러나 Matches 복사 완료 후 **작업 결과 보기에서 정수 RowIndex에 대한 DataGridView FormattingApplied 처리로 반복 FormatException**이 발생했다. 정상 UI 종료 시도로 해소되지 않아 복사 완료된 작업 소유 helper PID 7264만 종료했고, 시험 Excel 두 인스턴스의 Close/Quit 호출은 정상 반환했다. Excel 프로세스를 강제 종료하지 않았다. [실제 결함 기록](../artifacts/file-collect-live/sdi-b/result-ui-regression.txt)을 보존하며, 결과 UI 수정·새 MSI 빌드·185개 자동 회귀를 완료했고 후속 최종 MSI의 실제 결과 보기/닫기 회귀도 통과했다. c895… 후보는 전체 UI 합격이나 최종 배포 해시로 사용하지 않는다. M6 세 경로의 실행은 완료했으며 아래 등록/제거 실패와 비활성 상태 보존 미확정이 남았다. 상세 상태는 [검증 기록](FILE_COLLECT_VALIDATION.md)에 있다.

최종 UI 수정 MSI(233a93be…)에서는 새 Excel x64 프로세스에서 자동 로드와 양쪽 컨텍스트 메뉴의 제품 메뉴 각각 1개를 확인했다. Matches A6:A7을 선택하고 **실제 셀 우클릭 → 파일목록 → 선택한 파일 복사**를 클릭했다. 목적지 화면은 2행/44 bytes를 표시했고 “여기에 복사” 후 두 파일이 원본 fixture의 SHA-256과 같았다. 결과창의 2행에는 순번 1/2와 한국어 “복사 완료”가 표시됐으며 예외 없이 열리고 정상 닫혔다. [시작 증거](../artifacts/file-collect-live/ui-final/started.json), [복사·닫기 검증](../artifacts/file-collect-live/ui-final/copy-verification.json), [결과창 접근성 기록](../artifacts/file-collect-live/ui-final/result-grid-accessibility.txt)에 기록했다.

완료 helper는 Alt+F4로 정상 종료했다. 시험 Excel은 Quit 반환·창 닫힘 후 상당한 지연 뒤 최종 부재를 확인했다. 시험 harness와 helper도 최종 조회에서 없었으며 Excel을 강제 종료하지 않았다. 다른 Excel automation과 동시 상태였고 지연 원인은 확정하지 않았다. [중간 잔류 기록](../artifacts/file-collect-live/ui-final/process-exit-pending.txt)과 [최종 부재 확인](../artifacts/file-collect-regressions/final-machine-state.json)을 함께 보존하며 결과창 닫기와 Excel 프로세스 소멸 시점을 구분한다.

M6의 strict 본체만 설치 시험은 두 번 실패했다. HKCU\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel 등록은 MSI 테이블/로그에 존재하지만, PowerShell과 reg.exe 양 비트수의 HKCU/HKCR/HKLM 및 직접 사용자 Classes 조회에서 키가 없었다. Directory/Background 키는 정상이다. [직접 비교 증거](../artifacts/file-collect-regressions/m6-registry-provider-diagnosis.txt)와 [독립 레지스트리 증거](../artifacts/file-collect-regressions/m6-independent-registry-views.json)를 보존했다. 과거 유사한 로컬 기록이 있어도 이번 원인을 환경이나 provider 문제로 단정하지 않는다.

별도 [설치된 기능 시험](../artifacts/file-collect-regressions/m6-functional-output.txt)은 목록·중복·Matches·선택 파일·SQLite 재사용·원본 유지·helper 종료를 통과했고, [설치된 네이티브 표준 시험](../artifacts/file-collect-regressions/m6-installed-native.json)도 253개를 통과했다. 이 결과들은 실제 메뉴 등록 조건 실패를 상쇄하지 않는다. OFF의 설치·기능·복구·제거 전 단계는 완료했다. Windows Installer의 설치/복구/제거 각각은 종료 코드 0이었고, 설치된 네이티브 표준 253개, 목록·중복·Matches·선택 파일·캐시·원본/통합문서 보존·helper 종료와 선택형 Excel 기능 부재를 확인했다. 그러나 설치/복구의 일반 파일 메뉴 등록 실패 2건을 누적하여 시험 전체는 **종료 코드 1**이다. [OFF 최종 출력](../artifacts/file-collect-regressions/m6-off-final-output.txt), [MSI 로그와 시험 산출물](../artifacts/installer-smoke/59e06e4a1f464ba08dbbf3f3ea32b4d1/)에 기록했다.

시험 도구의 기존 네이티브 `&` 호출에서 이 호스트의 LASTEXITCODE가 설정되지 않는 문제는 `Start-Process -Wait -PassThru`의 명시적인 ExitCode 확인으로 수정했다. 수정 후 253개 검사를 정상 확인했으며, 이는 앞선 개발 전용 `--dispatch`의 installed 호출 오류와 별개이고 메뉴 등록 누락의 해결도 아니다. 진단 계속 옵션은 누적 실패를 최종 종료 코드 1로 유지하며 정식 게이트를 green으로 바꾸지 않는다.

ON의 설치·기능·복구·제거도 완료했고 각각의 MSI 호출은 0, 기능·네이티브 253개 시험은 통과했다. 설치/복구에서 일반 파일 메뉴 키 누락 경고도 지속됐다. 최초 Excel 32/64비트 등록과 LoadBehavior=3을 확인했다. 시험에서 2로 바꾼 LoadBehavior를 복구 후 PowerShell에서 2로 읽었지만 같은 MSI 복구 로그는 AppSearch와 쓰기 값 #3을 기록하므로 **비활성 상태 보존 검증은 미확정**이다. 제거 뒤 시험이 만든 DWORD 2만 남아 등록 부재 단언이 실패하여 전체 종료 코드는 1이다. 이 값만 존재하고 하위 키·제품 설치 폴더가 없음을 확인한 뒤 시험 값을 직접 정리했으며, 이를 제거 합격으로 바꾸지 않았다. [ON 실행 출력](../artifacts/file-collect-regressions/m6-on-output.txt), [시험 값 정리](../artifacts/file-collect-regressions/m6-test-value-cleanup.txt)에 기록했다.

실제 1.1.0→1.2.0 업그레이드는 ProductCode 교체, Excel 기능 자동 추가 방지, 기존 기능·캐시·원본/통합문서 보존·제거 확인을 통과했다. 각 MSI 호출은 0이지만 이전/신규 버전 모두 일반 파일 메뉴 등록 누락으로 2건을 누적하여 전체 종료 코드는 1이다. [업그레이드 출력](../artifacts/file-collect-regressions/m6-upgrade-output.txt), [로그](../artifacts/upgrade-smoke/bdd8a47a7afc4813b47c474c3dae3160/)에 기록했다. 실제 복사본과 통합문서 4개·내부 보고서 7개를 합친 34개 산출물은 M6 후에도 모두 SHA-256이 같았다. [보존 확인](../artifacts/file-collect-live/retention-final-after-m6.json), [M6 최종 요약](../artifacts/file-collect-regressions/m6-final-summary.json)이 최종 상태를 정리한다.

레지스트리 조회와 MSI 로그의 불일치는 원인 미확정이다. [읽기 전용 조사](../artifacts/file-collect-regressions/m6-msix-readonly-diagnosis.md)에서 도구 자식 프로세스는 패키지 ID가 없고 MSI와 동일 사용자 SID였으므로, 상위 앱의 MSIX 패키지만으로 원인이라고 단정하지 않는다. [최종 머신 상태](../artifacts/file-collect-regressions/final-machine-state.json)는 시험 값 정리 후 제품 설치 폴더·x86/x64 COM 등록·LoadBehavior가 없음을 확인하며, 이 최종 부재는 앞선 ON 제거 실패를 통과로 바꾸지 않는다.

## 6. 산출물과 남은 수용 항목

| 구분 | 산출물/상태 |
| --- | --- |
| 통합 소스·시험 | 현재 저장소의 변경 파일. 공개 푸시·릴리스는 이 작업의 완료 사실로 포함하지 않음 |
| 실제 M5에 사용한 MSI | [별도 보관 패키지](../artifacts/file-collect-live/m5-tested-msi/FileListToExcel-1.2.0-m5-tested.msi). 후속 후보와 구분 |
| SDI 실제 시험에 사용한 후보 MSI | [보관된 이전 SDI MSI](../artifacts/FileListToExcel-1.2.0-sdi-final.msi) — 51,842,841 bytes, ProductCode {ABF0839A-8787-4D9A-9912-A4BD8E41CE50} |
| SDI 실측 후보 SHA-256 | c89582f4837dbf297a7649289943f83d42dba4c128bac4a1822b910f28486801 — 메뉴/복사 실측 완료, 결과 UI 실패로 후속 후보가 필요 |
| 최종 UI 수정 후보 MSI | [FileListToExcel-1.2.0-win-x64.msi](../artifacts/release/FileListToExcel-1.2.0-win-x64.msi), 51,842,841 bytes, ProductCode {E16D738C-89DF-410B-9562-C1227CAC2C3C}. [별도 보관본](../artifacts/FileListToExcel-1.2.0-ui-final.msi) |
| 최종 후보 SHA-256 | 233a93be0b56068ac0a13ae37bc0f85668b6b9bf080dbc80e57f93670610ea5a — 실제 결과 UI 통과. M6 세 경로 전체 시험 실패, 등록/제거 원인 및 비활성 상태 보존 미확정 |
| 최신 후보 빌드 | 루트 폴더 행/결과 UI/SDI 메뉴 변경 포함. [네이티브 SDI 결과](../artifacts/file-collect-regressions/native-review/excel-sdi-native.json), [독립 검토](../artifacts/file-collect-regressions/native-review/sdi-review.txt), [최종 UI 수정 MSI 메타정보](../artifacts/file-collect-regressions/native-review/msi-metadata-ui-final.json). 실제 SDI 메뉴/복사는 통과. 결과 UI 예외 수정·185개 회귀·재빌드 및 최종 실제 결과 UI 재시험 통과 |
| 네이티브 바이너리 동일성 | 최종 UI 수정 후보의 x86/x64 추가 기능 DLL은 SDI 실측 후보와 동일. [패키지 증거](../artifacts/file-collect-regressions/native-review/final-package-ui-fix.json), [동결 해시](../artifacts/file-collect-regressions/native-review/native-frozen-identities.json). helper 변경 후 최종 실제 UI 재시험도 별도로 통과 |
| 사용자/설치 안내 | [README](../README.md), [installer 안내](../installer/README.md), [ADR](adr/0003-excel-file-collect.md), [스키마](FILE_COLLECT_SCHEMA.md) |
| 로컬 실제 증거 | [M1/M5 세션](../artifacts/file-collect-live/), [네이티브·MSI 메타정보](../artifacts/file-collect-regressions/native-review/). 원시 경로 정보는 공개 전에 제거 필요 |
| 서명·배포 승인 | 서명되지 않은 평가 빌드. 게시자 인증서/사내 정책 승인/배포 승인은 별도 필요 |

깨끗한 프로필/VM, 재로그인/재부팅, 실제 Excel x86, 전체 동시성/타제품 제거 순서, Excel 미설치, 실제 SMB/cloud/EFS/디스크 부족, 동일 조건 성능 전후 비교는 미실행이다. M6의 OFF·ON·업그레이드 자체는 모두 실행했으나 등록/제거 실패와 비활성 상태 보존 미확정이 남았다. 최종 Excel의 종료 지연 원인을 분리하는 추가 live 시험은 동시 작업 때문에 실행하지 않았다. 최근 시험 도구의 RCW/소유권 개선은 빌드만 확인한 변경이며 제품 MSI 변경이나 실제 Excel 재시험 합격이 아니다. 현재 전체 수용 완료 또는 출시 가능으로 보고하지 않는다.
