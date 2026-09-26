# 2026-09-24 평가 후보 기록

아래는 1.2.0 정식 릴리즈 전 평가 기록을 그대로 보존한 것입니다. 당시 실패·미확정 상태는 해당 후보와 실행 환경의 기록이며 현재 릴리즈 판단은 [검증 기록](FILE_COLLECT_VALIDATION.md)을 참조합니다.

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


---

# File Collect 검증 기록 — v1.2 평가 후보

기록일: 2026-09-24. [작업지시서](../FileListToExcel_FileCollect_WorkOrder.md)의 M0–M7 기준으로 구현, 자동 시험, 설치된 제품의 실제 Excel 시험, 미실행 항목을 구분한다.

**구현·평가 MSI는 완료했고 실제 Excel 자동 로드·복사·결과 UI는 통과했지만, 전체 출시 게이트는 통과하지 못했다.** M6의 OFF·ON·1.1→1.2 업그레이드 경로를 모두 실행했으며 세 경로의 전체 종료 코드는 모두 1이다. 일반 파일 메뉴 등록 누락과 ON 제거 후 시험 값 잔류는 실패로, 비활성 상태 보존은 미확정으로 남았다. 원인은 확정하지 않았다. 최종 UI 시험 Excel은 Quit 반환·창 닫힘 뒤 상당한 지연 후 최종 부재를 확인했으며 지연 원인은 미확정이다. 이전 M5, SDI 시험 후보, 최종 UI 수정 후보를 구분한다.

## 환경과 증거 범위

- 실제 호스트: Windows x64, Microsoft Office Click-to-Run Excel x64 16.0.20326.20158. 기록 JSON의 Excel 객체 모델은 Version 16.0 / Build 20326을 반환했다.
- 개발 도구: .NET SDK 10.0.401, 네이티브 C++ 추가 기능 x86/x64, 자체 포함 Windows x64 helper, WiX MSI.
- 시험은 이 작업이 만든 합성 파일·통합문서와 제품 시험 설치를 대상으로 했다. 사용자 문서를 복사하거나 변경하는 시험은 하지 않았다.
- 설치기가 COM/추가 기능 등록을 수행했다. 실제 Excel 시험에서 VBA/VBS 실행, 초기화 매크로, 추가 기능 수동 등록/Connect=true, Trust Center·매크로 정책 변경을 사용하지 않았다.
- 깨끗한 사용자 프로필/격리 VM은 아니다. 재로그인·재부팅 증거도 없으므로 작업지시서의 전체 자동 로드 수용 행렬을 모두 통과했다고 주장하지 않는다.
- 개발자용 Excel 수용 시험 도구가 시험 통합문서의 선택·필터·정렬 상태를 설정하고 관찰했다. 이는 실제 설치된 Excel/추가 기능의 시험이며 최종 사용자에게 개발 SDK나 시험 스크립트를 요구하는 배포 절차가 아니다.

로컬 증거는 저장소의 ignored artifacts 아래에 남긴다. 원시 JSON·MSI 로그·보고서는 원본 절대 경로와 로컬 환경 정보를 포함할 수 있으므로 그대로 공개 이슈나 릴리스에 첨부하지 않는다. 이 문서에는 개인 원본 경로를 전사하지 않는다.

## 마일스톤별 결과

| 단계 | 실행 및 관찰 | 상태/한계 |
| --- | --- | --- |
| M0 | 기존 코드·명세 조사, 선택형 COM 추가 기능 ADR, 기존 본체와 설치 소유권 분리 | 구현. [ADR](adr/0003-excel-file-collect.md) |
| M1 | MSI 설치 후 새 Excel 프로세스에서 Connected=true, 제품 메뉴 존재, 파일 행 선택 시 활성화. M1 Excel 종료 후 M5의 다른 Excel 프로세스에서도 자동 로드 | 실제 Excel x64 확인. 깨끗한 프로필/재로그인/재부팅은 미실행 |
| M2 | Files/분할 Files, Duplicates, Matches의 숨김 행 계약·짧은 사용자 지정 속성. Open XML 검증 및 기존 writer 회귀 | 자동 시험 통과. 정렬/열 이동/시트 이름 변경/Save As는 아래 M5 실측 포함 |
| M3 | 현재 Excel 메모리의 필터·선택을 캡처. 숨김 행 제외, 겹치는 셀의 행 중복 제거, 전체 행의 표 교차. 관련 없는 셀 혼합은 실제 거절 | SDI 후보에서 일반/미래 버전/기술 열 삭제/ Summary 비활성화와 현재 창 재활성화 확인 |
| M4 | 원본·조상 핸들 검증, 새 하위 폴더, 충돌 구분/no-overwrite, 취소·개별 결과, 내부 보고서 분리 | Windows 파일시스템 시험 통과. 실제 SMB/cloud/EFS/디스크 부족은 미실행 |
| M5 | 설치된 추가 기능 → helper → 목적지 선택 → 실제 복사 및 해시 대조 → 결과 보기/정상 닫기 | 최종 233a… MSI로 실제 우클릭·복사·결과 UI 회귀 통과. 이전 후보 이력과 분리 |
| M6 | 관리 코드·개발 네이티브·설치된 기능 시험과 엄격한 MSI 시험을 별도로 실행 | **OFF·ON·실제 1.1→1.2 업그레이드 실행 완료. 세 경로 모두 전체 종료 코드 1. 등록/제거 실패 및 비활성 상태 보존 미확정** |
| M7 | 소스·시험·후보 패키지·스키마·설치 안내·이 기록 | 평가 산출물. 서명/배포 승인 및 최종 수용 판단은 별도 |

## 자동 시험 기록

횟수는 서로 다른 시점의 실행 결과이며 더해서 최종 전체 시험 수로 표시하지 않는다.

| 실행 | 결과 | 범위 |
| --- | --- | --- |
| 변경 전 기준 | Core 90/90, App 31/31 | 기존 목록/중복 찾기/요청 동작 기준 |
| 기능 통합 후 전체 실행 | Core 135/135, App 44/44 | 당시 소스 기준 관리 코드 전체 시험 |
| writer 및 계약 집중 실행 | 47/47 | 기존 목록·중복 workbook 회귀와 계약. 65,001행 분할 및 10,000개 실제 파일 포함 |
| 엄격한 UUID 검사 후 계약 실행 | 13/13 | 버전·필드·중복 JSON·정밀도·길이·숨김 기술 열·등록 대상 구분 |
| 추가 안전 회귀 후 복사 시험 | 35/35 | ACL 시험 3개 포함. 전체 Core의 후속 집중 실행이며 별도 합산하지 않음 |
| 루트 폴더 행 수정 후 App 실행 | 46/46 | 요청 소비·경로/이름 대응·폴더 행 및 기존 App 회귀 |
| 결과 grid 예외 수정 전 전체 실행 | Core 138/138 + App 46/46 = 184/184, 실패/건너뜀 0, 종료 코드 0 | 루트 폴더 행·보고서 UI 수정 후 전체 솔루션 Release 실행. 네이티브 SDI 수정/최종 MSI 검증과는 별개 |
| 개발 네이티브 셸 | 414개 검사 통과 | 인접 RequestSink를 사용하는 helper dispatch 포함. 설치된 실제 helper 시험과 구분 |
| 설치된 네이티브 셸 | 253개 검사 통과 | 실제 설치 DLL의 표준 호출. 개발 전용 --dispatch를 적용한 실패 실행은 해당 시험 전제 조건 불일치이며 414개 합격으로 바꾸지 않음 |
| 설치된 helper 기능 | 별도 시험 통과 | 목록/중복/Matches/선택 파일, SQLite 캐시 재사용, 원본 유지 및 helper 종료. 등록 키 실패를 상쇄하지 않음 |
| Excel 추가 기능 네이티브 | x64 및 x86 실행 성공 | COM 수명주기/exports/요청 식별자. 실제 x86 Office 로드 증거가 아님 |
| 결과 grid 예외 수정 후 최신 전체 실행 | **Core 138/138 + App 47/47 = 185/185, 실패/건너뜀 0, 종료 코드 0** | 실제 숨김 STA DataGridView의 FormattedValue/DataError 회귀 추가. 최종 MSI 실제 UI 수용은 별도 |
| 최종 SDI 수정 네이티브 | x64 및 x86 mock 시험 통과, 별도 읽기 전용 검토에서 추가 차단 결함 없음 | 현재 창 컨트롤, 중복 콜백, byref Target, 시작 화면, 재진입 연결 해제. 실제 Excel 시험과 구분 |

복사 시험은 일반/UNC 문자열 경로와 장치·ADS·상대 경로 거절, 이름 충돌과 동일 경로 중복, dotfile/0바이트/확장자 없음/한글/긴 경로, read-only/Zone.Identifier 유지, 삭제·변경·폴더·Offline 제외, 원본 write-lock, 검증 후 교체, 진행 중 원본 변경 시도 차단, 큰 파일 취소, 목적지 취소, junction 조상, 경쟁 목적 파일 보존, 동시 작업 분리, 실제 접근 거부 ACL을 포함한다. Offline 속성 합성 시험은 실제 cloud provider의 no-recall 검증을 대신하지 않는다. 허용 경로 문자열의 UNC 시험도 실제 SMB 전송 시험을 대신하지 않는다.

최신 전체 실행 증거: [185개 요약](../artifacts/validation/collect/final-185-summary.md), [전체 명령 출력](../artifacts/validation/collect/final-185-output.txt), [Core/App TRX 디렉터리](../artifacts/validation/collect/final-185/), [결과 grid 수정 전 실패/수정 후 통과 증거](../artifacts/validation/collect/report-grid-regression-summary.md). 앞선 [184개 전체 실행](../artifacts/validation/collect/final-184-summary.md)은 이전 이력으로 보존한다.

로컬 증거: [기준 TRX 보관 폴더](../artifacts/test-results/), [셸 414개 결과](../artifacts/file-collect-regressions/native-review/shell-dispatch.json), [x86/x64 COM 시험](../artifacts/file-collect-regressions/native-review/excel-native.json), [정적/네이티브 검토](../artifacts/file-collect-regressions/native-review/readiness.md). 이전 집중 실행과 최신 전체 실행은 별도 증거로 보관하며, 오래된 TRX를 최신 실행 증거로 재명명하지 않았다.

## 실제 설치된 Excel 결과

| 시나리오 | 실제 결과 | 로컬 증거 |
| --- | --- | --- |
| MSI 후 첫 Excel 시작 | Connected=true, 메뉴 존재, 데이터 선택 활성화 | [M1 시작](../artifacts/file-collect-live/m1-session/started.json), [선택](../artifacts/file-collect-live/m1-session/result-1.json) |
| 종료 후 별도 Excel 프로세스 | M1 종료 기록, M5 새 프로세스의 자동 연결/메뉴 | [M1 종료](../artifacts/file-collect-live/m1-session/closed.json), [M5 시작](../artifacts/file-collect-live/full-session/started.json) |
| 100개 중 필터로 7개 표시 | 저장하지 않은 필터 상태에서 f001–f007만 복사. 각 SHA-256 동일, 전달 폴더에 내부 보고서 없음 | [필터 상태](../artifacts/file-collect-live/full-session/result-1.json), [복사 검증](../artifacts/file-collect-live/full-session/copy-verification.json) |
| 필터 결과에서 한 행 수동 숨김 + 같은 행을 두 열에서 중복 선택 | 목적지 화면의 대상이 6개. 목적지 취소 후 6행 Cancelled 기록 | [겹치는 선택](../artifacts/file-collect-live/full-session/result-3.json), [실행](../artifacts/file-collect-live/full-session/result-4.json). [세 작업 보고서 감사](../artifacts/validation/collect/live-report-audit.json) |
| 내림차순 Table 정렬 + 전체경로 열 이동 + 시트 이름 변경 + Save As + 전체 행 2:3 선택 | f099/f100 복사본만 생성 | [전체 행 선택](../artifacts/file-collect-live/full-session/result-10.json), [결과](../artifacts/file-collect-live/full-session/transformed-copy.json) |
| 정상 표 셀과 관계없는 셀 혼합 | 실제 오류 안내, helper 미실행 | [거절 기록](../artifacts/file-collect-live/full-session/mixed-selection-rejected.txt) |
| 기술 열 삭제 후 실제 명령 클릭 | 실제 Excel automation 오류로 거절. 메뉴 비활성화 상태와는 별도 | [클릭 거절](../artifacts/file-collect-live/full-session/technical-deletion-rejected.txt) |
| 정상 종료와 helper 수명 | 시험 통합문서 Close/Quit 후 해당 Excel PID와 helper 프로세스 없음 | [프로세스 종료 확인](../artifacts/file-collect-live/full-session/process-exit.json) |
| M5 시험 설치 정상 제거 | Windows Installer 제거 종료 코드 0 | [M5 제거 로그](../artifacts/file-collect-live/m5-uninstall.log). 최종 후보의 전체 M6 회귀와 구분 |
| 일반 통합문서/기술 열 삭제의 메뉴 표시 (이전 M5) | 이전 창 상태가 남는 SDI 메뉴 문제 발견. 후속 실제 회귀는 다음 절에 별도 기록 | [일반 표 관찰](../artifacts/file-collect-live/full-session/result-19.json), [기술 열 삭제 관찰](../artifacts/file-collect-live/full-session/result-21.json). **합격 아님** |

M5 시험 설치 패키지는 [보관된 M5 MSI](../artifacts/file-collect-live/m5-tested-msi/FileListToExcel-1.2.0-m5-tested.msi)로 분리했다. 후속 루트 폴더 행/결과 UI/SDI 메뉴 변경과 이후 수정은 이 MSI의 실측에 포함되지 않는다. 한 번의 메뉴/연결 관찰은 모든 이벤트, 여러 Excel 프로세스의 동시 사용 또는 다른 추가 기능 공존을 증명하지 않는다.

## SDI 수정 후보의 실제 Excel 재시험과 새 결과 UI 결함

시험 MSI의 ProductCode는 {ABF0839A-8787-4D9A-9912-A4BD8E41CE50}, SHA-256은 c89582f4837dbf297a7649289943f83d42dba4c128bac4a1822b910f28486801이다. 이 패키지는 결과 UI 수정 후보로 교체되었으며 최종 배포 해시로 사용하지 않는다.

| 실제 시험 | 결과 | 증거 |
| --- | --- | --- |
| 서로 다른 Excel 프로세스 2개 | 두 인스턴스 모두 Connected=true, EnableEvents=true, Cell/List Range Popup의 제품 메뉴 각각 1개 | [SDI-A 시작](../artifacts/file-collect-live/sdi-a/started.json), [SDI-B 시작](../artifacts/file-collect-live/sdi-b/started.json) |
| 일반 workbook와 지원 workbook 전환 | 일반 workbook 비활성, 지원 workbook로 돌아오면 활성 | [일반](../artifacts/file-collect-live/sdi-a/result-2.json), [지원 표 재활성](../artifacts/file-collect-live/sdi-a/result-3.json) |
| 미래 버전 999/스키마 1 복구 | 미래 버전은 선택 변경 후에도 비활성, 스키마 1 복구 후 활성 | [미래 버전](../artifacts/file-collect-live/sdi-a/result-4.json), [미래 버전 선택](../artifacts/file-collect-live/sdi-a/result-5.json), [복구](../artifacts/file-collect-live/sdi-a/result-6.json) |
| 미저장 100→7 필터와 실제 복사 | 7개만 복사, fixture와 동일 | [필터](../artifacts/file-collect-live/sdi-a/result-8.json), [명령](../artifacts/file-collect-live/sdi-a/result-9.json), [해시 대상 MSI를 명시한 복사 검증](../artifacts/file-collect-live/sdi-a/final-copy-verification.json) |
| 기술 열 삭제 | 메뉴 비활성 | [실측](../artifacts/file-collect-live/sdi-a/result-10.json) |
| 새 창 Duplicates / Summary | Duplicates 3개 선택·복사, Summary는 비활성 | [선택](../artifacts/file-collect-live/sdi-a/result-12.json), [실행](../artifacts/file-collect-live/sdi-a/result-13.json), [Summary](../artifacts/file-collect-live/sdi-a/result-14.json) |
| 다른 Excel 인스턴스의 Matches | A6:A7에 해당하는 2개 복사 완료 및 보고서 저장 | [선택](../artifacts/file-collect-live/sdi-b/result-1.json), [실행](../artifacts/file-collect-live/sdi-b/result-2.json) |
| 작업 결과 보기 | **실패**: DataGridView의 정수 RowIndex와 FormattingApplied=true 처리에서 반복 FormatException | [실제 결함 기록](../artifacts/file-collect-live/sdi-b/result-ui-regression.txt). 복사 완료 후 발생. 수정/재빌드 및 185개 회귀 완료. 다음 절의 최종 233a… MSI 실제 UI 재시험에서 해결 확인 |
| 시험 종료 | 두 시험 Excel의 Close/Quit 호출이 정상 반환. 반복 결과 UI 오류의 정상 종료 시도가 실패하여 복사 완료된 이 작업 소유 helper PID 7264만 종료 | [A 종료](../artifacts/file-collect-live/sdi-a/closed.json), [B 종료](../artifacts/file-collect-live/sdi-b/closed.json), [종료 사유](../artifacts/file-collect-live/sdi-b/result-ui-regression.txt). Excel 프로세스는 강제 종료하지 않음 |

SDI 메뉴의 이전 창 상태 문제는 이 실제 재시험에서 해결을 확인했다. 복사 자체의 성공, 메뉴 상태 검증, 상세 결과 UI의 실패를 별도로 기록하며 이 후보의 전체 UI 수용을 합격으로 표시하지 않는다.

## 최종 UI 수정 MSI의 실제 우클릭·복사·결과 화면 재시험

최종 시험 대상은 ProductCode {E16D738C-89DF-410B-9562-C1227CAC2C3C}, SHA-256 233a93be0b56068ac0a13ae37bc0f85668b6b9bf080dbc80e57f93670610ea5a의 MSI다.

| 실제 시험 | 결과 | 증거 |
| --- | --- | --- |
| 새 Excel 정상 시작 | Excel 16.0 / Build 20326, Connected=true, EnableEvents=true, Cell/List Range Popup 메뉴 각각 1개 | [시작](../artifacts/file-collect-live/ui-final/started.json) |
| Matches A6:A7 선택 | 메뉴 활성. 실제 셀 우클릭 → 파일목록 → 선택한 파일 복사 UI 클릭 | [선택](../artifacts/file-collect-live/ui-final/result-1.json) 및 실제 UI 시험 |
| 목적지 선택·복사 | 화면에 2행/44 bytes, “여기에 복사” 클릭 후 2개 복사. 두 파일 모두 fixture SHA-256과 동일 | [MSI 해시를 명시한 복사 검증](../artifacts/file-collect-live/ui-final/copy-verification.json) |
| 작업 결과 보기·닫기 | 실제 결과창 2행에 선택 순번 1/2, 한국어 “복사 완료” 표시. FormatException/DataError 없이 열리고 정상 닫힘 | [접근성 기록](../artifacts/file-collect-live/ui-final/result-grid-accessibility.txt), [정상 닫기 확인](../artifacts/file-collect-live/ui-final/copy-verification.json) |
| helper·Excel 종료 | 완료 helper는 Alt+F4로 정상 종료. Excel은 Quit 반환·창 닫힘 후 상당한 지연 뒤 최종 부재 확인. 이전 harness·helper도 없음. Excel 강제 종료 없음. 다른 Excel automation과 동시 상태였으며 지연 원인 미확정 | [종료 요청](../artifacts/file-collect-live/ui-final/closed.json), [프로세스 확인](../artifacts/file-collect-live/ui-final/process-exit.json), [중간 잔류 관찰](../artifacts/file-collect-live/ui-final/process-exit-pending.txt), [최종 부재](../artifacts/file-collect-regressions/final-machine-state.json). 결과창 닫기와 프로세스 소멸 시점을 구분 |

최종 MSI에서 이전 결과 grid 결함의 실제 재현 절차를 통과했다. 이 합격은 실제로 수행한 2행 Matches 흐름의 결과이며, 미실행 환경이나 설치 M6 행렬까지 확장하지 않는다.

## 최종 후보와 설치 회귀 — 엄격한 게이트 실패 및 추가 진단

SDI 실측에 사용한 당시 후보 파일은 다음과 같다. [보관된 SDI 시험 MSI](../artifacts/FileListToExcel-1.2.0-sdi-final.msi) — 51,842,841 bytes, ProductCode {ABF0839A-8787-4D9A-9912-A4BD8E41CE50}, SHA-256 c89582f4837dbf297a7649289943f83d42dba4c128bac4a1822b910f28486801. [SDI 시험 당시 MSI 메타정보](../artifacts/file-collect-regressions/native-review/msi-metadata-sdi-final.json)를 확인했다.

결과 grid는 RowIndex 표시를 CurrentCulture 문자열로 바꾸어 수정했으며, 실제 숨김 STA Form/DataGridView의 모든 셀 형식·다섯 결과 라벨·원래 행 인덱스 유지·DataError 부재를 시험했다. 수정 전 회귀 실패와 수정 후 통과를 별도로 보존했다.

**최종 UI 수정 후보:** [FileListToExcel-1.2.0-win-x64.msi](../artifacts/release/FileListToExcel-1.2.0-win-x64.msi), 51,842,841 bytes, ProductCode {E16D738C-89DF-410B-9562-C1227CAC2C3C}, SHA-256 233a93be0b56068ac0a13ae37bc0f85668b6b9bf080dbc80e57f93670610ea5a. [별도 보관본](../artifacts/FileListToExcel-1.2.0-ui-final.msi) 및 [패키지 증거](../artifacts/file-collect-regressions/native-review/final-package-ui-fix.json)를 남겼다. 네이티브 x86/x64 DLL은 실제 SDI 시험 때와 동일하며 [동결 바이너리 해시](../artifacts/file-collect-regressions/native-review/native-frozen-identities.json)로 확인한다. helper 변경 후에도 위 최종 실제 우클릭·결과 UI 시험을 별도로 수행했다.

SDI 수정 후 x86/x64 [네이티브 결과](../artifacts/file-collect-regressions/native-review/excel-sdi-native.json)와 [독립 검토](../artifacts/file-collect-regressions/native-review/sdi-review.txt)가 통과했다. 이 후보의 실제 설치·SDI 재시험은 위에 기록했지만 상세 결과 UI는 실패했다. UI 수정 최종 MSI의 실제 결과 보기·닫기 재시험은 위에서 통과했다. 아래 결과는 최종 패키지를 사용한 실행 로그와 [M6 최종 요약](../artifacts/file-collect-regressions/m6-final-summary.json)에 따른다. 패키지 이름이나 빌드 성공, 개별 msiexec 종료 코드만으로 전체 합격 처리하지 않는다.

| 확인 항목 | 현재 기록 |
| --- | --- |
| 최종 소스에 대응하는 배포 MSI 파일·SHA-256 | **고정: E16D738C… / 233a93be… 위 최종 후보. 실제 UI 통과, M6 세 경로 모두 전체 시험 실패** |
| 최종 MSI 정상 설치 → 새 Excel 자동 로드 → 복사 → 결과 보기 | **실제 UI 통과: 2행/44 bytes, 두 SHA-256 동일, 결과창 정상 열기·닫기** |
| 기본 본체만 설치/복구/제거 | **전체 단계 완료: 설치/복구/제거 MSI는 각각 0, 등록 실패 2건을 누적한 시험 전체는 종료 코드 1** |
| Excel 구성 요소 포함 설치/복구/제거, 비활성 LoadBehavior 유지 | **실행 완료: 설치/복구/제거 MSI 각각 0. 비활성 상태 보존 미확정, 제거 후 시험 LoadBehavior=2 잔류로 전체 종료 코드 1** |
| 이전 버전 무인 업그레이드에서 Excel 기능이 켜지지 않음 | **실제 1.1→1.2 ProductCode 교체·Excel 기능 자동 추가 방지 확인. 기능/보존/제거 통과, 이전·신규 버전의 메뉴 등록 실패 2건으로 전체 종료 코드 1** |
| SDI 메뉴 활성 상태/결과 UI | **SDI 메뉴 실제 회귀 통과. FormatException 수정·STA 회귀 및 최종 MSI의 실제 결과 보기·닫기 통과** |

### M6 전 경로 실행 결과와 실패·미확정 항목

일반 파일 메뉴 키 HKCU\Software\Classes\*\shellex\ContextMenuHandlers\FileListToExcel을 검사하는 strict OFF 시험이 두 번 실패했다. 최초 출력은 PowerShell null 호출이었고 추가 검사에서 키 자체가 없음을 확인했다. MSI Registry 테이블과 설치 로그에는 해당 등록이 올바르게 들어 있으나, reg.exe 64/32비트의 HKCU/HKCR/HKLM 및 직접 HKU 사용자 SID Classes 조회 모두 해당 별표(*) 제품 키를 찾지 못했다. Directory와 Directory Background 등록은 존재했다.

증거: [첫 실패](../artifacts/file-collect-regressions/m6-off-first-output.txt), [실패 위치](../artifacts/file-collect-regressions/m6-off-trace-output.txt), [PowerShell/직접 레지스트리 비교](../artifacts/file-collect-regressions/m6-registry-provider-diagnosis.txt), [뷰별 조회](../artifacts/file-collect-regressions/m6-independent-registry-views.json), [직접 사용자 Classes 조회](../artifacts/file-collect-regressions/m6-independent-direct-user-classes.json), [부모 키](../artifacts/file-collect-regressions/m6-independent-parent-keys.json), [MSI Registry 테이블](../artifacts/file-collect-regressions/m6-independent-msi-registry-table.json).

이는 PowerShell provider만의 문제라고 단정할 수 없다. [기존 v1 검증](VALIDATION.md)에 비슷한 로컬 관찰이 있지만 이번 실패의 원인이 같다는 증거도 없다. 환경 탓이나 알려진 무해한 실패로 분류하지 않으며, 등록 조건이 미충족인 출시 게이트 실패로 유지한다.

별도로 [설치된 기능 시험](../artifacts/file-collect-regressions/m6-functional-output.txt)은 목록/중복/Matches/선택 파일/캐시 재사용/원본 유지/helper 종료를 통과했다. [설치된 네이티브 시험](../artifacts/file-collect-regressions/m6-installed-native.json)은 표준 호출에서 253개를 통과했다. 개발용 --dispatch는 DLL 옆 RequestSink를 전제로 하므로 설치된 실제 helper에 적용한 실행은 올바른 installed 시험이 아니며, 개발 414개와 설치 253개를 섞어 기록하지 않는다.

OFF 후속 실행은 설치·기능·복구·제거 전 단계를 완료했다. [OFF 최종 출력](../artifacts/file-collect-regressions/m6-off-final-output.txt)과 [해당 설치 시험 폴더](../artifacts/installer-smoke/59e06e4a1f464ba08dbbf3f3ea32b4d1/)에 기록했다. Windows Installer의 설치/복구/제거는 각각 종료 코드 0이었고, 목록·중복·Matches·선택 파일·SQLite 캐시·원본/생성 통합문서 보존·helper 종료 및 설치된 네이티브 표준 253개 시험이 통과했다. 선택형 Excel 기능은 OFF 상태의 설치 전후 모두 없었다. 그러나 설치 및 복구 뒤 일반 파일 메뉴 등록 키가 없는 실패 2건을 누적하여 시험 전체는 **종료 코드 1**을 반환했다. MSI 작업의 종료 코드 0을 전체 수용 통과로 해석하지 않는다.

시험 도구의 기존 네이티브 `&` 호출은 이 호스트에서 LASTEXITCODE가 설정되지 않아 종료 코드 검사에 문제가 있었다. `Start-Process -Wait -PassThru`의 명시적인 ExitCode 확인으로 수정한 후속 실행에서 253개를 정상 확인했다. 이는 앞선 개발 전용 `--dispatch`의 잘못된 installed 호출 이력과 다른 시험 도구 문제이며, 일반 파일 메뉴 등록 누락을 해결한 변경이 아니다.

ON 경로도 설치·기능·복구·제거를 실행했다. 설치/복구에서는 일반 파일 메뉴 키 누락 경고도 계속 발생했다. MSI 설치/복구/제거 각각은 종료 코드 0이었고 기능 및 설치된 네이티브 253개 시험은 통과했다. 최초 x86/x64 Excel COM 등록과 LoadBehavior=3도 확인했다. 시험이 만든 LoadBehavior=2를 복구 후 PowerShell에서 2로 읽은 단언은 통과했지만, 같은 복구 로그의 AppSearch와 RegAddValue는 #3을 기록했다. 따라서 **사용자 비활성 상태를 복구가 보존했다는 검증은 미확정**이다. 제거 후 제품의 나머지 등록 값·설치 파일은 사라졌으나 이 시험이 만든 DWORD 2만 남아 Excel 등록 부재 단언이 실패했고, 전체 종료 코드는 1이다. [ON 출력](../artifacts/file-collect-regressions/m6-on-output.txt), [ON 설치/복구/제거 로그](../artifacts/installer-smoke/035e3bd22438401db39e3438f93347b9/), [잔류 값](../artifacts/file-collect-regressions/m6-on-leftover-excel-key.txt)에 기록했다.

잔류 키가 이 시험에서 만든 LoadBehavior=2 단독 값이고 하위 키·제품 설치 폴더가 없음을 확인한 뒤 그 시험 값만 정리했다. 사용자/Office의 기존 값은 아니며, [정리 기록](../artifacts/file-collect-regressions/m6-test-value-cleanup.txt)은 제거 수용 실패를 통과로 바꾸지 않는다.

업그레이드는 실제 1.1.0 설치 → 1.2.0 업그레이드 → 제거를 수행했다. 각 MSI 호출은 종료 코드 0이며 ProductCode 교체, 새 Excel 기능의 자동 추가 방지, 목록/중복/Matches/선택 기능, 캐시, 원본·생성 통합문서 유지 및 제거 후 보존을 확인했다. 그러나 이전 1.1과 신규 1.2 모두 일반 파일 메뉴 등록 키 누락 경고가 있어 2건을 누적한 전체 시험은 종료 코드 1이다. [업그레이드 출력](../artifacts/file-collect-regressions/m6-upgrade-output.txt), [실제 업그레이드 로그](../artifacts/upgrade-smoke/bdd8a47a7afc4813b47c474c3dae3160/)를 보존한다.

M6 후 실제 복사본·생성 통합문서 4개·내부 보고서 7개를 합친 34개 산출물은 모두 SHA-256이 변하지 않았다. [보존 확인](../artifacts/file-collect-live/retention-final-after-m6.json)과 [최종 전체 요약](../artifacts/file-collect-regressions/m6-final-summary.json)에 기록했다. 진단 계속 옵션은 독립 단계를 관찰하되 실패를 누적하고 전체 종료 코드 1을 유지했다. 실패 단언을 생략하거나 합격으로 바꾸지 않았다.

[읽기 전용 실행 문맥 조사](../artifacts/file-collect-regressions/m6-msix-readonly-diagnosis.md)에서 도구 자식 프로세스는 패키지 ID가 없고 MSI와 같은 사용자 SID임을 확인했다. 상위 앱의 MSIX 패키지 ID만으로 레지스트리 불일치의 원인이라고 단정할 수 없다. 현재 원인은 미확정이며 설치기의 등록·비활성 상태 보존·제거 수용을 통과했다고 주장하지 않는다.

[최종 머신 상태](../artifacts/file-collect-regressions/final-machine-state.json)에서 제품 설치 폴더, x86/x64 Excel COM 등록, LoadBehavior 값이 모두 없음을 확인했다. 이는 시험 값 정리 후의 최종 상태이며 앞선 ON 제거 단언 실패와 구분한다. 최종 시험 Excel·이전 harness·helper 프로세스도 모두 없었다. Excel은 강제 종료하지 않았다.

[MSI 메타정보 검사](../artifacts/file-collect-regressions/native-review/msi-metadata-final.json)는 1.1.0/1.2.0의 서로 다른 ProductCode, 같은 UpgradeCode, 필수 Complete Level 1 및 선택 ExcelIntegration Level 2를 확인한다. 실제 설치/업그레이드 시험을 대체하지 않는다. 기존 셸 등록 복구의 로컬 호스트 차이는 [기존 검증 기록](VALIDATION.md)에 남아 있으며, 해당 단언을 생략한 실행으로 정식 회귀 합격을 주장하지 않는다.

## 미실행 항목과 남은 제한

- 깨끗한 Windows 사용자 프로필/격리 VM, 재로그인/재부팅 후 자동 로드, 실제 Excel x86. 동시 Excel 프로세스 2개의 기본 자동 로드·별도 복사는 확인했으나 모든 동시성 시나리오를 수행한 것은 아니다.
- 실제 SMB 지연·끊김, 실제 cloud-only provider/hydration, 디스크 부족, EFS/DRM 및 대상 파일시스템별 보존, 조직 서명·차단 정책 행렬.
- Excel 미설치/대체 앱의 이번 버전 설치 흐름, 명단대조기/Workspace와의 동시 설치 및 서로 다른 제거 순서. 이들 제품을 설치하거나 제거한 것으로 기록하지 않는다.
- 실제 Excel의 Files_2 및 각 표의 전체 선택 조합, totals-only/다중 표/메타정보 복사·손상 전체 수용 행렬. Files/Duplicates/Matches의 위 기본 복사와 미래 버전 999의 메뉴 비활성화만 실측했다. 자동 계약 시험이나 일부 live 관찰을 전부의 합격으로 넓히지 않는다.
- 새 기술 열의 성능 비용을 비교하는 동일 조건 전후 기준 측정. 대량 회귀 실행은 성능 악화율 측정이 아니다.
- 최종 UI 시험의 Excel 종료 지연 원인을 분리하는 추가 실제 Excel 시험은 동시 작업 때문에 실행하지 않았다. 후속 수용 시험 도구의 RCW/소유권 개선은 빌드만 확인했으며 제품 MSI 변경이나 새 live 합격 증거가 아니다.
- v1 구현 한도는 **선택한 보이는 데이터 행 10,000개 및 요청 32 MiB**이다. 작업지시서의 “고유 파일 10,000개”보다 엄격하므로, 같은 원본이 반복되어도 전송 전 행 한도에 걸릴 수 있다.
- 모든 reparse/recall 조상·파일은 보수적으로 제외한다. 로컬에 내려받은 cloud 항목도 reparse가 남아 있으면 제외된다. 일반 원본 hardlink는 regular file로 허용하며, 요청 파일 hardlink는 거절한다.
- 극단적으로 긴 확장자의 이름 충돌은 CollisionNameTooLong으로 제외할 수 있다. 지원하지 않는 별칭·장치·ADS·특수 경로 및 필요한 파일 식별자/메타정보 보존을 제공하지 않는 파일시스템은 거절한다.
- 크기/UTC 수정시간 및 현재 핸들 검증은 목록 생성 당시 바이트의 인증이나 과거 버전 복원이 아니다. 제품 표식은 형식 식별 정보이며 출처 인증이 아니다.
- 평가 DLL/helper/MSI는 서명되지 않았다. 게시자 인증서, 조직 승인 및 배포 승인은 제공되지 않았다. 이 작업에서 공개 릴리스·푸시를 수행한 것으로 기록하지 않는다.
