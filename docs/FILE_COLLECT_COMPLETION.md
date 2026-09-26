# File Collect 구현·납품 보고 — v1.2.0

기록일: 2026-09-26. [작업지시서](../FileListToExcel_FileCollect_WorkOrder.md)에 따라 기존 File List to Excel에 **Excel 목록에서 선택한 실제 파일의 복사본을 새 폴더에 모으는 기능**을 통합했다. Excel 통합문서의 Ctrl+S를 대신하는 기능이 아니다.

**1.2.0의 선택형 Excel 파일 복사 구현, 자동 회귀, 깨끗한 Windows의 설치·복구·제거·실제 1.1.0 업그레이드와 CI MSI의 실제 Excel 복사를 확인했다.** 이전 평가의 실패와 후속 실행 문맥 진단은 [검증 기록](FILE_COLLECT_VALIDATION.md) 및 [평가 이력](FILE_COLLECT_EVALUATION_HISTORY.md)에 구분해 보존했다.

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

## 5. 최종 검증

- Core 138 + App 47, 총 185개 Release 테스트 통과. 셸/Excel x64/Excel x86의 네이티브 CTest도 모두 통과했다. 선택 캡처 11개 시나리오와 ExcelAcceptance 도구의 CI 컴파일을 추가했다.
- [후보 CI](https://github.com/prozac0401/File-List-To-Excel/actions/runs/36221029040)는 기본/Excel 포함 설치·복구·제거, 합성 0.9.0 및 실제 배포 1.1.0에서의 업그레이드를 모두 통과했다. 비활성 LoadBehavior=2 보존은 값 조회와 MSI 로그를 함께 확인했다.
- 후보 MSI의 실제 Excel x64에서 자동 연결, 메뉴 각각 1개, 미저장 100→7 필터 선택, 실제 우클릭 복사, 7개 SHA-256 일치, 상세 결과 순번/상태, helper와 Excel 정상 종료를 확인했다. 100개 원본은 변하지 않았다.
- 이전 로컬 등록 실패는 동일 사용자·세션의 프로세스 실행 문맥에 따른 조회 차이로 분리했다. 독립 프로세스의 직접 조회와 실제 Excel에서는 등록/자동 로드를 확인했고, 같은 실행 경로의 로컬 strict OFF·ON 및 실제 1.1.0 업그레이드도 모두 통과했다. 내부 기작을 추정하지 않았고 보안 정책이나 등록 값을 바꾸지 않았다.
- 설치 화면의 잘못된 MIT 문구를 저장소 LICENSE의 기존 조건으로 정정했다. 셸 버전 메타정보를 1.2.0으로 맞췄으며 테스트가 0개 실행되는 빌드를 거절하도록 보강했다.

## 6. 배포와 제한

배포: [v1.2.0 Release](https://github.com/prozac0401/File-List-To-Excel/releases/tag/v1.2.0)의 **FileListToExcel-1.2.0-win-x64.msi**, **SHA256SUMS.txt**. main에 포함된 v1.2.0 태그의 필수 자동 게이트를 통과한 패키지만 게시한다. 공개 자산에는 사용자 경로를 담은 로컬 보고서나 테스트 원본을 포함하지 않는다.

실제 Excel x86, 조직 보안 정책, 재로그인/재부팅, 원격 SMB 장애, 다른 제품의 모든 설치·제거 공존 조합 등 미실행 환경은 [검증 기록](FILE_COLLECT_VALIDATION.md)에 명시했다. x86 빌드·COM 시험을 실제 x86 Office 수용으로 표시하지 않는다. 바이너리는 서명되지 않았다.

소스 형식은 [스키마](FILE_COLLECT_SCHEMA.md), 선택 구성 요소 및 자동 로드 경계는 [ADR](adr/0003-excel-file-collect.md)을 참조한다. 평가 당시의 후보 해시·실패·실행 증거는 [평가 이력](FILE_COLLECT_EVALUATION_HISTORY.md)에 남겼다.
