# File List to Excel

Windows 탐색기에서 선택한 파일과 폴더의 목록을 Excel 표로 열고, 내용이 같은 파일을 찾는 Windows 11 x64 도구입니다. 설정 창, 계정, 상주 서비스 없이 실행합니다.

## 설치와 사용

1. [Releases](https://github.com/prozac0401/File-List-To-Excel/releases)에서 MSI를 내려받아 실행합니다.
2. 파일이나 폴더 선택 후 **우클릭 → 더 많은 옵션 표시**에서 Excel 목록 메뉴를 선택합니다.
3. 폴더 안 빈 공간의 메뉴는 현재 폴더를 사용합니다. 하위 폴더 메뉴는 재귀 탐색합니다.
4. Excel에서 필터·정렬하고 필요한 결과를 **다른 이름으로 저장**합니다.

사용자별 설치이며 관리자 권한과 별도 .NET 설치가 필요하지 않습니다. Windows 설정 → 앱 → 설치된 앱에서 제거할 수 있습니다. 설치 직후 기존 탐색기 창에 메뉴가 보이지 않으면 새 창을 열거나 다음 로그인 후 확인합니다. 가상 폴더와 압축파일 내부는 대상이 아닙니다.

Excel Desktop을 우선 실행합니다. 설치되어 있지 않으면 기본 .xlsx 연결 앱을 사용하고, 연결 앱도 없으면 생성 경로를 안내합니다. v1은 x64 Windows 탐색기용이며 Windows ARM64 네이티브 탐색기는 지원하지 않습니다.

## 동작

- 파일 다중 선택: 선택한 항목만 포함하며 같은 폴더의 다른 파일은 포함하지 않습니다.
- 폴더 선택: 직접 자식 또는 전체 하위 트리를 선택할 수 있습니다. 폴더 행도 포함합니다.
- 파일과 폴더 혼합 선택: 기본 명령은 선택 항목 자체를 나열하고, 재귀 명령은 선택 파일과 선택 폴더의 하위 항목을 포함합니다.
- 이름·종류·byte 크기·표시 크기·수정일·생성일·상위 폴더·상대경로·전체경로·깊이·속성·기준 폴더를 기록합니다.
- 숫자와 날짜는 Excel 값으로 저장합니다. 표, 필터, 첫 행 고정, 원본 파일/폴더 링크를 제공합니다.
- 약 700ms 이상 걸리는 작업은 발견 개수와 현재 경로, 취소 버튼을 표시합니다.
- 기존 파일 목록 명령은 파일 내용을 읽지 않습니다. 다운로드를 요청하거나 해시·문서 분석을 실행하지 않습니다.
- 읽기 실패는 Errors 시트에 기록하고 계속합니다. 원본 파일을 변경하거나 삭제하지 않습니다.

생성 파일은 %TEMP%/FileListToExcel에 저장합니다. 사용자가 저장한 결과와 임시 Excel 파일은 제거 시 삭제하지 않습니다. Windows의 임시 파일 정리 대상이 될 수 있으므로 보관할 결과는 별도 저장하세요. 진단 오류는 %LOCALAPPDATA%/FileListToExcel/Logs/last-error.log에 저장하며 네트워크로 전송하지 않습니다.

## 중복 파일 찾기 (v1.1)

- 파일 하나 → **같은 파일 찾기**: 해당 파일의 현재 폴더와 하위 폴더에서 같은 내용을 찾습니다. 기준 파일은 결과 상단에 표시하고 결과 행에서는 제외합니다.
- 파일 여러 개 → **선택 파일 중 중복 찾기**: 선택한 파일끼리만 비교합니다.
- 폴더 하나 또는 여러 개 → **중복 파일 찾기**: 선택 폴더 전체를 합쳐 재귀 검사합니다.
- 폴더 빈 공간 → **현재 폴더 중복 파일 찾기**. 파일·폴더 혼합 선택에서는 기존 목록 메뉴를 제공합니다.

파일명·확장자와 관계없이 **크기 → Quick Fingerprint → SHA-256**으로 비교합니다. 크기가 혼자인 파일은 읽지 않습니다. 1 MiB 이하 후보는 SHA-256을 한 번 계산하고, 큰 파일은 시작·중간·끝 64 KiB씩 읽은 지문이 같은 후보만 전체 해시합니다. 0 byte 파일은 읽지 않고 그룹화합니다. 읽기는 4 MiB 버퍼와 최대 2개 worker를 사용하며 UNC 경로는 1개 worker를 사용합니다.

결과는 `Duplicates` 또는 `Matches`, `Summary`, `Errors`, `Skipped` 시트로 구성됩니다. 큰 중복 그룹부터 표시하며 숫자 크기, 날짜, 필터, 고정 행과 파일 링크를 제공합니다. 예상 중복 용량은 그룹마다 한 파일을 남긴다는 가정의 계산값입니다. **원본을 삭제·이동·수정하거나 정리를 권고하지 않습니다.**

Hash Cache는 `%LOCALAPPDATA%/FileListToExcel/hash_cache.sqlite`에 저장됩니다. 경로·크기·UTC 수정시간·알고리즘이 같을 때 재사용하고, 손상된 캐시는 보관 후 재생성합니다. 사용할 수 없으면 캐시 없이 검사합니다. 캐시와 생성한 Excel은 제거 후에도 보존됩니다. 앱을 종료한 상태에서 캐시를 삭제해도 다음 검사에서 다시 생성됩니다.

Cloud-only(Offline/Recall) 파일을 내려받지 않습니다. 중복 검사에서는 안전을 위해 모든 파일/폴더 reparse point와 그 경유 경로를 제외하므로, 로컬에 내려받은 OneDrive 항목도 reparse 속성이 남아 있으면 제외됩니다. 접근 거부·사용 중·검사 중 변경된 파일은 기록하고 나머지를 계속 검사합니다. 약 700 ms 이후 단계별 진행 창과 취소 버튼이 나타납니다.


## Excel에서 선택한 파일 복사 (v1.2, 선택 설치)

v1.2.0은 Excel에서 선택한 목록 행의 실제 파일을 새 폴더로 복사하는 선택형 연동을 추가합니다. 릴리즈는 자동 테스트와 본체/Excel 연동의 설치·복구·제거, 기존 배포본 업그레이드 검사를 통과해야 게시됩니다. 실제 시험 결과와 환경별 제한은 [구현·검증 보고](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/FILE_COLLECT_COMPLETION.md)에 기록합니다.

설치 화면에서 **Excel에서 선택한 파일 복사 (선택)** 구성 요소를 선택하면 다음 Excel 시작부터 **셀 우클릭 → 파일목록 → 선택한 파일 복사…**가 준비됩니다. 기본 선택은 꺼져 있으며 기존 사용자의 무인 업그레이드에 자동 추가되지 않습니다. Excel Desktop이 필요합니다. 작업 중인 통합문서를 저장하고 Excel을 다시 여는 것 외에 VBA/.vbs 실행, 매크로 허용, 추가 기능 찾아보기/수동 체크는 필요하지 않습니다.

새로 만든 Files / Duplicates / Matches 결과표에서 파일명 셀만 선택해도 됩니다. Ctrl 다중 선택을 지원하고 필터 및 수동 숨김 행은 제외합니다. 메뉴를 누른 뒤 목적 폴더 화면에서 개수·예상 크기·제외 수를 확인하고 **여기에 복사**를 누르면, 새 `모은파일_날짜_시간_식별자` 하위 폴더에 실제 파일 복사본을 모읍니다. 원본 이동·이름 변경·삭제와 기존 파일 덮어쓰기는 하지 않습니다. 같은 이름은 `이름 (2).확장자`로 구분합니다.

결과 파일은 매크로 없는 일반 .xlsx입니다. 연동을 설치하지 않아도 목록과 중복 찾기는 그대로 동작합니다. 이전 결과에는 새 행 식별 정보가 없어 새 목록을 만들어야 합니다. 식별 정보는 출처 인증이 아니며 경로·파일 종류·크기·UTC 수정시간을 다시 검사합니다. cloud-only/recall/reparse 경로, 변경·삭제된 원본은 제외합니다. 내려받은 OneDrive 파일도 reparse 속성이 남아 있으면 제외될 수 있습니다. 목적 파일시스템은 ADS와 영구 ACL을 지원해야 합니다. 원본의 소유자·접근 권한 목록 자체를 복제하는 기능은 아닙니다.

전달 폴더에는 복사 파일만 들어갑니다. 내부 원본 경로가 포함된 상세 결과는 `%LOCALAPPDATA%/FileListToExcel/Reports`에 보관하고 완료 화면의 **작업 결과 보기**에서 확인합니다. 취소 시 완료본을 유지하고 자체 미완료본만 정리합니다. 원본 파일에 접근할 수 없는 다른 PC에서는 목록 파일만으로 복사할 수 없습니다.

일반 Windows x64에서 Excel x86/x64용 컴파일 추가 기능을 패키징합니다. 조직 추가 기능 정책·서명 요구·Office 비활성화 정책은 존중하며 자동으로 보안 설정을 낮추거나 차단을 해제하지 않습니다. 배포 바이너리는 서명되지 않았습니다. 실제로 시험한 환경과 미실행 항목은 [파일 복사 검증 기록](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/FILE_COLLECT_VALIDATION.md), 구현 경계는 [ADR](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/adr/0003-excel-file-collect.md), 형식은 [스키마](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/FILE_COLLECT_SCHEMA.md)를 참조하세요.

## 대량 및 예외 처리

한 번에 선택해 전달할 수 있는 입력 경로는 최대 100,000개이며, 셸 요청 파일은 32MiB까지 지원합니다. 폴더 재귀 탐색 결과 행 수에는 이 입력 개수 제한이 적용되지 않습니다.

Excel 링크 제한을 피하기 위해 시트당 65,000행을 기록하고 Files_2 등으로 자동 분할합니다. 매우 긴 링크는 전체경로를 보존하고 Errors에 기록합니다. 링크/정션 자체는 목록에 표시하되 루프 방지를 위해 안으로 따라가지 않습니다. Cloud 디렉터리는 OS reparse tag를 구분해 메타데이터를 열거합니다. 네트워크 I/O 취소는 현재 OS 호출이 반환된 후 반영될 수 있으며 창은 계속 응답합니다.

날짜는 해당 PC의 현지 시간입니다. Excel이 날짜로 표시할 수 없는 1900년 이전 값은 문자로 보존합니다. Excel 셀의 32,767자 제한을 넘는 문자열은 제한 길이로 잘립니다. 수식처럼 보이는 파일명도 수식이 아닌 문자로 저장합니다. 빈 폴더도 헤더가 있는 정상 통합문서를 만듭니다.

## CLI

    FileListToExcel.exe --files "C:/자료/보고서.docx" "C:/자료/결과.pdf"
    FileListToExcel.exe --folder "C:/자료" "D:/자료"
    FileListToExcel.exe --recursive "C:/자료"
    FileListToExcel.exe --recursive "C:/자료" --no-open --output "C:/결과/목록.xlsx"
    FileListToExcel.exe --matches "C:/자료/기준.txt"
    FileListToExcel.exe --duplicates "C:/자료" "D:/백업"
    FileListToExcel.exe --duplicate-files "C:/자료/원본.txt" "D:/백업/복사본.txt"
    FileListToExcel.exe --duplicates "C:/자료" --no-cache --no-open --output "C:/결과/중복.xlsx"

--no-open은 진행 창과 자동 열기를 사용하지 않는 자동화 모드입니다. --output은 기존 파일을 덮어쓰지 않습니다. 종료 코드: 0 생성 성공(Errors 시트의 개별 오류 포함), 1 실행 실패, 2 취소. 셸 연동의 --request는 제품의 GUID 요청 파일만 수용합니다.

`--no-cache`는 중복 검사에서 캐시 조회와 저장을 생략하는 진단 옵션입니다. 일반 우클릭 실행에서는 옵션 창이 추가되지 않습니다.

## 개발 및 검증

필수: Windows x64, .NET SDK 10.0.401, Visual Studio 2022 C++ Build Tools + Windows SDK + CMake. 로컬 .tools/dotnet에 있는 동일 SDK도 사용합니다. 네이티브 빌드는 선택적으로 LLVM-MinGW를 지원합니다.

    dotnet test FileListToExcel.sln -c Release
    powershell -NoProfile -File scripts/Build.ps1 -TestInstaller

Build.ps1은 관리 코드 테스트, 자체 포함 앱 게시, 네이티브 셸 테스트, WiX MSI 빌드 및 SHA-256 생성을 수행합니다. -TestInstaller는 기존 설치가 없는 환경에서 설치·복구·제거를 검증합니다. 결과는 artifacts/release에 생성됩니다. CI는 Windows에서 같은 검증을 실행하고 main에 포함된 v* 태그만 릴리스합니다.

명세: [FileListToExcel_SPEC_v1.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/FileListToExcel_SPEC_v1.md). 설계: [docs/ARCHITECTURE.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/ARCHITECTURE.md). 검증 상태: [docs/VALIDATION.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/VALIDATION.md).

중복 기능 명세: [작업지시서](FileListToExcel_DuplicateFinder_WorkOrder.md). [구현 보고서](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/DUPLICATE_FINDER_REPORT.md), [Milestone별 검증](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/DUPLICATE_FINDER_VALIDATION.md), [성능 측정 재현](tests/FileListToExcel.Benchmarks/README.md).

## 배포 상태

MSI와 실행 파일의 Authenticode 서명은 게시자 인증서를 별도로 연결해야 합니다. 서명되지 않은 빌드에는 Windows의 알 수 없는 게시자/SmartScreen 경고가 표시될 수 있습니다. 체크섬은 파일 무결성 확인용이며 게시자 서명을 대체하지 않습니다. 지원 환경별 검증 범위는 검증 문서를 참고하세요.
