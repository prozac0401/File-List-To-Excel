# File List to Excel

Windows 탐색기에서 선택한 파일과 폴더의 목록을 Excel 표로 여는 Windows 11 x64 도구입니다. 설정 창, 계정, 상주 서비스 없이 실행합니다.

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
- 파일 내용은 읽지 않습니다. 다운로드를 요청하거나 해시·문서 분석을 실행하지 않습니다.
- 읽기 실패는 Errors 시트에 기록하고 계속합니다. 원본 파일을 변경하거나 삭제하지 않습니다.

생성 파일은 %TEMP%/FileListToExcel에 저장합니다. 사용자가 저장한 결과와 임시 Excel 파일은 제거 시 삭제하지 않습니다. Windows의 임시 파일 정리 대상이 될 수 있으므로 보관할 결과는 별도 저장하세요. 진단 오류는 %LOCALAPPDATA%/FileListToExcel/Logs/last-error.log에 저장하며 네트워크로 전송하지 않습니다.

## 대량 및 예외 처리

한 번에 선택해 전달할 수 있는 입력 경로는 최대 100,000개이며, 셸 요청 파일은 32MiB까지 지원합니다. 폴더 재귀 탐색 결과 행 수에는 이 입력 개수 제한이 적용되지 않습니다.

Excel 링크 제한을 피하기 위해 시트당 65,000행을 기록하고 Files_2 등으로 자동 분할합니다. 매우 긴 링크는 전체경로를 보존하고 Errors에 기록합니다. 링크/정션 자체는 목록에 표시하되 루프 방지를 위해 안으로 따라가지 않습니다. Cloud 디렉터리는 OS reparse tag를 구분해 메타데이터를 열거합니다. 네트워크 I/O 취소는 현재 OS 호출이 반환된 후 반영될 수 있으며 창은 계속 응답합니다.

날짜는 해당 PC의 현지 시간입니다. Excel이 날짜로 표시할 수 없는 1900년 이전 값은 문자로 보존합니다. Excel 셀의 32,767자 제한을 넘는 문자열은 제한 길이로 잘립니다. 수식처럼 보이는 파일명도 수식이 아닌 문자로 저장합니다. 빈 폴더도 헤더가 있는 정상 통합문서를 만듭니다.

## CLI

    FileListToExcel.exe --files "C:/자료/보고서.docx" "C:/자료/결과.pdf"
    FileListToExcel.exe --folder "C:/자료" "D:/자료"
    FileListToExcel.exe --recursive "C:/자료"
    FileListToExcel.exe --recursive "C:/자료" --no-open --output "C:/결과/목록.xlsx"

--no-open은 진행 창과 자동 열기를 사용하지 않는 자동화 모드입니다. --output은 기존 파일을 덮어쓰지 않습니다. 종료 코드: 0 생성 성공(Errors 시트의 개별 오류 포함), 1 실행 실패, 2 취소. 셸 연동의 --request는 제품의 GUID 요청 파일만 수용합니다.

## 개발 및 검증

필수: Windows x64, .NET SDK 10.0.401, Visual Studio 2022 C++ Build Tools + Windows SDK + CMake. 로컬 .tools/dotnet에 있는 동일 SDK도 사용합니다. 네이티브 빌드는 선택적으로 LLVM-MinGW를 지원합니다.

    dotnet test FileListToExcel.sln -c Release
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts/Build.ps1 -TestInstaller

Build.ps1은 관리 코드 테스트, 자체 포함 앱 게시, 네이티브 셸 테스트, WiX MSI 빌드 및 SHA-256 생성을 수행합니다. -TestInstaller는 기존 설치가 없는 환경에서 설치·복구·제거를 검증합니다. 결과는 artifacts/release에 생성됩니다. CI는 Windows에서 같은 검증을 실행하고 main에 포함된 v* 태그만 릴리스합니다.

명세: [FileListToExcel_SPEC_v1.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/FileListToExcel_SPEC_v1.md). 설계: [docs/ARCHITECTURE.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/ARCHITECTURE.md). 검증 상태: [docs/VALIDATION.md](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/VALIDATION.md).

## 배포 상태

MSI와 실행 파일의 Authenticode 서명은 게시자 인증서를 별도로 연결해야 합니다. 서명되지 않은 빌드에는 Windows의 알 수 없는 게시자/SmartScreen 경고가 표시될 수 있습니다. 체크섬은 파일 무결성 확인용이며 게시자 서명을 대체하지 않습니다. 지원 환경별 검증 범위는 검증 문서를 참고하세요.
