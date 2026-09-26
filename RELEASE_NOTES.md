# File List to Excel v1.2.0

Excel 결과표에서 선택한 실제 파일을 새 폴더에 복사하는 **선택형 Excel 연동**을 추가했습니다. 기존 탐색기 파일 목록과 중복 찾기는 그대로 사용할 수 있습니다.

- 설치 화면에서 **Excel에서 선택한 파일 복사**를 선택한 뒤 Excel을 다시 시작하면 셀 우클릭의 **파일목록 → 선택한 파일 복사…** 메뉴가 준비됩니다. 기본값은 꺼짐이며 기존 버전의 무인 업그레이드에서 자동 추가되지 않습니다.
- 새 Files / Duplicates / Matches 결과표에서 파일명 셀만 선택해도 됩니다. 현재 Excel의 필터·정렬·Ctrl 다중 선택을 반영하고 숨긴 행은 제외합니다.
- 목적지를 한 번 선택하면 새 하위 폴더에 복사합니다. 같은 이름은 번호로 구분하며 원본 이동·삭제와 기존 파일 덮어쓰기는 하지 않습니다.
- 삭제·변경된 파일, cloud-only/recall 및 reparse 경유 경로는 제외합니다. 취소 시 완료본은 유지하고 자체 미완료본만 정리합니다.
- 상세 결과는 제품 LocalAppData의 Reports에 보관하며 전달 폴더에는 복사한 파일만 들어갑니다. 완료 화면에서 작업 결과를 확인할 수 있습니다.
- Excel x86/x64용 네이티브 추가 기능을 포함합니다. VBA 실행, 매크로 허용 또는 수동 COM 등록은 필요하지 않습니다.

**다운로드:** FileListToExcel-1.2.0-win-x64.msi 및 SHA256SUMS.txt. Windows 11 x64 사용자별 설치이며 관리자 권한과 별도 .NET 설치가 필요하지 않습니다. Excel 연동에는 Excel Desktop과 이번 버전에서 새로 만든 결과표가 필요합니다.

릴리즈 워크플로는 관리 코드·네이티브 테스트, MSI 검증, 기본/Excel 포함 설치·복구·제거, 시험용 이전 버전 및 실제 1.1.0 배포본 업그레이드를 모두 통과한 MSI만 게시합니다. 실제 Excel x64의 자동 연결·필터 선택 복사·결과 화면도 별도로 검증했습니다. 범위와 상세 결과는 [1.2.0 검증 기록](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/FILE_COLLECT_VALIDATION.md)을 참고하세요.

**제한:** 선택한 보이는 데이터 행은 최대 10,000개, 요청은 32 MiB입니다. 모든 reparse 경유 경로를 보수적으로 제외하므로 로컬 OneDrive 파일도 제외될 수 있습니다. 실제 Excel x86, 조직 보안 정책, 재부팅/재로그인, 원격 SMB 장애와 다른 추가 기능의 모든 공존 조합은 검증하지 않았습니다. MSI와 실행 파일은 Authenticode 서명되지 않았으며 ARM64 탐색기는 지원하지 않습니다.

[설치·사용 안내](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/README.md) · [구현 보고](https://github.com/prozac0401/File-List-To-Excel/blob/v1.2.0/docs/FILE_COLLECT_COMPLETION.md)
