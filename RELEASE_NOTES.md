# File List to Excel v1.1.0

기존 파일 목록 내보내기에 SHA-256 기반 중복 파일 찾기를 추가했습니다.

- 파일 하나: 같은 폴더와 하위 폴더에서 동일한 파일 찾기
- 파일 여러 개: 선택한 파일 안에서 중복 찾기
- 폴더 선택 또는 폴더 빈 공간: 하위 폴더를 포함해 중복 찾기
- 파일 크기 → Quick Fingerprint → 후보만 SHA-256 검사, SQLite Hash Cache 재사용
- Duplicates / Matches, Summary, Errors, Skipped 시트와 원본 하이퍼링크
- 단계별 진행 상황과 취소, 클라우드 전용 파일 및 재분석 지점 건너뛰기
- 기존 파일 목록 메뉴와 Excel 형식 유지, 원본 삭제·이동·수정 없음

다운로드: **FileListToExcel-1.1.0-win-x64.msi**. SHA256SUMS.txt로 무결성을 확인할 수 있습니다. Windows 11 x64의 **더 많은 옵션 표시** 메뉴를 사용하며, .NET을 포함한 사용자별 설치로 관리자 권한이 필요하지 않습니다.

5,102개 파일(총 4.52 GB, 실제 2 GiB 초과 파일 2개 포함)에서 첫 검사 44.72초, 캐시 재검사 0.87초를 측정했습니다. 크기 후보 102개 중 전체 해시는 6개만 계산했고, 재검사는 원본 내용을 읽지 않았습니다. 측정 환경과 OS 파일 캐시의 영향을 포함한 상세 조건은 [검증 기록](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/DUPLICATE_FINDER_VALIDATION.md)을 참고하세요.

캐시는 `%LOCALAPPDATA%\FileListToExcel\hash_cache.sqlite`에 저장됩니다. 파일 경로·크기·최종 수정 시각이 같으면 재사용하며, `--no-cache`로 우회할 수 있습니다. 제거 시 사용자가 만든 Excel 파일과 캐시는 보존됩니다.

보수적인 정책으로 모든 파일·폴더 재분석 지점(일부 로컬 OneDrive 파일 포함)을 건너뜁니다. 실제 OneDrive cloud-only 계정, 원격 SMB/외장 SSD 및 Excel 미설치 PC는 이번 검증에 포함되지 않았습니다. MSI/실행 파일은 Authenticode 서명되지 않았고 ARM64 탐색기는 지원하지 않습니다.

설치·사용법은 [README](https://github.com/prozac0401/File-List-To-Excel/blob/v1.1.0/README.md), 구현·테스트·제한사항은 [완료 보고서](https://github.com/prozac0401/File-List-To-Excel/blob/main/docs/DUPLICATE_FINDER_REPORT.md)를 참고하세요.

최종 Windows CI에서 기존 기능, 설치·손상 등록 복구·제거, 실제 1.0.0 배포본에서의 업그레이드가 통과했습니다. 로컬 테스트 PC에서는 파일 우클릭 handler 등록 조회 문제가 재현되어 원인이 미해결입니다(폴더/배경 등록은 확인). 이 환경의 실패와 진단 결과를 완료 보고서에 명시했으며, 테스트 기준을 완화하지 않았습니다.
