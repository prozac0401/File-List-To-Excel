# File Collect 검증 — v1.2.0

검증일: 2026-09-26. 정식 릴리즈는 아래 자동 게이트를 통과한 태그 빌드의 MSI만 게시한다. 이전 평가 후보의 실패 기록은 [평가 이력](FILE_COLLECT_EVALUATION_HISTORY.md)에 보존한다.

## 자동 릴리즈 게이트

소스 후보 `280369160877fecd7a4350d18281036ac46ccd0d`의 [Windows CI 36221029040](https://github.com/prozac0401/File-List-To-Excel/actions/runs/36221029040)과 [PR CI 36221067350](https://github.com/prozac0401/File-List-To-Excel/actions/runs/36221067350)가 모두 성공했다.

| 검사 | 결과 |
| --- | --- |
| 관리 코드 Release | Core 138 + App 47 = **185개 통과**, 실패/건너뜀 0 |
| 네이티브 | 셸, Excel x64, Excel x86의 CTest 3개 모음 모두 통과 |
| 추가 선택 회귀 | 실제 Capture 경로의 다중 영역 중복 제거·행 순서·숨김 행·전체 시트 교차·헤더/관계없는 범위·손상 기술 정보·10,000행 제한 등 11개 시나리오 |
| 개발용 Excel 수용 도구 | Release 빌드 성공; CI에서는 Office를 실행하지 않음 |
| MSI | WiX ICE 검증, 경고를 오류로 처리 |
| 기본 본체만 설치 | 설치·원본 DLL의 253개 셸 검사·목록/중복/Matches/선택 파일·캐시 재사용·손상 등록 복구·제거·산출물 보존 통과 |
| Excel 포함 설치 | x86/x64 COM 등록, 최초 LoadBehavior=3, 복구 시 비활성 값 2 유지, 제거 후 등록/파일 부재까지 통과 |
| 업그레이드 | 합성 0.9.0→1.2.0 및 실제 공개 1.1.0→1.2.0 모두 통과. ProductCode 교체, Excel 자동 추가 방지, 기존 기능/캐시/산출물 보존 확인 |

공개 1.1.0 시험 MSI는 SHA-256 `14af126d51684e61f39a334db63a572c69926a016d043d217e8866ee05a33ab0`으로 검증한다. Excel 포함 복구 로그는 AppSearch의 `#2`, PreserveExcelLoad2 실행, `RegAddValue LoadBehavior=#2`를 모두 기록한다. 이전 평가 환경의 모순된 조회 결과로 합격을 추정한 것이 아니다.

빌드는 BUILD_TESTING=ON을 강제하고 CTest에 --no-tests=error를 사용한다. 설치 검증 스크립트가 exit 1로 종료해도 호출 빌드가 즉시 실패하도록 종료 코드를 검사한다. 실제 PowerShell 격리 프로브에서 exit 1/throw의 후속 실행 차단과 exit 0/정상 반환의 통과를 확인했다. 진단용 등록 실패 계속 옵션을 정식 CI에서는 사용하지 않는다. 설치 라이선스 화면은 저장소 LICENSE의 기존 권리 유보 조건과 동일하며 새 오픈소스 라이선스를 부여하지 않는다.

## 실제 Excel과 로컬 실행 환경

CI 후보 MSI는 50,880,300 bytes, SHA-256 `42cfc5f6d33513471122e75dffaf132ae0b26d11886eac8092a79f66def2c5b3`다. 이 후보를 설치한 실제 Excel x64 16.0.20326.20158에서 아래 결과를 확인했다. 최종 배포 MSI의 식별 정보는 GitHub Release의 SHA256SUMS.txt를 기준으로 한다.


| 실제 Excel 수용 | 결과 |
| --- | --- |
| 새 정상 Excel 시작 | 추가 기능 Connected=true, EnableEvents=true, Cell/List Range Popup 메뉴 각각 1개 |
| 미저장 필터 | 100행을 7행으로 필터 후 A2:A101 선택; 실제 셀 우클릭 메뉴로 실행 |
| 목적지·실제 복사 | 7행/322 bytes 표시. 새 하위 폴더에 f001~f007만 복사; 7개 모두 SHA-256 일치 |
| 원본/보고서 경계 | 100개 원본 내용 무변경. 전달 폴더에는 복사 파일 7개만 존재 |
| 상세 결과 UI | 순번 1~7, 한국어 복사 완료와 Copied 상태; 결과창 정상 닫기 |
| 정상 종료 | helper 종료, Excel Quit 반환 후 12,252ms에 실제 프로세스 종료 확인(exit0); 강제 종료 없음 |
| 기존 기능 | 설치된 셸 DLL 253개 검사 및 새 100행 목록·중복 CLI 정상 |

자동화 도구가 직접 시작한 프로세스에서는 레지스트리의 파일 메뉴/Excel 등록이 보이지 않는 현상을 재현했다. 같은 사용자 SID와 같은 대화형 세션의 독립 프로세스에서는 동일 설치의 LoadBehavior=3과 올바른 파일 메뉴 CLSID를 직접 조회했다. 그 실행 경로에서 Excel의 자동 로드와 위 복사를 통과했다. 따라서 이번 후보의 설치 등록 누락으로 해석하지 않는다. 정확한 내부 리디렉션 기작은 단정하지 않으며 보안 정책·실행 정책·등록 값을 변경하지 않았다. 최초 도구 경로의 실패 기록도 별도로 보존한다.

동일한 일반 프로세스 실행 경로에서 **로컬 strict OFF·ON 설치/복구/제거와 실제 1.1.0→1.2.0 업그레이드도 각각 종료 코드 0으로 통과**했다. 등록 실패 계속 옵션은 사용하지 않았으며 OFF/ON 각각 설치된 셸 253개 검사, 비활성 LoadBehavior=2의 복구 보존, 제거 후 등록/설치 폴더/프로세스 부재와 원본·통합문서·캐시 보존을 확인했다. 실행 상태는 artifacts/release-1.2-gates/ordinary-strict-gates.json에 기록했다.

이전 평가에서 실제 Excel x64 16.0.20326.20158의 자동 연결, 100→7 필터 복사, 숨김/중복 선택, 정렬·열 이동·시트 이름 변경·Save As, 두 Excel 프로세스의 별도 복사, Matches 결과창 표시/정상 닫기를 확인했다. 해당 후보별 MSI와 결과를 [평가 이력](FILE_COLLECT_EVALUATION_HISTORY.md)에 구분했다. 이 이력을 새 MSI 자체의 검증으로 바꾸어 기록하지 않는다.

## 미실행 환경과 지원 경계

- 실제 Excel x86, 깨끗한 로컬 Windows 사용자/VM의 실제 Office, 재로그인/재부팅 후 자동 로드. x86 네이티브 COM 시험은 실제 x86 Office 시험과 다르다.
- 실제 원격 SMB의 지연·끊김, cloud-only provider/hydration, 디스크 부족, EFS/DRM, 조직 서명/차단 정책 및 다른 제품의 모든 설치·제거 공존 조합.
- 실제 Excel Files_2의 모든 선택 조합, totals-only/복수 표/손상 정보 전체 행렬, 동일 조건 목록 생성 성능 전후 측정.
- 보이는 선택 데이터 행 10,000개 및 요청 32 MiB 한도. 고유 파일 수보다 엄격하며 기존 목록 생성 한도와 별개다.
- 모든 reparse/recall 경유 경로는 보수적으로 제외한다. 내려받은 OneDrive 파일도 해당될 수 있다. 목적지는 ADS와 영구 ACL을 지원해야 하며 원본 소유자/ACL 자체를 복제하지 않는다.
- 파일 경로·크기·수정시각은 과거 바이트의 인증이 아니며, 제품 메타정보는 출처 인증이 아니다.
- Authenticode 서명 없음. ARM64 탐색기 미지원. 조직 정책이나 보안 설정을 완화하지 않는다.

## 재현 및 증거

`dotnet test FileListToExcel.sln -c Release` 및 깨끗한 Windows 계정의 `powershell -NoProfile -File scripts/Build.ps1 -Version 1.2.0 -TestInstaller`를 사용한다. 실제 1.1.0 배포본 업그레이드는 GitHub workflow에 SHA-256과 함께 고정되어 있다.

GitHub 각 실행의 windows-test-evidence에는 TRX, 네이티브 JUnit/CTest, 설치/복구/제거 및 업그레이드 로그가 포함된다. 로컬 상세 증거는 artifacts/validation/release-1.2.0와 artifacts/release-1.2-gates에 보관하며 공개 릴리즈에는 MSI와 체크섬만 첨부한다.
