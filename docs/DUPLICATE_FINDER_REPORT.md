# File List to Excel 중복 파일 찾기 구현·검증 보고서

대상 버전: **1.1.0** · 검증일: **2026-09-24**

작업 명세: `FileListToExcel_DuplicateFinder_WorkOrder.md` 전체 25절. 이 보고서는 명세 24절의 아홉 항목을 따릅니다.

**현재 상태:** Milestone 1–6 구현·검증 완료. **[M7_PENDING]** 최종 MSI 설치·업그레이드·제거 검증 진행 중. **[PUBLISH_PENDING]** GitHub Main 반영 및 GitHub Release/MSI 공개 대기. 이 두 항목은 성공 증거 확인 후 최종 결과로 갱신합니다.

## 1. 구현한 기능

- 파일 하나를 선택하면 그 파일의 현재 폴더와 하위 폴더에서 내용이 같은 파일을 찾습니다. 기준 파일은 Matches 상단에 표시하고 결과 행에서는 제외합니다.
- 여러 파일을 선택하면 선택한 파일끼리 비교합니다. 하나 또는 여러 폴더를 선택하면 하위 폴더까지 합친 범위에서 중복을 찾습니다. 폴더 빈 공간에서도 현재 폴더 중복 검사를 실행할 수 있습니다.
- 기존 Explorer 선택 전달, 메타데이터 열거, 지연 표시 Progress UI, 취소 처리, OOXML 생성기와 설치 구조를 확장했습니다. 기존 파일 목록 기능을 다시 작성하지 않았습니다.
- 결과에는 Duplicates 또는 Matches, Summary, Errors, Skipped 시트가 있습니다. 숫자 크기·Excel 날짜·Table·AutoFilter·고정 헤더·원본 하이퍼링크와 65,000개 링크 단위 분할을 적용했습니다.
- 원본 파일을 삭제·이동·수정하는 기능은 없습니다. 검사 실패는 파일별로 기록하고 나머지 검사를 계속합니다. 취소 시 완성되지 않은 Excel 결과를 공개하거나 열지 않습니다.
- **M7 진행 상태:** [M7_PENDING] 버전 1.1.0 MSI의 실제 설치·업그레이드·제거 결과는 아직 확정하지 않았습니다.
- **배포 진행 상태:** [PUBLISH_PENDING] Main 커밋 및 Release URL·MSI·SHA256SUMS 확인 결과를 배포 완료 후 추가합니다.

## 2. 변경한 파일

아래 경로는 저장소 루트 기준입니다. 사용자 제공 작업 명세도 함께 보존합니다.

**Core: 비교 엔진·안전한 읽기·캐시·Excel 출력**

```text
src/FileListToExcel.Core/DuplicateModels.cs
src/FileListToExcel.Core/DuplicateDetector.cs
src/FileListToExcel.Core/DuplicateScanService.cs
src/FileListToExcel.Core/FileContentHasher.cs
src/FileListToExcel.Core/SqliteHashCache.cs
src/FileListToExcel.Core/DuplicateWorkbookWriter.cs
src/FileListToExcel.Core/FileListToExcel.Core.csproj
src/FileListToExcel.Core/FileScanner.cs
src/FileListToExcel.Core/Models.cs
src/FileListToExcel.Core/WorkbookWriter.cs
```

**앱·Explorer: 선택 모드, 요청 전달, 실행, 진행 상황**

```text
src/FileListToExcel.App/CommandLine.cs
src/FileListToExcel.App/Program.cs
src/FileListToExcel.App/ProgressWindow.cs
src/FileListToExcel.Shell/ShellExtension.cpp
src/FileListToExcel.Shell/CMakeLists.txt
src/FileListToExcel.Shell/Version.rc
src/FileListToExcel.Shell/README.md
src/FileListToExcel.Shell/tests/ShellTests.cpp
```

**자동 검증·재현 가능한 성능 측정 도구**

```text
tests/FileListToExcel.Core.Tests/DuplicateEngineTests.cs
tests/FileListToExcel.Core.Tests/DuplicateScanTests.cs
tests/FileListToExcel.Core.Tests/SafeContentReaderTests.cs
tests/FileListToExcel.Core.Tests/HashCacheTests.cs
tests/FileListToExcel.Core.Tests/DuplicateCacheIntegrationTests.cs
tests/FileListToExcel.Core.Tests/CacheSidecarSafetyTests.cs
tests/FileListToExcel.Core.Tests/DuplicateWorkbookTests.cs
tests/FileListToExcel.Core.Tests/DuplicateSafetyRegressionTests.cs
tests/FileListToExcel.App.Tests/DuplicateCommandLineTests.cs
tests/FileListToExcel.Benchmarks/FileListToExcel.Benchmarks.csproj
tests/FileListToExcel.Benchmarks/Program.cs
tests/FileListToExcel.Benchmarks/README.md
```

**빌드·설치·업그레이드·CI**

```text
Directory.Build.props
.gitattributes
.github/workflows/windows.yml
installer/Product.wxs
installer/README.md
scripts/Build.ps1
scripts/Test-Installer.ps1
scripts/Test-Upgrade.ps1
scripts/Test-InstalledDuplicates.ps1
```

**문서·명세·새 SQLite 의존성의 라이선스 원문**

```text
FileListToExcel_DuplicateFinder_WorkOrder.md
README.md
RELEASE_NOTES.md
THIRD-PARTY-NOTICES.md
docs/ARCHITECTURE.md
docs/VALIDATION.md
docs/DUPLICATE_FINDER_VALIDATION.md
docs/DUPLICATE_FINDER_REPORT.md
docs/licenses/sqlite/README.md
docs/licenses/sqlite/Microsoft.Data.Sqlite-LICENSE.txt
docs/licenses/sqlite/SQLitePCLRaw-LICENSE.txt
docs/licenses/sqlite/SQLitePCLRaw-NOTICE.txt
docs/licenses/sqlite/SQLite-PUBLIC-DOMAIN.html
```

## 3. 중복 탐지 알고리즘

1. 파일 크기가 같은 파일이 두 개 이상인 그룹만 남깁니다. 기준 파일 찾기는 기준 파일과 같은 크기의 후보로 범위를 좁힙니다.
2. 1 MiB 이하 후보는 전체 SHA-256을 한 번 계산합니다. 1 MiB 초과 후보는 파일 크기의 64비트 표현과 시작·중앙·끝의 각 64 KiB를 SHA-256으로 계산하여 Quick Fingerprint를 만듭니다.
3. 크기와 Quick Fingerprint가 모두 같은 후보만 전체 SHA-256으로 확인합니다. 기준 파일 찾기는 기준 파일과 같은 지문 그룹만 이 단계로 진행합니다.
4. 같은 크기·최종 SHA-256별로 그룹화하고 예상 중복 용량 내림차순, 그룹 및 상대경로 순으로 결과를 표시합니다. 빈 파일은 내용 읽기 없이 묶습니다.

전체 내용은 4 MiB 버퍼로 스트리밍합니다. 일반 경로는 최대 두 worker, UNC 경로는 한 worker를 사용합니다. 검사 전후의 크기·UTC 수정 시각을 비교하고, 열린 Windows 파일 핸들의 메타데이터도 확인합니다. 변경된 파일은 ChangedDuringScan으로 제외합니다.

cloud-only 표시 및 모든 Reparse Point를 보수적으로 제외합니다. 파일 읽기는 Windows의 읽기 전용 핸들, OPEN_REPARSE_POINT, OPEN_NO_RECALL을 사용하며 이미 열린 핸들의 속성을 첫 내용 읽기 전에 재확인합니다. 폴더와 상위 경로도 검사하여 Junction/Symbolic Link 재귀를 막습니다.

## 4. Cache 방식

- 저장소는 SQLite이며 기본 위치는 Windows LocalAppData 아래 `FileListToExcel/hash_cache.sqlite`입니다. Microsoft.Data.Sqlite 10.0.12를 사용합니다.
- 조회 기준은 정규화된 절대 경로, 파일 크기, 최종 수정 시각의 UTC ticks, 지문/해시 알고리즘 버전입니다. Windows 타임스탬프 정밀도인 100 ns를 유지합니다.
- Quick Fingerprint, 최종 SHA-256, 마지막 확인 시각을 저장합니다. 같은 경로의 크기나 시각이 달라지면 이전 전체 해시를 새 버전으로 이어받지 않습니다.
- WAL과 원자적 upsert를 사용하고 DB 접근을 직렬화합니다. 잠금 대기는 1초로 제한하며 캐시 사용 실패 시 원본 검사로 계속합니다.
- 손상 또는 잘못된 스키마는 기존 파일을 격리 보존한 뒤 재생성을 시도합니다. 더 높은 버전의 스키마는 그대로 두고 해당 실행에서 캐시를 사용하지 않습니다.
- 캐시 DB·상위 경로·기존 WAL/SHM/journal의 cloud/reparse 속성을 검사합니다. 내부 DB와 관련 파일은 검사 대상에서 제외합니다. `--no-cache`는 진단용 우회 옵션입니다.

캐시는 메타데이터에 근거한 최적화입니다. 파일 내용이 바뀌어도 크기와 수정 시각이 의도적으로 유지되거나 파일 시스템이 변경 시각을 아직 확정하지 않았다면 이전 해시가 재사용될 수 있습니다. 원자적인 파일 시스템 스냅샷이나 변조 검출을 제공하지 않습니다.

## 5. Explorer 메뉴 변화

| 선택 상태 | 새 메뉴 | 검사 범위 |
| --- | --- | --- |
| 파일 하나 | 같은 파일 찾기 | 기준 파일의 현재 폴더와 하위 폴더 |
| 파일 두 개 이상 | 선택 파일 중 중복 찾기 | 선택한 파일만 |
| 폴더 하나 이상 | 중복 파일 찾기 | 선택 폴더들을 합친 전체 하위 트리 |
| 폴더 내부 빈 공간 | 현재 폴더 중복 파일 찾기 | 현재 폴더와 하위 폴더 |

기존 파일 목록 메뉴와 전달 방식을 유지했습니다. 혼합 선택 등 지원하지 않는 선택에서는 해당 중복 메뉴를 노출하지 않습니다. Windows 11에서는 기존 클래식 우클릭 메뉴, 즉 **더 많은 옵션 표시** 경로를 사용합니다. 별도 상주 서비스는 추가하지 않았습니다.

## 6. 테스트한 케이스

| 범위 | 실제 결과 | 주요 검증 |
| --- | --- | --- |
| 기존 기준선 | Core 26, App 9, native 159 통과 | 구현 전 정상 상태 확인 |
| 최종 Core Release | **90/90 통과, 실패·건너뜀 0** | 기존 26개 포함, 새 캐시 sidecar 6개까지 포함 |
| 최종 App Release | **31/31 통과, 실패·건너뜀 0** | 기존 9개 포함, 모드·인수·일회성 요청 전달 |
| native Explorer 통합 | **414개 검사 통과** | 표시 조건, 요청 전달, 선택 수, Unicode/ANSI verb |
| 최종 MSI 설치·업그레이드·제거 | **[M7_PENDING] 진행 중** | 설치 패키지에 대한 최종 성공 결과는 아직 기록하지 않음 |

실제 파일을 사용하여 같은 내용/다른 이름·확장자, 같은 크기/다른 내용, 다른 크기, 0 byte, 1 MiB 경계, Quick Fingerprint 샘플 밖의 변경, 한글·emoji·긴 경로, 10단계 하위 폴더, 빈 폴더, 여러/겹치는 폴더, 실제 Junction 및 ACL 접근 거부, 읽기 잠금, 파일 삭제·수정 중 검사 등을 확인했습니다.

캐시는 재실행, 같은 크기 수정, 크기 변경, DB 삭제·손상·스키마 오류, 접근 불가, 실제 쓰기 잠금, 병렬 호출, 취소 및 검사 중 변경에 대한 결과를 확인했습니다. 기존 cache sidecar에 실제 Offline 속성 또는 디렉터리가 있으면 DB를 만들거나 원본 내용을 쓰지 않고 캐시 사용을 중단하는 여섯 사례도 통과했습니다.

Excel 출력은 Open XML 스키마, 숫자/날짜, Table/AutoFilter/고정 헤더, Unicode·#·%·emoji 하이퍼링크, 기준 파일 제외, 오류/건너뜀, 빈 결과, 65,001행 분할, 취소 시 임시 파일 정리, 기존 출력 덮어쓰기 방지를 검증했습니다. 실제 최신 helper는 세 중복 모드와 기존 `--folder`에서 각각 4·1·2·5행의 결과를 만들고 모두 종료 코드 0을 반환했습니다.

실제 Excel Desktop에서 Duplicates와 Matches가 복구 경고 없이 열렸고 Summary와 기준 파일 표시를 확인했습니다. 이미 Excel이 열린 상태에서 helper가 결과를 열고 종료했습니다. 특수 문자가 포함된 하이퍼링크를 눌러 예상한 원본 내용을 Notepad에서 확인했습니다.

진행 UI의 별도 취소 검증에서는 두 개의 16 GiB **sparse 파일**을 사용했습니다. 전체 SHA-256 단계에서 UI가 응답했고 취소 후 helper 종료 코드 2, 완성 결과·partial 파일 없음, helper 잔류 없음이 확인되었습니다. 이 sparse 데이터는 다음 성능 수치에 포함하지 않았습니다.

증거: `artifacts/duplicate-final-tests/core-final.trx`, 최종 App의 자동 생성 TRX, `artifacts/test-results/native-duplicates.xml`, `artifacts/milestone5-cli.json`, `artifacts/milestone6-ui.json`. 단계별 기록은 [검증 문서](DUPLICATE_FINDER_VALIDATION.md)에 있습니다.

## 7. 실제 성능 측정 결과

환경: DT-PROZAC, Windows 11 build 22631, .NET 10.0.12, 로컬 볼륨. 별도 실행 도구 `tests/FileListToExcel.Benchmarks`로 측정했습니다. 증거는 `artifacts/duplicate-performance.json`입니다.

- 전체 **5,102개 파일, 4,517,316,068 bytes**: 크기가 모두 다른 작은 파일 5,000개, 2 MiB 파일 100개, 내용이 동일한 **2,147,549,184-byte 파일 두 개**입니다.
- 대용량 두 파일은 각각 2 GiB + 64 KiB이며 전체 내용을 실제로 기록했습니다. Sparse/Compressed 속성이 없음을 확인했습니다.
- 2 MiB 집합에는 실제 동일 파일과 일부 샘플만 같은 비동일 파일을 포함했습니다.
- 데이터 준비 144.02초는 아래 검사 시간에서 제외했습니다. 시간은 OS 파일 캐시의 영향을 포함하므로 물리 디스크의 완전히 비어 있는 캐시 상태를 보장하지 않습니다.

| 측정값 | 첫 검사 | 캐시 재검사 |
| --- | ---: | ---: |
| 열거를 포함한 전체 시간 | **44,715.53 ms** | **869.37 ms** |
| 비교 엔진 시간 | 42,462.65 ms | 382.91 ms |
| 크기 후보 → Quick Fingerprint 통과 후보 | 102 → 6 | 102 → 6 |
| Quick / 전체 SHA-256 계산 횟수 | 102 / 6 | 0 / 0 |
| 캐시 재사용 파일 수 | 0 | **102** |
| 원본 내용 읽기량 | 4,323,540,992 bytes | **0 bytes** |
| 샘플링한 최대 프로세스 Working Set | 74,141,696 bytes | 75,517,952 bytes |

두 실행 모두 **중복 그룹 2개, 그룹 내 파일 4개**를 찾았습니다. 한 파일만 남긴다고 가정한 계산값은 **2,149,646,336 bytes**이며 삭제 권고가 아닙니다. 첫 검사에서 동시에 활성화된 내용 읽기 작업은 최대 두 개였습니다. 캐시 실행은 102개 후보 모두의 저장된 해시를 재사용했습니다.

검사 전후 5,102개 원본의 크기·수정 시각·속성이 동일했습니다. 열거 취소는 128개 항목 뒤 중단했으며 **전체 소요 시간은 10.52 ms**였습니다. 두 대용량 전체 해시가 실행되는 동안 취소했을 때 **취소 요청 이후 24.21 ms**에 종료했고, 결과를 반환하지 않았으며 파일 핸들이 해제되었음을 배타적 재열기로 확인했습니다.

## 8. 아직 남아 있는 제한사항

- 실제 OneDrive cloud-only 계정, 실제 원격 SMB/네트워크 드라이브, 외장 SSD, Excel 미설치 PC에서는 검증하지 않았습니다. 테스트한 Offline 속성 및 안전한 파일 열기 검증을 실제 cloud provider의 종단 간 검사로 간주하지 않습니다.
- 이 PC는 Symbolic Link 생성 권한이 없어 실제 symlink 생성 시험을 완료하지 못했습니다. 실제 Junction과 Reparse Point 방지 검증은 수행했습니다.
- 모든 Reparse Point를 건너뛰므로 로컬에 내려받은 OneDrive 파일·폴더도 해당 속성이 남아 있으면 제외될 수 있습니다.
- 파일 시스템 스냅샷을 사용하지 않습니다. 계속 수정되는 트리에서는 검사 도중 변경된 파일을 제외하며, 동일한 크기·수정 시각을 유지한 내용 변경은 캐시가 감지하지 못할 수 있습니다.
- Excel 결과의 이론적 중복 용량은 명세의 `(파일 수 - 1) × 파일 크기` 계산값입니다. 자동 삭제·이동·최신 파일 선택이나 유사한 내용 검색은 제공하지 않습니다.
- Windows x64 클래식 Explorer 메뉴를 대상으로 합니다. ARM64 Explorer 및 Windows 11 현대식 메뉴 직접 통합은 지원 범위 밖입니다. MSI/실행 파일은 Authenticode 서명되지 않았습니다.
- **[M7_PENDING] / [PUBLISH_PENDING]** 최종 설치 패키지 검증과 Main/Release 공개는 아직 이 보고서의 완료 항목으로 표시하지 않았습니다.

## 9. 수동으로 확인해야 할 항목

- **[M7_PENDING]** 최종 MSI 실제 설치 후 기존 파일 목록과 새 중복 명령 실행, 1.0.0 → 1.1.0 업그레이드, 제거 후 Explorer 등록·메뉴 및 helper 잔류 확인. 자동 설치 시험과 실제 확인 결과를 받아 최종 상태를 갱신합니다.
- **[PUBLISH_PENDING]** GitHub Main 커밋, v1.1.0 Release URL, MSI 파일 및 SHA256SUMS의 공개·다운로드 가능 여부.
- 실제 OneDrive cloud-only 파일과 폴더를 대상으로 실행 전후 다운로드 상태·네트워크 사용량이 변하지 않는지 확인합니다.
- 권한이 있는 별도 환경에서 파일/폴더 Symbolic Link와 순환 경로를 확인합니다. 실제 SMB 공유 및 외장 SSD에서는 권한·연결 중단·취소를 점검합니다.
- Excel이 설치되지 않은 별도 PC에서 안내/파일 탐색기 대체 동작을 확인합니다. 해당 환경을 현재 PC에서 흉내 낸 결과를 실제 미설치 검증으로 대체하지 않았습니다.

원본을 정리하거나 삭제하는 후속 동작은 수행하지 않았습니다. 성능/UI 검증용으로 생성한 대용량 fixture는 별도의 artifacts/benchmarks 경로에 보관되어 있습니다.
