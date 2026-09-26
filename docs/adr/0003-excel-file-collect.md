# ADR 0003 — 선택형 Excel 파일 복사 연동

상태: 구현 및 환경별 검증 중 · 2026-09-24

## 범위

사용자가 승인한 FileCollect 작업지시서 v1.0에 따라 기존 v1 명세의 “Excel Add-in 설치 금지”를 **선택형 Excel 파일 복사 구성 요소에 한해서** 변경한다. 목록/중복 찾기 본체는 Excel 추가 기능 없이 실행하며, 결과는 실행 매크로가 없는 일반 .xlsx이다. 명단대조기 및 Workspace에 의존하지 않는다.

## 결정

- 기존 네이티브 C++ 도구 체계로 IDTExtensibility2 COM 추가 기능을 컴파일한다. 자체 포함 x64 .NET helper를 재사용한다. VSTO는 추가 런타임/배포 신뢰 전제 조건이 필요하므로 기본 경로로 택하지 않는다.
- 추가 기능은 Excel 시작 시 메뉴와 컴파일된 콜백만 준비한다. 사용자 클릭 전 파일/네트워크 검사와 복사를 수행하지 않는다. 연결된 현재 Excel Application의 메모리 상태만 캡처한다.
- 별도 CLSID/ProgID/고유 메뉴 Tag를 사용한다. 제품 컨트롤과 이벤트만 해제하며 기본 메뉴 Reset, VBA OnAction 매크로, GetActiveObject 기반 인스턴스 추정은 사용하지 않는다.
- 기존 UpgradeCode와 탐색기 등록 소유권을 유지한다. MSI의 Excel 연동 기능은 기본 OFF이며 신규 설치/기존 무인 업그레이드에 자동 추가하지 않는다. 사용자가 한 번 선택 설치하면 정상 LoadBehavior 등록은 설치기가 처리한다.
- 사용자별 HKCU/LocalAppData 설치이며 x64 Windows의 Excel x64/x86용 DLL을 분리한다. 보안 설정, Trust Center, 정책, 신뢰 위치, VBA 접근 설정, 차단 목록을 변경하지 않는다. Office가 비활성화한 항목의 강제 재활성화 감시는 하지 않는다.
- 등록 실패, 정책 차단, 비트수 불일치, 누락 파일/런타임은 배포/지원 진단에서 구분한다. 평가 빌드는 서명되지 않았으며 서명이 필요한 조직에는 배포 승인된 서명 빌드가 필요하다.

## 파일 계약

FLT.Product / FLT.ActionSchemaVersion / FLT.WorkbookId 속성과 FLT.Table.<기존 표 이름> 역할 등록으로 스키마 1 표를 찾는다. 같은 Table의 숨김 __FLT_ItemId와 __FLT_SourceRecord는 정렬 시 함께 이동한다. SourceRecord는 itemId, kind, absolutePath, 문자열 sizeBytes/modifiedUtcTicks를 포함한다. 큰 JSON은 자르지 않고 unavailable로 기록한다.

이 식별 정보는 서명이나 출처 인증이 아니다. 표시 이름/경로, 행 식별자, JSON 구조, 버전, 실제 파일 속성과 경로를 각각 검증한다. 헤더만 비슷한 일반 표를 승격하지 않는다. 과거/미래/손상된 계약은 새 목록 생성 안내로 거절한다.

## helper와 복사

--collect-request는 기존 --request 작업들과 구별한다. 소유 Requests 폴더의 GUID JSON만 크기/행/버전/구조를 검사한 후 동일 핸들로 한 번 소비한다. 폴더 선택을 생략하는 CLI는 제공하지 않는다.

목적지 선택 화면에 사전 검증 개수/크기/제외 수를 표시하며 “여기에 복사”로 확정한다. 새로운 작업 하위 폴더에서 파일 이름 충돌을 분리하며 no-overwrite를 OS 호출에도 적용한다. 원본은 수정하지 않는다. cloud/recall/reparse/변경 파일을 제외하고 복사 전후 핸들 정보를 확인한다. 내부 원본 경로 보고서는 LocalAppData Reports에 보관하며 전달 폴더에 넣지 않는다.

## 검증 구분

M1 자동 로드가 우선 출시 조건이다. 개발용 COM 수동 연결이나 소스 단위 시험은 설치 후 자동 로드 시험을 대체하지 않는다. 실제 환경/깨끗한 프로필/재로그인/재부팅/x86/타제품 공존의 실측과 미실행 범위를 FILE_COLLECT_VALIDATION.md에서 별도 기록한다. 현재 사용자의 열린 Excel이나 저장하지 않은 통합문서는 강제 종료/자동 저장하지 않는다.
