# Excel VBA 매크로 모음 📊

Excel 업무 자동화를 위한 다양한 VBA 매크로 프로젝트 모음입니다.

## 📑 목차

- [프로젝트 구조](#-프로젝트-구조)
- [매크로 목록](#-매크로-목록)
  - [1. 볼드체 처리 매크로](#1-볼드체-처리-매크로)
  - [2. 품목코드 그룹화 매크로](#2-품목코드-그룹화-매크로)
  - [3. 엑셀 파일 병합 매크로](#3-엑셀-파일-병합-매크로)
- [설치 및 사용법](#-설치-및-사용법)
- [기술 스택](#-기술-스택)
- [버전 관리](#-버전-관리)

---

## 📂 프로젝트 구조

```
Excel_macro/
├── bold.bas                    # HTML 볼드 태그 변환 매크로
├── ItemCode_GroupBy.bas        # 품목코드 그룹화 매크로
├── Merge_sheet.bas             # 엑셀 파일 병합 매크로
├── ExcelVBA/
│   └── MyMacro.xlsm           # 매크로 통합 파일
├── personal_vba_home/          # 개인 VBA 매크로 라이브러리
│   └── src/
│       ├── Modules/           # 40개 이상의 모듈
│       └── Forms/             # 사용자 폼
└── README.md
```

---

## 🚀 매크로 목록

### 1. 볼드체 처리 매크로

**파일**: `bold.bas`

#### 📋 기능 설명
Excel 셀 내의 HTML 볼드 태그(`<b>텍스트</b>`)를 찾아 태그를 제거하고, 해당 텍스트에 실제 Excel 볼드 서식을 적용합니다.

#### ✨ 주요 기능
- HTML 태그 자동 감지 및 제거
- Excel 네이티브 볼드 서식 자동 적용
- 전체 시트 일괄 처리
- 태그가 있는 셀만 선택적 처리

#### 💡 사용 예시

**처리 전:**
```
일반텍스트<b>볼드텍스트</b>일반텍스트
```

**처리 후:**
```
일반텍스트볼드텍스트일반텍스트
            ^^^^^^^^
           (볼드 서식)
```

#### 🎯 사용 방법
1. Excel 파일 열기
2. `Alt + F11`로 VBA 편집기 열기
3. `bold.bas` 파일 가져오기 (파일 → 파일 가져오기)
4. `볼드체처리하기()` 매크로 실행

#### ⚙️ 기술 상세
- **처리 대상**: 활성 시트의 모든 사용된 셀(UsedRange)
- **알고리즘**: 
  1. 태그 위치 계산
  2. HTML 태그 제거
  3. Characters 객체로 부분 서식 적용
- **성능**: 큰 파일도 빠른 처리 (UsedRange 최적화)

---

### 2. 품목코드 그룹화 매크로

**파일**: `ItemCode_GroupBy.bas`

#### 📋 기능 설명
품목코드 앞 2자리를 기준으로 데이터를 자동 그룹화하고, 각 그룹별로 별도 시트를 생성합니다.

#### ✨ 주요 기능
- **자동 그룹화**: 품목코드 앞 2자리 기준 분류
- **시트 자동 생성**: 각 그룹별 독립 시트
- **하이퍼링크 생성**: 
  - F열: 각 행에서 해당 시트로 바로 이동
  - I열: 그룹별 품목명 인덱스 링크
- **스마트 시트명**: 각 그룹의 첫 번째 품목명을 시트명으로 사용
- **에러 처리**: 강화된 데이터 검증 및 오류 복구

#### 📊 데이터 형식

**입력 요구사항:**
```
1행: 회사명 정보 (예: "회사명 : 주식회사 썸유 / 포천창고 / 2025/10/10")
2행: 헤더 (품목코드, 품목명, 규격, 재고수량, 포천창고)
3행~: 실제 데이터
```

**입력 예시:**

| 품목코드 | 품목명 | 규격 | 재고수량 |
|---------|--------|------|----------|
| 00073 | C017(둔한파랑) | kg | 10 |
| 001005000002 | C002(연노랑) | kg | 10 |
| 001005000004 | C004(진노랑) | kg | 10 |

**생성 결과:**
- **C017(둔한파랑)** 시트: 00으로 시작하는 품목들
- **C002(연노랑)** 시트: 10으로 시작하는 품목들
- F열: 각 행별 시트 바로가기 링크
- I열: 그룹별 품목명 인덱스

#### 🎯 사용 방법
1. 데이터 준비 (위의 형식대로)
2. `Alt + F11`로 VBA 편집기 열기
3. `ItemCode_GroupBy.bas` 파일 가져오기
4. `SplitByPrefix_WithRowAndIndexLinks()` 함수 실행

#### 🛡️ 안전 기능
- **데이터 검증**: 빈 데이터, 잘못된 형식 자동 감지
- **에러 복구**: 작업 중단 시 설정 자동 복원
- **중복 처리**: 같은 시트명 자동 번호 부여
- **진행 표시**: StatusBar에 작업 진행률 표시
- **메모리 최적화**: 화면 업데이트 중지로 성능 향상

#### 📝 버전 히스토리

**v2.0 (2025-10-10) - 개선판**
- ✅ 전역 에러 처리 추가
- ✅ 데이터 유효성 검증 강화
- ✅ 중복 시트명 자동 처리
- ✅ 메모리 안전 처리
- ✅ 진행 상황 표시
- ✅ 빈 데이터 행 스킵
- ✅ 2행 헤더 지원

**v1.7 (2025-10-10) - 초기 버전**
- 기본 그룹화 기능
- 하이퍼링크 생성

---

### 3. 엑셀 파일 병합 매크로

**파일**: `Merge_sheet.bas`

#### 📋 기능 설명
선택한 폴더 내의 모든 엑셀 파일들을 하나의 워크북으로 병합합니다. 각 파일의 모든 시트를 개별 시트로 복사하여 통합 관리가 가능합니다.

#### ✨ 주요 기능
- **폴더 선택 UI**: 사용자 친화적인 폴더 선택 대화상자
- **다중 파일 처리**: 폴더 내 모든 엑셀 파일 자동 검색
- **다중 시트 지원**: 각 파일의 모든 시트를 개별 복사
- **자동 명명**: 복사된 시트 이름을 원본 파일명으로 자동 설정
- **포맷 유지**: 원본 데이터의 모든 서식과 내용 보존
- **파일 형식 지원**: .xls, .xlsx, .xlsm 등 모든 엑셀 형식

#### 💡 사용 시나리오

**활용 사례:**
- 여러 지점의 월별 보고서를 하나로 통합
- 부서별 데이터 파일을 단일 워크북으로 병합
- 프로젝트별 엑셀 파일을 통합 관리
- 다수의 클라이언트 데이터를 일괄 정리

#### 🎯 사용 방법

1. **매크로 실행**
   - `Alt + F8` 키 입력
   - `엑셀화일시트별병합` 선택 및 실행

2. **폴더 선택**
   - 대화상자에서 병합할 파일들이 있는 폴더 선택
   - "확인" 버튼 클릭

3. **자동 처리**
   - 매크로가 자동으로 모든 파일 처리
   - 진행 완료 메시지 확인

4. **결과 저장**
   - 새로 생성된 워크북 검토
   - 원하는 위치에 저장

#### ⚙️ 기술 상세

**처리 프로세스:**
```vba
1. FileDialog로 폴더 선택
2. 새 워크북 생성
3. Dir 함수로 파일 검색 (*.xls*)
4. 각 파일별 처리:
   - 파일 열기
   - 모든 시트 순회
   - UsedRange 복사
   - 시트명 설정
   - 파일 닫기
5. 기본 빈 시트 삭제
6. 완료 메시지 표시
```

**주요 변수:**
- `FolderPath`: 선택된 폴더 경로
- `wbSource`: 원본 워크북 (읽기용)
- `wbDest`: 대상 워크북 (통합 결과)
- `wsSource` / `wsDest`: 원본/대상 워크시트

**최적화 기법:**
- UsedRange를 사용한 효율적인 데이터 복사
- 중복 시트명 오류 방지 (On Error Resume Next)
- DisplayAlerts 제어로 불필요한 경고 제거

#### 📊 처리 예시

**폴더 구조:**
```
선택한폴더/
├── 1월보고서.xlsx (시트1, 시트2)
├── 2월보고서.xlsx (시트1)
└── 3월보고서.xlsx (시트1, 시트2, 시트3)
```

**생성 결과:**
```
통합워크북.xlsx
├── 1월보고서 (원본 시트1 복사)
├── 1월보고서 (원본 시트2 복사)
├── 2월보고서 (원본 시트1 복사)
├── 3월보고서 (원본 시트1 복사)
├── 3월보고서 (원본 시트2 복사)
└── 3월보고서 (원본 시트3 복사)
```

#### 🛡️ 주의사항

- **중복 시트명**: 같은 파일명이 여러 개 있을 경우 시트명이 자동으로 번호가 붙음
- **메모리**: 대용량 파일이 많을 경우 메모리 사용량에 주의
- **원본 보존**: 원본 파일은 변경되지 않고 그대로 유지됨
- **수동 저장**: 통합된 워크북은 자동 저장되지 않으므로 수동으로 저장 필요

#### 💪 권장 사용 환경

- **파일 개수**: 최대 100개 파일
- **파일 크기**: 각 파일 10MB 이하 권장
- **시트 개수**: 전체 시트 수 255개 이하 (Excel 제한)
- **처리 시간**: 파일 10개 기준 약 10-30초

#### 🔧 커스터마이징 팁

**특정 시트만 복사하려면:**
```vba
' 78행 수정 예시
For Each wsSource In wbSource.Sheets
    If wsSource.Name = "Sheet1" Then  ' 특정 시트만
        ' ... 복사 로직
    End If
Next wsSource
```

**시트명 규칙 변경:**
```vba
' 77행 수정 예시
' 파일명 + 시트명 조합
FileTitle = Left(FileName, InStrRev(FileName, ".") - 1) & "_" & wsSource.Name
```

---

## 🔧 설치 및 사용법

### 공통 설치 절차

1. **VBA 편집기 열기**
   - Excel에서 `Alt + F11` 키 입력

2. **매크로 보안 설정**
   - 파일 → 옵션 → 보안 센터 → 보안 센터 설정
   - "모든 매크로 포함" 또는 "디지털 서명된 매크로만" 선택

3. **매크로 가져오기**
   - VBA 편집기에서: 파일 → 파일 가져오기
   - 원하는 `.bas` 파일 선택
   - 또는: 파일 내용을 복사하여 새 모듈에 붙여넣기

4. **매크로 실행**
   - Excel에서 `Alt + F8` 키 입력
   - 실행할 매크로 선택
   - "실행" 버튼 클릭

### 빠른 실행 버튼 만들기

1. **개발 도구 탭 표시**
   - 파일 → 옵션 → 리본 사용자 지정
   - "개발 도구" 체크

2. **버튼 추가**
   - 개발 도구 → 삽입 → 단추(양식 컨트롤)
   - 시트에 버튼 그리기
   - 매크로 연결

---

## 🛠 기술 스택

- **언어**: VBA (Visual Basic for Applications)
- **플랫폼**: Microsoft Excel 2010 이상
- **호환성**: Windows / Mac Excel
- **라이브러리**: 
  - Scripting.Dictionary (자동 참조)
  - VBA 표준 라이브러리

---

## 📊 성능 최적화 팁

### 대용량 데이터 처리 시

```vba
' 매크로 시작 시
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' ... 작업 수행 ...

' 매크로 종료 시
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```

### 권장 데이터 크기
- **볼드체 처리**: 최대 10만 셀 (대부분 1초 이내)
- **품목코드 그룹화**: 최대 5만 행 (약 30초)
- **엑셀 파일 병합**: 최대 100개 파일, 각 10MB 이하 (약 1-2분)

---

## 🔍 문제 해결

### 일반적인 오류

| 오류 | 원인 | 해결 방법 |
|------|------|-----------|
| "매크로를 찾을 수 없음" | 파일이 제대로 로드 안됨 | VBA 편집기에서 파일 확인 |
| "개체 변수가 설정되지 않음" | 시트/범위가 없음 | 데이터 존재 여부 확인 |
| "시트를 만들 수 없음" | 시트명 중복/금지문자 | 기존 시트 삭제 또는 이름 변경 |
| "메모리 부족" | 데이터가 너무 큼 | 데이터를 나눠서 처리 |

### 디버깅 모드

1. VBA 편집기에서 `F8` 키로 한 줄씩 실행
2. 변수 값 확인: 디버그 → 조사식 추가
3. 중단점 설정: 코드 줄 클릭

---

## 📚 추가 리소스

### 학습 자료
- [Microsoft VBA 공식 문서](https://docs.microsoft.com/ko-kr/office/vba/api/overview/excel)
- [Excel VBA 튜토리얼](https://www.excel-easy.com/vba.html)

### 커뮤니티
- [Stack Overflow - Excel VBA 태그](https://stackoverflow.com/questions/tagged/excel-vba)
- [Reddit - r/vba](https://www.reddit.com/r/vba/)

---

## 🤝 기여하기

### 버그 리포트
이슈를 등록할 때 다음 정보를 포함해 주세요:
- Excel 버전
- 오류 메시지 (스크린샷)
- 재현 방법
- 샘플 데이터 (가능한 경우)

### 기능 제안
새로운 기능이나 개선 사항 제안을 환영합니다!

---

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.

```
MIT License

Copyright (c) 2025 Excel VBA Macro Project

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👤 작성자

**신진우** - Excel VBA 매크로 개발자

---

## 🔄 업데이트 내역

| 날짜 | 버전 | 내용 |
|------|------|------|
| 2025-10-11 | 2.1 | Merge_sheet.bas 추가 - 엑셀 파일 병합 매크로 |
| 2025-10-10 | 2.0 | bold.bas 추가, README 전체 개편 |
| 2025-10-10 | 1.7 | ItemCode_GroupBy.bas v2.0 개선 |
| 2025-10-10 | 1.0 | 초기 프로젝트 생성 |

---

## ⭐ 스타를 눌러주세요!

이 프로젝트가 도움이 되었다면 GitHub에서 ⭐를 눌러주세요!

---

**최종 수정일**: 2025년 10월 11일
