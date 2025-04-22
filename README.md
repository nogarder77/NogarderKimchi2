# RoyalPalaceKimchi

PDF 파일을 검색하고, 특정 문자열이 포함된 페이지를 필터링하여 다양한 방법으로 출력할 수 있는 콘솔 어플리케이션입니다.

## 기능

- PDF 파일 내 텍스트 검색
- 검색 결과 페이지만 포함하는 새 PDF 파일 생성
- 검색 결과 페이지 번호 목록 출력
- 검색 결과 페이지만 인쇄

## 사용법

```
RoyalPalaceKimchi.exe [PDF파일경로] [검색문자열] [명령]
```

### 명령어

- `-save` : 검색된 페이지만 포함하는 새 PDF 파일 생성
- `-page` : 검색된 페이지 번호 목록을 콘솔에 출력
- `-print` : 검색된 페이지를 기본 프린터로 인쇄
- `-help` : 도움말 메시지 표시
- `/?` : 도움말 메시지 표시
- `-version` : 버전 정보 표시

### 검색 연산자

- `&` : AND 연산자 (예: "단어1 & 단어2")
- `|` : OR 연산자 (예: "단어1 | 단어2")

### 예시

```
RoyalPalaceKimchi.exe "C:\문서\보고서.pdf" "예산 & 2024" -save
RoyalPalaceKimchi.exe "보고서.pdf" "수출 | 수입" -page
```

## 요구사항

- Windows 운영 체제
- .NET Core 8.0 이상 런타임

## 제작 정보

- **제작자**: 김용식
- **연락처**: nogarder@gmail.com
- **저작권**: © 2025 김용식, All Rights Reserved 