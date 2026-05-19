# automated_pricetag
약국에서 약명을 검색해 공공데이터포털 의약품 정보를 불러오고, 약사님이 문구와 가격을 직접 수정한 뒤 가격표로 출력할 수 있는 React 앱입니다.

## 실행 방법

1. 의존성을 설치합니다.

```bash
npm install
```

2. 필요한 경우 `.env.example`을 복사해 `.env` 파일을 만들고 공공데이터포털 API 키를 입력합니다.
   앱 화면에서도 API 키를 직접 입력하고 저장할 수 있으므로 `.env` 설정은 선택사항입니다.

```env
VITE_PUBLIC_DATA_SERVICE_KEY=공공데이터포털_API_키
```

3. 개발 서버를 실행합니다.

```bash
npm run dev
```

## Windows exe 만들기

1. 의존성을 설치합니다.

```bash
npm install
```

2. Electron 실행 파일 폴더를 생성합니다.

```bash
npm run dist
```

3. 생성된 실행 파일을 엽니다.

```text
release\win-unpacked\약국 가격표 생성기.exe
```

배포할 때는 `release\win-unpacked` 폴더 전체를 압축해서 전달하세요. exe 파일만 단독으로 복사하면 실행에 필요한 Electron 파일이 빠져 실행되지 않을 수 있습니다.

### 단일 portable exe가 필요한 경우

아래 명령어를 사용할 수 있습니다.

```bash
npm run dist:portable
```

단, Windows 권한 설정에 따라 심볼릭 링크 생성 오류가 발생할 수 있습니다. 그 경우 관리자 권한 터미널 또는 Windows 개발자 모드를 켠 뒤 다시 실행하세요.

## 주요 기능

- **약명 검색**: 공공데이터포털 `DrbEasyDrugInfoService`에서 약 정보를 조회합니다.
- **가격표 수정**: 약명, 주요 증상, 약 분류, 단위, 가격을 직접 수정할 수 있습니다.
- **가격표 미리보기**: 수정한 내용이 즉시 가격표 디자인에 반영됩니다.
- **인쇄 출력**: 출력 시 편집 화면은 숨기고 가격표만 인쇄합니다.