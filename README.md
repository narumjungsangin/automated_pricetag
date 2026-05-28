# 약국 가격표 생성기

약국에서 약명을 검색해 공공데이터포털 의약품 정보를 자동으로 불러오고, AI로 효능 문구를 다듬어 보기 좋은 가격표를 빠르게 만들어 출력할 수 있는 React 앱입니다.

## 주요 기능

- **약명 검색**: 공공데이터포털 `DrbEasyDrugInfoService`에서 약명·효능을 자동 조회합니다.
- **원본 복원**: 약명을 수정한 뒤 공공데이터 원본 이름으로 되돌릴 수 있습니다.
- **AI 설명 생성**: Gemini(무료) 또는 OpenAI로 긴 효능 설명을 짧고 쉬운 문구로 자동 변환합니다.
- **AI 추가 지시**: AI 대화 로그 아래 입력창에서 "더 짧게 해줘", "분류를 진통제로 바꿔줘" 등 후속 지시를 입력하면 해당 필드만 수정합니다.
- **AI 프롬프트 커스텀**: AI 지시문을 직접 편집하고 저장할 수 있습니다.
- **엑셀 대량 가져오기**: `.xlsx` / `.csv` 파일을 업로드하면 각 행이 가격표 1개로 자동 변환됩니다.
- **색상 테마**: 클래식 네이비, 포레스트 그린, 딥 블루 등 9가지 프리셋 테마를 클릭 한 번으로 적용합니다.
- **폰트 크기 조정**: 약명·증상·분류·하단 각각 슬라이더로 세밀하게 조정합니다.
- **가격표 미리보기**: 수정한 내용이 즉시 가격표 디자인에 반영됩니다.
- **인쇄 출력**: 출력 시 편집 화면은 숨기고 가격표만 인쇄합니다.
- **사용 설명서**: 앱 내 설명서 탭에서 전체화면으로 확인하거나 PDF로 저장할 수 있습니다.

## 실행 방법

1. 의존성을 설치합니다.

```bash
npm install
```

2. 필요한 경우 `.env.example`을 복사해 `.env` 파일을 만들고 공공데이터포털 API 키를 입력합니다.
   앱 화면에서도 API 키를 직접 입력·저장할 수 있으므로 `.env` 설정은 선택사항입니다.

```env
VITE_PUBLIC_DATA_SERVICE_KEY=공공데이터포털_API_키
```

3. 개발 서버를 실행합니다.

```bash
npm run dev
```

## API 키 설정

| 키 | 용도 | 발급처 |
|---|---|---|
| 공공데이터포털 API | 약명·효능 자동 검색 | [data.go.kr](https://www.data.go.kr/data/15075057/openapi.do) — 무료 |
| Gemini API | AI 설명 생성 (권장) | [aistudio.google.com](https://aistudio.google.com/apikey) — 무료 |
| OpenAI API | AI 설명 생성 (선택) | [platform.openai.com](https://platform.openai.com/api-keys) — 유료 |

키는 브라우저 로컬 스토리지에만 저장되며 외부 서버로 전송되지 않습니다.

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

```bash
npm run dist:portable
```

Windows 권한 설정에 따라 심볼릭 링크 생성 오류가 발생할 수 있습니다. 그 경우 관리자 권한 터미널 또는 Windows 개발자 모드를 켠 뒤 다시 실행하세요.

## 기술 스택

- **React 19** + **TypeScript** + **Vite**
- **Electron** (데스크톱 앱 빌드)
- **Google Gemini API** / **OpenAI API**
- **공공데이터포털** 의약품 개요정보(e약은요) API
- **xlsx** (엑셀 파싱)