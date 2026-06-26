import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

type DrugInfo = {
  id: string;
  itemName: string;
  originalItemName?: string;
  symptoms: string;
  category: string;
  unit: string;
  price: string;
};

type EasyDrugApiItem = {
  itemName?: string;
  efcyQesitm?: string;
  useMethodQesitm?: string;
};

type EasyDrugApiResponse = {
  response?: {
    header?: { resultCode?: string; resultMsg?: string };
    body?: { items?: EasyDrugApiItem[] | EasyDrugApiItem };
  };
  header?: { resultCode?: string; resultMsg?: string };
  body?: { items?: EasyDrugApiItem[] | EasyDrugApiItem };
};

type EasyDrugApiPayload = {
  resultCode?: string;
  resultMsg?: string;
  items?: EasyDrugApiItem[] | EasyDrugApiItem;
};

type DrugInfoField = keyof Omit<DrugInfo, 'id'>;

type Toast = { id: string; message: string; type: 'success' | 'error' | 'info' };
type AiLog = { id: string; time: string; provider: string; drugName: string; prompt: string; request: string; response: string; isError: boolean; };

type TagStyle = {
  headerBg: string;
  headerText: string;
  categoryBg: string;
  categoryText: string;
  bodyBg: string;
  bodyText: string;
  borderColor: string;
  fontHeader: number;
  fontSymptoms: number;
  fontCategory: number;
  fontFooter: number;
};

type ExcelDrugRow = { itemName: string; subtitle: string; symptoms: string; extra: string };

type AppTab = 'editor' | 'help';
type AiProvider = 'openai' | 'gemini';

const DEFAULT_TAG_STYLE: TagStyle = {
  headerBg: '#546576',
  headerText: '#ffffff',
  categoryBg: '#6a7b8c',
  categoryText: '#ffffff',
  bodyBg: '#ffffff',
  bodyText: '#222222',
  borderColor: '#546576',
  fontHeader: 26,
  fontSymptoms: 18,
  fontCategory: 13,
  fontFooter: 16,
};

const DEFAULT_DRUG_INFO: DrugInfo = {
  id: crypto.randomUUID(),
  itemName: '',
  symptoms: '',
  category: '',
  unit: '',
  price: '',
};

const SERVICE_KEY_STORAGE_KEY = 'pharmacy-price-tag-service-key';
const OPENAI_KEY_STORAGE_KEY = 'pharmacy-openai-key';
const GEMINI_KEY_STORAGE_KEY = 'pharmacy-gemini-key';
const AI_PROVIDER_STORAGE_KEY = 'pharmacy-ai-provider';
const AI_PROMPT_STORAGE_KEY = 'pharmacy-ai-prompt';
const DEFAULT_AI_PROMPT = `당신은 약국 가격표 문구 전문가입니다. 약국 손님(환자)이 쉽게 이해할 수 있는 짧고 명확한 한국어 증상·효능 문구를 작성하세요. 반드시 50자 이내, 줄바꿈 없이, 핵심 증상/효능만 쉼표로 구분하여 출력하세요. 다른 말은 절대 붙이지 마세요.`;
const FOLLOWUP_AI_PROMPT = `당신은 약국 가격표 편집 전문가입니다. 유저의 지시에 따라 가격표 필드를 수정하세요. 반드시 아래 JSON 형식으로만 응답하세요. 유저가 명시적으로 변경을 요청한 필드만 채우고, 나머지는 반드시 null로 두세요. 다른 말은 절대 붙이지 마세요.
{"itemName": null, "symptoms": null, "category": null, "unit": null, "price": null}`;
const TAG_STYLE_STORAGE_KEY = 'pharmacy-tag-style';
const TAG_SIZE_STORAGE_KEY = 'pharmacy-tag-size';
const TAG_HEIGHT_STORAGE_KEY = 'pharmacy-tag-height';
const EXCEL_DB_PATH_KEY = 'pharmacy-excel-db-path';

type ElectronFileResult = { filePath: string; data: string } | null;
declare global {
  interface Window {
    electronAPI?: {
      openExcelDialog: () => Promise<ElectronFileResult>;
      readExcelFile: (filePath: string) => Promise<ElectronFileResult>;
    };
  }
}

const isElectron = typeof window !== 'undefined' && !!window.electronAPI;

const parseExcelDbFromBase64 = (base64: string): ExcelDrugRow[] => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  const workbook = XLSX.read(bytes, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1 }) as string[][];
  return rows
    .filter((row) => row.some((cell) => cell !== undefined && cell !== ''))
    .map((row) => ({
      itemName: String(row[0] ?? '').trim(),
      subtitle: String(row[1] ?? '').trim(),
      symptoms: String(row[2] ?? '').trim(),
      extra: String(row[3] ?? '').trim(),
    }));
};

const REVIEW_VALUES: Record<DrugInfoField, string[]> = {
  itemName: ['새 가격표', ''],
  originalItemName: [''],
  symptoms: ['주요 증상을 입력하세요', '증상 정보 없음', ''],
  category: ['약 분류', ''],
  unit: ['단위', ''],
  price: ['0', ''],
};

const COLOR_THEMES: { label: string; headerBg: string; headerText: string; categoryBg: string; categoryText: string; bodyBg: string; bodyText: string; borderColor: string }[] = [
  { label: '클래식 네이비',  headerBg: '#546576', headerText: '#ffffff', categoryBg: '#6a7b8c', categoryText: '#ffffff', bodyBg: '#ffffff', bodyText: '#222222', borderColor: '#546576' },
  { label: '포레스트 그린',  headerBg: '#2d6a4f', headerText: '#ffffff', categoryBg: '#52b788', categoryText: '#ffffff', bodyBg: '#f0fdf4', bodyText: '#1a3c2b', borderColor: '#2d6a4f' },
  { label: '딥 블루',        headerBg: '#1e3a5f', headerText: '#ffffff', categoryBg: '#2563eb', categoryText: '#ffffff', bodyBg: '#eff6ff', bodyText: '#1e3a5f', borderColor: '#1e3a5f' },
  { label: '버건디',         headerBg: '#7f1d1d', headerText: '#ffffff', categoryBg: '#b91c1c', categoryText: '#ffffff', bodyBg: '#fff1f2', bodyText: '#3b0a0a', borderColor: '#7f1d1d' },
  { label: '다크 슬레이트',  headerBg: '#1e293b', headerText: '#f8fafc', categoryBg: '#334155', categoryText: '#e2e8f0', bodyBg: '#f8fafc', bodyText: '#1e293b', borderColor: '#1e293b' },
  { label: '라벤더',         headerBg: '#5b21b6', headerText: '#ffffff', categoryBg: '#8b5cf6', categoryText: '#ffffff', bodyBg: '#f5f3ff', bodyText: '#2e1065', borderColor: '#5b21b6' },
  { label: '선셋 오렌지',    headerBg: '#c2410c', headerText: '#ffffff', categoryBg: '#ea580c', categoryText: '#ffffff', bodyBg: '#fff7ed', bodyText: '#431407', borderColor: '#c2410c' },
  { label: '민트 프레시',    headerBg: '#0f766e', headerText: '#ffffff', categoryBg: '#14b8a6', categoryText: '#ffffff', bodyBg: '#f0fdfa', bodyText: '#0f3d38', borderColor: '#0f766e' },
  { label: '화이트 클린',    headerBg: '#e2e8f0', headerText: '#1e293b', categoryBg: '#cbd5e1', categoryText: '#334155', bodyBg: '#ffffff', bodyText: '#1e293b', borderColor: '#94a3b8' },
];

const SIZE_PRESETS = [
  { label: '소 (60mm)', value: 60, heightMm: 39.2, fontHeader: 20, fontSymptoms: 14, fontCategory: 10, fontFooter: 12 },
  { label: '중 (70mm)', value: 70, heightMm: null, fontHeader: 26, fontSymptoms: 18, fontCategory: 13, fontFooter: 16 },
  { label: '대 (85mm)', value: 85, heightMm: null, fontHeader: 32, fontSymptoms: 22, fontCategory: 16, fontFooter: 19 },
  { label: '특대 (100mm)', value: 100, heightMm: null, fontHeader: 38, fontSymptoms: 26, fontCategory: 19, fontFooter: 22 },
];

const stripHtml = (value: string) => value.replace(/<[^>]*>?/gm, '').replace(/\s+/g, ' ').trim();
const getShortItemName = (value: string) => value.split('(')[0].trim();
const needsFieldReview = (field: DrugInfoField, value: string) => REVIEW_VALUES[field].includes(value.trim());
const needsTagReview = (tag: DrugInfo) => (Object.keys(REVIEW_VALUES) as DrugInfoField[]).filter((f) => f !== 'originalItemName').some((f) => needsFieldReview(f, tag[f] ?? ''));
const getInputClassName = (field: DrugInfoField, value: string) => needsFieldReview(field, value) ? 'needs-review' : undefined;
const getDisplayValue = (value: string, fallback: string) => value.trim() || fallback;
const renderPreviewValue = (value: string, fallback: string) => value.trim() || <span className="print-hidden-placeholder">{fallback}</span>;
const renderPreviewPrice = (price: string) => price.trim() ? `${price}원` : <span className="print-hidden-placeholder">가격원</span>;

const getFieldClassName = (baseClassName: string, field: DrugInfoField, value: string) => {
  const classes = [baseClassName];
  if (needsFieldReview(field, value)) classes.push('needs-review');
  if (!value.trim()) classes.push('is-empty');
  return classes.join(' ');
};

const normalizeServiceKey = (serviceKey: string) => {
  const t = serviceKey.trim();
  try { return encodeURIComponent(decodeURIComponent(t)); } catch { return encodeURIComponent(t); }
};

const createDrugSearchUrl = (serviceKey: string, itemName: string) => {
  const params = new URLSearchParams({ itemName, type: 'json', pageNo: '1', numOfRows: '10' });
  return `https://apis.data.go.kr/1471000/DrbEasyDrugInfoService/getDrbEasyDrugList?ServiceKey=${normalizeServiceKey(serviceKey)}&${params.toString()}`;
};

const normalizeItems = (items: EasyDrugApiItem[] | EasyDrugApiItem | undefined): EasyDrugApiItem[] => {
  if (!items) return [];
  return Array.isArray(items) ? items : [items];
};

const getXmlText = (element: Element, tagName: string) =>
  element.getElementsByTagName(tagName)[0]?.textContent?.trim() || undefined;

const parseXmlApiResponse = (responseText: string): EasyDrugApiPayload => {
  const xml = new DOMParser().parseFromString(responseText, 'application/xml');
  if (xml.getElementsByTagName('parsererror')[0]) throw new Error('XML 응답을 해석하지 못했습니다.');
  const items = Array.from(xml.getElementsByTagName('item')).map((item) => ({
    itemName: getXmlText(item, 'itemName'),
    efcyQesitm: getXmlText(item, 'efcyQesitm'),
    useMethodQesitm: getXmlText(item, 'useMethodQesitm'),
  }));
  return { resultCode: getXmlText(xml.documentElement, 'resultCode'), resultMsg: getXmlText(xml.documentElement, 'resultMsg'), items };
};

const parseJsonApiResponse = (responseText: string): EasyDrugApiPayload => {
  const data = JSON.parse(responseText) as EasyDrugApiResponse;
  return {
    resultCode: data.response?.header?.resultCode ?? data.header?.resultCode,
    resultMsg: data.response?.header?.resultMsg ?? data.header?.resultMsg,
    items: data.response?.body?.items ?? data.body?.items,
  };
};

const parseApiResponse = (responseText: string): EasyDrugApiPayload => {
  const t = responseText.trim();
  if (t.startsWith('<')) return parseXmlApiResponse(t);
  try { return parseJsonApiResponse(t); } catch { throw new Error(`API 응답을 해석하지 못했습니다. 응답: ${responseText.slice(0, 120)}`); }
};

const HelpPage = () => {
  const handlePrintHelp = () => {
    document.body.classList.add('print-help-mode');
    window.print();
    setTimeout(() => document.body.classList.remove('print-help-mode'), 500);
  };

  return (
  <div className="help-page">
    <div className="help-hero">
      <p className="eyebrow">사용 설명서</p>
      <h2>약국 가격표 생성기 사용 가이드</h2>
      <p>약 정보를 자동 검색하고, AI로 문구를 다듬어 예쁜 가격표를 빠르게 만들어 출력하세요.</p>
      <button type="button" className="help-pdf-btn no-print" onClick={handlePrintHelp}>
        🖨️ 설명서 PDF로 저장
      </button>
    </div>

    <div className="help-sections">
      <div className="help-section">
        <div className="help-section-num">01</div>
        <div className="help-section-body">
          <h3>🔑 공공데이터 API 키 설정</h3>
          <p>약 정보 자동 검색을 위해 <strong>공공데이터포털</strong>에서 무료 API 키를 발급받아야 합니다.</p>
          <ol>
            <li><a href="https://www.data.go.kr/data/15075057/openapi.do" target="_blank" rel="noreferrer">data.go.kr</a> 접속 → 회원가입 후 "의약품 개요정보(e약은요)" API 신청</li>
            <li>마이페이지에서 발급된 <strong>Encoding 키</strong> 복사</li>
            <li>앱 상단 <strong>"공공데이터 API → 설정"</strong> 클릭 후 붙여넣기 → 저장</li>
          </ol>
          <div className="help-note">💡 키는 브라우저 로컬 스토리지에만 저장되며 외부 서버로 전송되지 않습니다.</div>
        </div>
      </div>

      <div className="help-section">
        <div className="help-section-num">02</div>
        <div className="help-section-body">
          <h3>🔍 약 검색 및 자동 입력</h3>
          <p>약 이름을 검색하면 공공데이터에서 약명과 효능을 자동으로 불러옵니다.</p>
          <ol>
            <li>상단 검색창에 약 이름 입력 (예: "타이레놀", "판콜", "게보린")</li>
            <li><strong>API 검색</strong> 버튼 클릭 또는 Enter</li>
            <li>결과가 여러 개면 목록에서 원하는 약 선택 → 약명·증상 자동 입력</li>
            <li>약명 옆 <strong className="badge-yellow">원본 복원</strong> 버튼으로 공공데이터 원본 약명으로 되돌리기 가능</li>
          </ol>
          <div className="help-note">💡 검색 결과가 없으면 약명을 더 짧게 입력해 보세요 (예: "타이레놀" → "타이레").</div>
        </div>
      </div>

      <div className="help-section">
        <div className="help-section-num">03</div>
        <div className="help-section-body">
          <h3>🤖 AI 설명 생성 (Gemini / OpenAI)</h3>
          <p>AI가 긴 효능 설명을 환자가 이해하기 쉬운 짧은 문구로 바꿔줍니다.</p>
          <div className="help-sub-section">
            <h4>Gemini (무료 권장)</h4>
            <ol>
              <li><a href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer">aistudio.google.com</a>에서 무료 API 키 발급</li>
              <li>앱 상단 <strong>"AI 설명 생성 → 설정"</strong>에서 Gemini 선택 후 키 저장</li>
              <li>편집기에서 약을 선택한 뒤 <strong>✨ AI로 설명 다듬기</strong> 클릭</li>
            </ol>
          </div>
          <div className="help-sub-section">
            <h4>AI 추가 지시 (대화형 편집)</h4>
            <ol>
              <li>AI 로그 아래 입력창에 추가 지시 입력</li>
              <li>예: <em>"더 짧게 해줘"</em>, <em>"분류를 진통제로 바꿔줘"</em>, <em>"영어로 번역해줘"</em></li>
              <li>Enter 또는 ➤ 버튼으로 전송 → 해당 필드만 변경</li>
            </ol>
          </div>
          <div className="help-sub-section">
            <h4>AI 프롬프트 커스텀</h4>
            <p>AI 설정 패널 → <strong>AI 프롬프트 → 편집</strong>에서 AI 지시문을 원하는 대로 수정할 수 있습니다.</p>
          </div>
          <div className="help-note">⚠️ OpenAI는 유료입니다. Gemini 무료 티어는 분당 15회 제한이 있습니다.</div>
        </div>
      </div>

      <div className="help-section">
        <div className="help-section-num">04</div>
        <div className="help-section-body">
          <h3>📊 엑셀 DB 검색</h3>
          <p>기존에 정리해 둔 약품 엑셀 파일을 DB로 불러와 검색할 수 있습니다.</p>
            <div className="help-table-wrap">
              <table className="help-table">
                <thead>
                  <tr><th>A열 (약명)</th><th>B열 (서브타이틀/분류)</th><th>C열 (증상/효능)</th><th>D열 (부가설명)</th></tr>
                </thead>
                <tbody>
                  <tr><td>포박신</td><td>안전한 한방 근육생약</td><td>심한 운동후 / 생리통</td><td>근육 경직 근육통 완화 생약</td></tr>
                  <tr><td>기넥신F 120mg</td><td>식물성캡슐 은행엽120mg</td><td>말초혈관 / 뇌혈관순환제</td><td>두통 어지러움 귀울림</td></tr>
                </tbody>
              </table>
            </div>
            <ol>
              <li>검색창 아래 <strong>📋 엑셀 DB 불러오기</strong> 버튼으로 파일 선택</li>
              <li>검색창에 약 이름 입력 후 검색 → 엑셀 DB + 공공데이터 API 결과가 함께 표시</li>
              <li><span className="search-source-badge excel">엑셀</span> 뱃지가 붙은 결과를 클릭하면 해당 약품 정보가 가격표에 자동 입력</li>
            </ol>
            <div className="help-note">💡 Electron 앱에서는 최초 1회 파일을 선택하면 경로가 저장되어 다음 실행 시 자동으로 불러옵니다. 엑셀 파일 수정 후 🔄 버튼으로 새로고침할 수 있습니다.</div>
        </div>
      </div>

      <div className="help-section">
        <div className="help-section-num">05</div>
        <div className="help-section-body">
          <h3>🎨 디자인 설정</h3>
          <p>모든 가격표에 동일하게 적용되는 디자인을 변경합니다.</p>
          <ul>
            <li><strong>크기 프리셋</strong>: 소(60mm) / 중(70mm) / 대(85mm) / 특대(100mm) — 폰트 크기도 자동 조정</li>
            <li><strong>색상</strong>: 헤더(약명 배경·글자) / 분류 배경·글자 / 본문 배경·글자 / 테두리 색상</li>
            <li><strong>폰트 크기</strong>: 헤더·증상·분류·하단 각각 슬라이더로 세밀 조정</li>
            <li><strong>초기화</strong> 버튼으로 기본값 복원</li>
          </ul>
          <div className="help-note">💡 설정은 자동 저장됩니다. 브라우저를 닫아도 유지됩니다.</div>
        </div>
      </div>

      <div className="help-section">
        <div className="help-section-num">06</div>
        <div className="help-section-body">
          <h3>🖨️ 출력 및 가격표 관리</h3>
          <ul>
            <li><strong>가격표 추가/복사/삭제</strong>: 편집기 상단 버튼으로 관리</li>
            <li><strong>노란 줄무늬</strong>: 수정이 필요한 항목 표시 — 출력 전 확인 권장</li>
            <li><strong>프린터로 출력하기</strong> 버튼 클릭 → 인쇄 대화상자 열림</li>
          </ul>
          <div className="help-note-list">
            <div className="help-note">🖨️ 인쇄 설정 권장: 용지 <strong>A4</strong> · 여백 <strong>최소</strong> · <strong>배경 그래픽 인쇄 체크</strong></div>
            <div className="help-note">💡 PDF로 저장하려면 프린터 목록에서 <strong>"PDF로 저장"</strong> 선택</div>
          </div>
        </div>
      </div>
    </div>

    <div className="help-footer-note">
      <strong>문의사항:</strong> <a href="mailto:joonst26@gmail.com">joonst26@gmail.com</a> ·
      <a href="https://github.com/narumjungsangin" target="_blank" rel="noreferrer"> GitHub @narumjungsangin</a>
    </div>
  </div>
  );
};

function App() {
  const [activeTab, setActiveTab] = useState<AppTab>('editor');
  const [searchQuery, setSearchQuery] = useState('');
  const [serviceKeyInput, setServiceKeyInput] = useState('');
  const [savedServiceKey, setSavedServiceKey] = useState('');
  const [isApiKeyBoxOpen, setIsApiKeyBoxOpen] = useState(false);
  const [hasServiceKeyError, setHasServiceKeyError] = useState(false);
  const [openAiKeyInput, setOpenAiKeyInput] = useState('');
  const [savedOpenAiKey, setSavedOpenAiKey] = useState('');
  const [geminiKeyInput, setGeminiKeyInput] = useState('');
  const [savedGeminiKey, setSavedGeminiKey] = useState('');
  const [aiProvider, setAiProvider] = useState<AiProvider>('gemini');
  const [isAiKeyBoxOpen, setIsAiKeyBoxOpen] = useState(false);
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [aiPrompt, setAiPrompt] = useState(DEFAULT_AI_PROMPT);
  const [isAiPromptOpen, setIsAiPromptOpen] = useState(false);
  const [aiLogs, setAiLogs] = useState<AiLog[]>([]);
  const [isAiLogOpen, setIsAiLogOpen] = useState(false);
  const [aiFollowUp, setAiFollowUp] = useState('');
  const [searchResults, setSearchResults] = useState<EasyDrugApiItem[]>([]);
  const [priceTags, setPriceTags] = useState<DrugInfo[]>([DEFAULT_DRUG_INFO]);
  const [selectedTagId, setSelectedTagId] = useState(DEFAULT_DRUG_INFO.id);
  const [toast, setToast] = useState<Toast | null>(null);
  const [isSearching, setIsSearching] = useState(false);
  const [tagStyle, setTagStyle] = useState<TagStyle>(DEFAULT_TAG_STYLE);
  const [tagSizeMm, setTagSizeMm] = useState(70);
  const [tagHeightMm, setTagHeightMm] = useState<number | null>(null);
  const [customWidthInput, setCustomWidthInput] = useState('');
  const [customHeightInput, setCustomHeightInput] = useState('');
  const [keepAspectRatio, setKeepAspectRatio] = useState(true);
  const [isSizeSliderOpen, setIsSizeSliderOpen] = useState(false);
  const [isDesignOpen, setIsDesignOpen] = useState(false);
  const excelDbInputRef = useRef<HTMLInputElement>(null);
  const [excelDb, setExcelDb] = useState<ExcelDrugRow[]>([]);
  const [excelDbPath, setExcelDbPath] = useState('');
  const [excelSearchResults, setExcelSearchResults] = useState<ExcelDrugRow[]>([]);

  const envServiceKey = import.meta.env.VITE_PUBLIC_DATA_SERVICE_KEY as string | undefined;
  const serviceKey = savedServiceKey || envServiceKey;
  const hasServiceKey = Boolean(serviceKey);
  const needsServiceKeyInput = !hasServiceKey || hasServiceKeyError;
  const selectedTag = priceTags.find((tag) => tag.id === selectedTagId) ?? priceTags[0];
  const canSearch = useMemo(() => searchQuery.trim().length > 0 && !isSearching, [isSearching, searchQuery]);
  const tagWidthPx = Math.round((tagSizeMm / 70) * 280);
  const tagFontScale = tagSizeMm / 70;

  useEffect(() => {
    const storedServiceKey = window.localStorage.getItem(SERVICE_KEY_STORAGE_KEY) ?? '';
    const storedOpenAiKey = window.localStorage.getItem(OPENAI_KEY_STORAGE_KEY) ?? '';
    const storedStyle = window.localStorage.getItem(TAG_STYLE_STORAGE_KEY);
    const storedSize = window.localStorage.getItem(TAG_SIZE_STORAGE_KEY);

    setSavedServiceKey(storedServiceKey);
    setServiceKeyInput(storedServiceKey);
    setIsApiKeyBoxOpen(!storedServiceKey && !envServiceKey);
    const storedGeminiKey = window.localStorage.getItem(GEMINI_KEY_STORAGE_KEY) ?? '';
    const storedProvider = (window.localStorage.getItem(AI_PROVIDER_STORAGE_KEY) ?? 'gemini') as AiProvider;
    setSavedOpenAiKey(storedOpenAiKey);
    setOpenAiKeyInput(storedOpenAiKey);
    setSavedGeminiKey(storedGeminiKey);
    setGeminiKeyInput(storedGeminiKey);
    setAiProvider(storedProvider);
    const storedPrompt = window.localStorage.getItem(AI_PROMPT_STORAGE_KEY);
    if (storedPrompt) setAiPrompt(storedPrompt);

    if (storedStyle) {
      try { setTagStyle(JSON.parse(storedStyle)); } catch { /* ignore */ }
    }
    if (storedSize) {
      const n = Number(storedSize);
      if (!isNaN(n)) setTagSizeMm(n);
    }
    const storedHeight = window.localStorage.getItem(TAG_HEIGHT_STORAGE_KEY);
    if (storedHeight) {
      const h = Number(storedHeight);
      if (!isNaN(h)) setTagHeightMm(h);
    }

    // Electron: 저장된 엑셀 DB 경로가 있으면 자동 로드
    const storedExcelPath = window.localStorage.getItem(EXCEL_DB_PATH_KEY);
    if (storedExcelPath && isElectron) {
      setExcelDbPath(storedExcelPath);
      window.electronAPI!.readExcelFile(storedExcelPath).then((result) => {
        if (!result) return;
        try {
          const dbRows = parseExcelDbFromBase64(result.data);
          if (dbRows.length > 0) setExcelDb(dbRows);
        } catch { /* ignore */ }
      });
    }
  }, [envServiceKey]);

  useEffect(() => {
    if (!toast) return;
    const duration = toast.type === 'error' ? 8000 : toast.type === 'info' ? 3500 : 2600;
    const id = window.setTimeout(() => setToast(null), duration);
    return () => window.clearTimeout(id);
  }, [toast]);

  const showToast = (message: string, type: Toast['type'] = 'info') =>
    setToast({ id: crypto.randomUUID(), message, type });

  const saveServiceKey = () => {
    const k = serviceKeyInput.trim();
    if (!k) { showToast('API 키를 입력해주세요.', 'info'); return; }
    window.localStorage.setItem(SERVICE_KEY_STORAGE_KEY, k);
    setSavedServiceKey(k);
    setHasServiceKeyError(false);
    showToast('API 키를 저장했습니다.', 'success');
  };

  const clearServiceKey = () => {
    window.localStorage.removeItem(SERVICE_KEY_STORAGE_KEY);
    setSavedServiceKey('');
    setServiceKeyInput('');
    setHasServiceKeyError(false);
    showToast('저장된 API 키를 삭제했습니다.', 'info');
  };

  const saveOpenAiKey = () => {
    const k = openAiKeyInput.trim();
    if (!k) { showToast('OpenAI API 키를 입력해주세요.', 'info'); return; }
    window.localStorage.setItem(OPENAI_KEY_STORAGE_KEY, k);
    setSavedOpenAiKey(k);
    showToast('OpenAI API 키를 저장했습니다.', 'success');
  };

  const clearOpenAiKey = () => {
    window.localStorage.removeItem(OPENAI_KEY_STORAGE_KEY);
    setSavedOpenAiKey('');
    setOpenAiKeyInput('');
    showToast('OpenAI API 키를 삭제했습니다.', 'info');
  };

  const saveGeminiKey = () => {
    const k = geminiKeyInput.trim();
    if (!k) { showToast('Gemini API 키를 입력해주세요.', 'info'); return; }
    window.localStorage.setItem(GEMINI_KEY_STORAGE_KEY, k);
    setSavedGeminiKey(k);
    showToast('Gemini API 키를 저장했습니다.', 'success');
  };

  const clearGeminiKey = () => {
    window.localStorage.removeItem(GEMINI_KEY_STORAGE_KEY);
    setSavedGeminiKey('');
    setGeminiKeyInput('');
    showToast('Gemini API 키를 삭제했습니다.', 'info');
  };

  const updateAiProvider = (provider: AiProvider) => {
    setAiProvider(provider);
    window.localStorage.setItem(AI_PROVIDER_STORAGE_KEY, provider);
  };

  const loadExcelDbFromFile = async (result: ElectronFileResult) => {
    if (!result) return;
    try {
      const dbRows = parseExcelDbFromBase64(result.data);
      if (dbRows.length === 0) { showToast('엑셀에 데이터가 없습니다.', 'error'); return; }
      setExcelDb(dbRows);
      setExcelDbPath(result.filePath);
      window.localStorage.setItem(EXCEL_DB_PATH_KEY, result.filePath);
      showToast(`엑셀 DB ${dbRows.length}개 약품 로드 완료`, 'success');
    } catch {
      showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'error');
    }
  };

  const handleExcelDbClick = async () => {
    if (isElectron) {
      const result = await window.electronAPI!.openExcelDialog();
      loadExcelDbFromFile(result);
    } else {
      excelDbInputRef.current?.click();
    }
  };

  const handleExcelDbUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1 }) as string[][];
        const dbRows: ExcelDrugRow[] = rows
          .filter((row) => row.some((cell) => cell !== undefined && cell !== ''))
          .map((row) => ({
            itemName: String(row[0] ?? '').trim(),
            subtitle: String(row[1] ?? '').trim(),
            symptoms: String(row[2] ?? '').trim(),
            extra: String(row[3] ?? '').trim(),
          }));
        if (dbRows.length === 0) { showToast('엑셀에 데이터가 없습니다.', 'error'); return; }
        setExcelDb(dbRows);
        showToast(`엑셀 DB ${dbRows.length}개 약품 로드 완료`, 'success');
      } catch {
        showToast('엑셀 파일을 읽는 중 오류가 발생했습니다.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = '';
  };

  const applyExcelDrug = (row: ExcelDrugRow) => {
    const symptoms = [row.symptoms, row.extra].filter(Boolean).join('\n');
    setPriceTags((prev) => prev.map((tag) => tag.id === selectedTag.id ? {
      ...tag,
      itemName: row.itemName,
      originalItemName: row.itemName,
      symptoms: symptoms || '증상 정보 없음',
      category: row.subtitle,
    } : tag));
    setSearchResults([]);
    setExcelSearchResults([]);
    showToast('엑셀 DB에서 약 정보를 불러왔습니다.', 'success');
  };

  const applyDrugToSelectedTag = (drug: EasyDrugApiItem) => {
    const rawSymptoms = drug.efcyQesitm ? stripHtml(drug.efcyQesitm) : '증상 정보 없음';
    setPriceTags((prev) => prev.map((tag) => tag.id === selectedTag.id ? {
      ...tag,
      itemName: drug.itemName ? getShortItemName(drug.itemName) : tag.itemName,
      originalItemName: drug.itemName ?? undefined,
      symptoms: rawSymptoms,
    } : tag));
    setSearchResults([]);
    showToast('데이터베이스에서 약 정보를 불러왔습니다.', 'success');
  };

  const searchDrug = async (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    const q = searchQuery.trim();
    if (!q) { showToast('약 이름을 입력해주세요.', 'info'); return; }

    // 엑셀 DB 검색 (항상 실행)
    const excelMatches = excelDb.filter((row) => row.itemName.toLowerCase().includes(q.toLowerCase()));
    setExcelSearchResults(excelMatches);

    // API 검색
    if (!serviceKey) {
      if (excelMatches.length > 0) {
        showToast(`엑셀 DB에서 ${excelMatches.length}개 발견 (API 키 미설정)`, 'info');
        return;
      }
      setIsApiKeyBoxOpen(true);
      showToast('API 키를 먼저 입력하고 저장해주세요.', 'error');
      return;
    }
    setIsSearching(true);
    try {
      const response = await fetch(createDrugSearchUrl(serviceKey, q));
      const responseText = await response.text();
      if (!response.ok) {
        setHasServiceKeyError(true); setIsApiKeyBoxOpen(true);
        throw new Error(`API 응답 오류: ${response.status} ${responseText.slice(0, 120)}`);
      }
      const data = parseApiResponse(responseText);
      if (data.resultCode && data.resultCode !== '00') {
        setHasServiceKeyError(true); setIsApiKeyBoxOpen(true);
        throw new Error(`API 오류: ${data.resultMsg ?? data.resultCode}`);
      }
      setHasServiceKeyError(false);
      const drugs = normalizeItems(data.items);
      if (drugs.length === 0 && excelMatches.length === 0) { showToast('검색 결과가 없습니다.', 'info'); return; }
      if (drugs.length === 1 && excelMatches.length === 0) { applyDrugToSelectedTag(drugs[0]); return; }
      setSearchResults(drugs);
      const total = drugs.length + excelMatches.length;
      showToast(`${total}개의 검색 결과가 있습니다. 적용할 약을 선택해주세요.`, 'info');
    } catch (error) {
      console.error('API 호출 에러:', error);
      showToast(error instanceof Error ? error.message : '데이터를 불러오는 중 문제가 발생했습니다.', 'error');
    } finally {
      setIsSearching(false);
    }
  };

  type AiPatch = { itemName?: string | null; symptoms?: string | null; category?: string | null; unit?: string | null; price?: string | null };

  const parseAiResult = (raw: string): AiPatch => {
    try {
      const clean = raw.replace(/```json|```/g, '').trim();
      return JSON.parse(clean) as AiPatch;
    } catch {
      return { symptoms: raw };
    }
  };

  const applyAiPatch = (patch: AiPatch, tag: DrugInfo): DrugInfo => ({
    ...tag,
    ...(patch.itemName != null ? { itemName: patch.itemName } : {}),
    ...(patch.symptoms != null ? { symptoms: patch.symptoms } : {}),
    ...(patch.category != null ? { category: patch.category } : {}),
    ...(patch.unit != null ? { unit: patch.unit } : {}),
    ...(patch.price != null ? { price: patch.price } : {}),
  });

  const callAi = async (systemPrompt: string, userContent: string): Promise<string> => {
    if (aiProvider === 'gemini') {
      if (!savedGeminiKey) { setIsAiKeyBoxOpen(true); throw new Error('Gemini API 키를 먼저 저장해주세요.'); }
      const maxRetries = 3;
      let lastErrMsg = '';
      let geminiRes: Response | null = null;
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        geminiRes = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${savedGeminiKey}`,
          { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ parts: [{ text: `${systemPrompt}\n\n${userContent}` }] }], generationConfig: { maxOutputTokens: 200, temperature: 0.4 } }) }
        );
        if (geminiRes.ok) break;
        if (geminiRes.status === 429 && attempt < maxRetries) {
          showToast(`요청 한도 초과 — ${attempt}/${maxRetries - 1}회 재시도 중...`, 'info');
          await new Promise((r) => setTimeout(r, 5000 * attempt));
          continue;
        }
        try {
          const errJson = await geminiRes.json() as { error?: { message?: string; status?: string } };
          const serverMsg = errJson.error?.message ?? '';
          const serverStatus = errJson.error?.status ?? '';
          if (geminiRes.status === 429) lastErrMsg = `[429] ${serverStatus || 'RESOURCE_EXHAUSTED'}: ${serverMsg || '요청 한도 초과'}`;
          else if (geminiRes.status === 400) lastErrMsg = `[400] ${serverMsg || 'Bad Request'}`;
          else if (geminiRes.status === 403) lastErrMsg = `[403] ${serverMsg || 'Forbidden'} — API 키 권한 확인`;
          else lastErrMsg = `[${geminiRes.status}] ${serverMsg || 'Gemini 오류'}`;
        } catch { lastErrMsg = `Gemini HTTP ${geminiRes.status}`; }
        throw new Error(lastErrMsg);
      }
      if (!geminiRes?.ok) throw new Error(lastErrMsg || 'Gemini 오류');
      const json = await geminiRes.json() as { candidates?: { content?: { parts?: { text?: string }[] } }[] };
      return json.candidates?.[0]?.content?.parts?.[0]?.text?.trim() ?? '';
    } else {
      if (!savedOpenAiKey) { setIsAiKeyBoxOpen(true); throw new Error('OpenAI API 키를 먼저 저장해주세요.'); }
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${savedOpenAiKey}` },
        body: JSON.stringify({ model: 'gpt-4o-mini', messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userContent }], max_tokens: 200, temperature: 0.4 }),
      });
      if (!res.ok) {
        let errMsg = `OpenAI 오류: ${res.status}`;
        try {
          const errJson = await res.json() as { error?: { message?: string; code?: string } };
          const code = errJson.error?.code;
          if (res.status === 429) errMsg = code === 'insufficient_quota' ? 'OpenAI 크레딧이 부족합니다.' : 'Rate Limit 초과. 잠시 후 재시도해주세요.';
          else if (res.status === 401) errMsg = 'OpenAI API 키가 잘못되었습니다.';
          else if (errJson.error?.message) errMsg = `OpenAI: ${errJson.error.message}`;
        } catch { /* ignore */ }
        throw new Error(errMsg);
      }
      const json = await res.json() as { choices: { message: { content: string } }[] };
      return json.choices?.[0]?.message?.content?.trim() ?? '';
    }
  };

  const followUpAi = async () => {
    const instruction = aiFollowUp.trim();
    if (!instruction) { showToast('추가 지시를 입력해주세요.', 'info'); return; }
    const tag = selectedTag;
    const systemPrompt = FOLLOWUP_AI_PROMPT;
    const userContent = `현재 가격표 정보:\n약명: ${tag.itemName || '(없음)'}\n증상: ${tag.symptoms || '(없음)'}\n분류: ${tag.category || '(없음)'}\n단위: ${tag.unit || '(없음)'}\n가격: ${tag.price || '(없음)'}\n\n유저 지시: ${instruction}\n\n유저가 명시적으로 변경을 요청한 필드만 JSON에 값을 넣고, 나머지는 null로 두세요.`;
    setIsAiLoading(true);
    try {
      const raw = await callAi(systemPrompt, userContent);
      if (!raw) throw new Error('AI 응답이 비어 있습니다.');
      const patch = parseAiResult(raw);
      setPriceTags((prev) => prev.map((t) => t.id === tag.id ? applyAiPatch(patch, t) : t));
      setAiLogs((prev) => [{ id: crypto.randomUUID(), time: new Date().toLocaleTimeString('ko-KR'), provider: aiProvider, drugName: tag.itemName || '(미입력)', prompt: systemPrompt, request: userContent, response: raw, isError: false }, ...prev].slice(0, 30));
      setAiFollowUp('');
      showToast('AI가 추가 지시를 반영했습니다.', 'success');
    } catch (error) {
      const msg = error instanceof Error ? error.message : 'AI 요청 중 오류가 발생했습니다.';
      setAiLogs((prev) => [{ id: crypto.randomUUID(), time: new Date().toLocaleTimeString('ko-KR'), provider: aiProvider, drugName: tag.itemName || '(미입력)', prompt: systemPrompt, request: userContent, response: msg, isError: true }, ...prev].slice(0, 30));
      showToast(msg, 'error');
    } finally {
      setIsAiLoading(false);
    }
  };

  const polishWithAi = async () => {
    const tag = selectedTag;
    const hasSymptoms = tag.symptoms.trim() && tag.symptoms.trim() !== '증상 정보 없음';
    if (!hasSymptoms && !tag.itemName.trim()) { showToast('약명을 먼저 입력해주세요.', 'info'); return; }
    const systemPrompt = aiPrompt.trim() || DEFAULT_AI_PROMPT;
    const userContent = hasSymptoms
      ? `다음 약 효능 설명을 약국 가격표용 짧은 문구로 다듬어 주세요.\n\n약명: ${tag.itemName || '(미입력)'}\n효능 원문: ${tag.symptoms}`
      : `다음 약의 주요 증상·효능을 약국 가격표용 짧은 문구로 작성해 주세요.\n\n약명: ${tag.itemName}`;
    setIsAiLoading(true);
    try {
      const raw = await callAi(systemPrompt, userContent);
      if (!raw) throw new Error('AI 응답이 비어 있습니다.');
      const symptomsText = raw.replace(/```json|```/g, '').trim();
      setPriceTags((prev) => prev.map((t) => t.id === tag.id ? { ...t, symptoms: symptomsText } : t));
      setAiLogs((prev) => [{ id: crypto.randomUUID(), time: new Date().toLocaleTimeString('ko-KR'), provider: aiProvider, drugName: tag.itemName || '(미입력)', prompt: systemPrompt, request: userContent, response: raw, isError: false }, ...prev].slice(0, 30));
      showToast(hasSymptoms ? 'AI가 설명을 다듬었습니다.' : 'AI가 설명을 생성했습니다.', 'success');
    } catch (error) {
      const msg = error instanceof Error ? error.message : 'AI 요청 중 오류가 발생했습니다.';
      console.error('[AI Error]', msg);
      setAiLogs((prev) => [{ id: crypto.randomUUID(), time: new Date().toLocaleTimeString('ko-KR'), provider: aiProvider, drugName: tag.itemName || '(미입력)', prompt: systemPrompt, request: userContent, response: msg, isError: true }, ...prev].slice(0, 30));
      showToast(msg, 'error');
    } finally {
      setIsAiLoading(false);
    }
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    setPriceTags((prev) => prev.map((tag) => tag.id === selectedTag.id ? { ...tag, [name]: value } : tag));
    setSearchResults([]);
  };

  const addPriceTag = () => {
    const newTag = { ...DEFAULT_DRUG_INFO, id: crypto.randomUUID() };
    setPriceTags((prev) => [...prev, newTag]);
    setSelectedTagId(newTag.id);
    setSearchQuery('');
    setSearchResults([]);
  };

  const duplicatePriceTag = () => {
    const newTag = { ...selectedTag, id: crypto.randomUUID() };
    setPriceTags((prev) => [...prev, newTag]);
    setSelectedTagId(newTag.id);
    setSearchResults([]);
  };

  const deletePriceTag = () => {
    if (priceTags.length === 1) { showToast('가격표는 최소 1개 이상 필요합니다.', 'info'); return; }
    const nextTags = priceTags.filter((tag) => tag.id !== selectedTag.id);
    setPriceTags(nextTags);
    setSelectedTagId(nextTags[0].id);
    setSearchResults([]);
  };


  const updateTagStyle = (key: keyof TagStyle, value: string | number) => {
    const next = { ...tagStyle, [key]: value };
    setTagStyle(next);
    window.localStorage.setItem(TAG_STYLE_STORAGE_KEY, JSON.stringify(next));
  };

  const resetTagStyle = () => {
    setTagStyle(DEFAULT_TAG_STYLE);
    window.localStorage.setItem(TAG_STYLE_STORAGE_KEY, JSON.stringify(DEFAULT_TAG_STYLE));
    showToast('기본 색상으로 초기화했습니다.', 'info');
  };

  const updateTagSize = (mm: number) => {
    setTagSizeMm(mm);
    setCustomWidthInput('');
    window.localStorage.setItem(TAG_SIZE_STORAGE_KEY, String(mm));
    const preset = SIZE_PRESETS.find((p) => p.value === mm);
    const presetHeight = preset?.heightMm ?? null;
    setTagHeightMm(presetHeight);
    setCustomHeightInput('');
    if (presetHeight !== null) {
      window.localStorage.setItem(TAG_HEIGHT_STORAGE_KEY, String(presetHeight));
    } else {
      window.localStorage.removeItem(TAG_HEIGHT_STORAGE_KEY);
    }
    if (preset) {
      const next = { ...tagStyle, fontHeader: preset.fontHeader, fontSymptoms: preset.fontSymptoms, fontCategory: preset.fontCategory, fontFooter: preset.fontFooter };
      setTagStyle(next);
      window.localStorage.setItem(TAG_STYLE_STORAGE_KEY, JSON.stringify(next));
    }
  };

  const currentRatio = tagHeightMm !== null ? tagHeightMm / tagSizeMm : 0.77;

  const handleCustomWidthChange = (val: string) => {
    setCustomWidthInput(val);
    if (keepAspectRatio && val !== '') {
      const w = parseFloat(val);
      if (!isNaN(w)) {
        setCustomHeightInput((w * currentRatio).toFixed(1));
      }
    }
  };

  const handleCustomHeightChange = (val: string) => {
    setCustomHeightInput(val);
    if (keepAspectRatio && val !== '') {
      const h = parseFloat(val);
      if (!isNaN(h)) {
        setCustomWidthInput((h / currentRatio).toFixed(1));
      }
    }
  };

  const applyCustomSize = () => {
    const w = parseFloat(customWidthInput);
    const h = parseFloat(customHeightInput);
    if (!isNaN(w) && w >= 30 && w <= 200) {
      setTagSizeMm(w);
      window.localStorage.setItem(TAG_SIZE_STORAGE_KEY, String(w));
    }
    if (!isNaN(h) && h >= 20 && h <= 200) {
      setTagHeightMm(h);
      window.localStorage.setItem(TAG_HEIGHT_STORAGE_KEY, String(h));
    } else if (customHeightInput === '') {
      setTagHeightMm(null);
      window.localStorage.removeItem(TAG_HEIGHT_STORAGE_KEY);
    }
  };

  const handlePrint = () => window.print();

  const tagInlineStyle = {
    '--tag-header-bg': tagStyle.headerBg,
    '--tag-header-text': tagStyle.headerText,
    '--tag-category-bg': tagStyle.categoryBg,
    '--tag-category-text': tagStyle.categoryText,
    '--tag-body-bg': tagStyle.bodyBg,
    '--tag-body-text': tagStyle.bodyText,
    '--tag-border': tagStyle.borderColor,
    '--tag-width-px': `${tagWidthPx}px`,
    '--tag-size-mm': `${tagSizeMm}mm`,
    ...(tagHeightMm !== null ? { '--tag-height-mm': `${tagHeightMm}mm` } : {}),
    '--tag-font-scale': `${tagFontScale}`,
    '--tag-font-header': `${tagStyle.fontHeader}px`,
    '--tag-font-symptoms': `${tagStyle.fontSymptoms}px`,
    '--tag-font-category': `${tagStyle.fontCategory}px`,
    '--tag-font-footer': `${tagStyle.fontFooter}px`,
  } as React.CSSProperties;

  return (
    <main className={activeTab === 'help' ? 'app-container help-mode' : 'app-container'} style={tagInlineStyle}>
      {toast && (
        <div className={`toast toast-${toast.type}`} key={toast.id} role="status">
          {toast.message}
        </div>
      )}

      <section className="control-panel no-print" aria-label="가격표 편집 영역">
        <div className="panel-heading">
          <p className="eyebrow">Pharmacy Label Maker</p>
          <h1>약국 가격표 생성기</h1>
          <p>약명을 검색한 뒤, 여러 가격표를 모아 한 번에 출력하세요.</p>
          <div className="tab-nav">
            <button type="button" className={activeTab === 'editor' ? 'tab-btn active' : 'tab-btn'} onClick={() => setActiveTab('editor')}>편집기</button>
            <button type="button" className={activeTab === 'help' ? 'tab-btn active' : 'tab-btn'} onClick={() => setActiveTab('help')}>사용 설명서</button>
          </div>
        </div>

        {activeTab === 'help' ? <HelpPage /> : (
          <>
            {/* ① 가격표 목록 */}
            <div className="tag-list">
              <div className="tag-list-header">
                <strong>가격표 목록</strong>
                <span>{priceTags.length}개</span>
              </div>
              <div className="tag-list-items">
                {priceTags.map((tag, index) => (
                  <button
                    className={`${tag.id === selectedTag.id ? 'tag-list-item active' : 'tag-list-item'} ${needsTagReview(tag) ? 'needs-review' : ''}`}
                    key={tag.id} type="button" onClick={() => setSelectedTagId(tag.id)}>
                    <span>{index + 1}. {getDisplayValue(tag.itemName, '이름 미입력')}</span>
                    <small>{needsTagReview(tag) ? '수정 필요' : `${tag.price}원`}</small>
                  </button>
                ))}
              </div>
              <div className="tag-actions">
                <button type="button" onClick={addPriceTag}>+ 새 가격표</button>
                <button type="button" onClick={duplicatePriceTag}>복제</button>
                <button type="button" onClick={deletePriceTag}>삭제</button>
              </div>
            </div>

            {/* ② 검색 */}
            <div className="search-row">
              <form className="search-box" onSubmit={searchDrug}>
                <input type="text" placeholder="약 이름 검색 (예: 타이레놀, 판콜)" value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)} />
                <button type="submit" disabled={!canSearch}>{isSearching ? '검색 중…' : '🔍 검색'}</button>
              </form>
              <div className="excel-db-row">
                {!isElectron && <input ref={excelDbInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={handleExcelDbUpload} />}
                <button type="button" className="excel-db-btn" onClick={handleExcelDbClick}>
                  📋 {excelDb.length > 0 ? `엑셀 DB (${excelDb.length}개)` : '엑셀 DB 불러오기'}
                </button>
                {isElectron && excelDbPath && excelDb.length > 0 && (
                  <button type="button" className="excel-db-refresh-btn" title="엑셀 파일 새로고침" onClick={async () => {
                    const result = await window.electronAPI!.readExcelFile(excelDbPath);
                    loadExcelDbFromFile(result);
                  }}>🔄</button>
                )}
              </div>
              {excelDbPath && <span className="excel-db-path" title={excelDbPath}>{excelDbPath.split(/[\\/]/).pop()}</span>}
            </div>

            {(excelSearchResults.length > 0 || searchResults.length > 0) && (
              <div className="search-results">
                <div className="search-results-header">
                  <strong>검색 결과</strong>
                  <span>{excelSearchResults.length + searchResults.length}개</span>
                </div>
                {excelSearchResults.map((row, index) => (
                  <button className="search-result-item excel" key={`excel-${index}-${row.itemName}`} type="button" onClick={() => applyExcelDrug(row)}>
                    <strong><span className="search-source-badge excel">엑셀</span> {row.itemName}</strong>
                    <span>{[row.subtitle, row.symptoms, row.extra].filter(Boolean).join(' · ')}</span>
                  </button>
                ))}
                {searchResults.map((drug, index) => (
                  <button className="search-result-item" key={`api-${index}-${drug.itemName}`} type="button" onClick={() => applyDrugToSelectedTag(drug)}>
                    <strong><span className="search-source-badge api">API</span> {getDisplayValue(drug.itemName ?? '', '제품명 없음')}</strong>
                    <span>{getDisplayValue(stripHtml(drug.efcyQesitm ?? ''), '효능 정보 없음')}</span>
                  </button>
                ))}
              </div>
            )}

            {/* ③ 편집 폼 */}
            <div className="edit-form">
              <div className="label-with-action">
                <label htmlFor="itemName">약명</label>
                {selectedTag.originalItemName && selectedTag.originalItemName !== selectedTag.itemName && (
                  <button type="button" className="restore-btn" onClick={() => setPriceTags((prev) => prev.map((t) => t.id === selectedTag.id ? { ...t, itemName: getShortItemName(selectedTag.originalItemName ?? '') } : t))}>
                    원본 복원
                  </button>
                )}
              </div>
              <textarea id="itemName" className={getInputClassName('itemName', selectedTag.itemName)} name="itemName" value={selectedTag.itemName} onChange={handleInputChange} rows={2} />

              <label htmlFor="symptoms">주요 증상</label>
              <textarea id="symptoms" className={getInputClassName('symptoms', selectedTag.symptoms)} name="symptoms" value={selectedTag.symptoms} onChange={handleInputChange} rows={3} />
              <button type="button" className={`ai-polish-btn${isAiLoading ? ' loading' : ''}`} onClick={polishWithAi} disabled={isAiLoading}>
                {isAiLoading ? '⏳ AI 처리 중...' : '✨ AI로 설명 다듬기'}
              </button>
              {aiLogs.length > 0 && (
                <div className="ai-log-inline">
                  <div className="ai-log-inline-header">
                    <span className="ai-log-inline-title">AI 대화 로그 <span className="ai-log-count">{aiLogs.length}</span></span>
                    <div style={{ display: 'flex', gap: 4 }}>
                      <button type="button" className="ai-prompt-reset-btn" style={{ fontSize: 11, padding: '3px 7px' }} onClick={() => setAiLogs([])}>삭제</button>
                      <button type="button" className="ai-prompt-toggle-btn" style={{ fontSize: 11, padding: '3px 7px' }} onClick={() => setIsAiLogOpen((p) => !p)}>{isAiLogOpen ? '접기' : '보기'}</button>
                    </div>
                  </div>
                  {isAiLogOpen && (
                    <div className="ai-log-list">
                      {aiLogs.map((log) => (
                        <div key={log.id} className={log.isError ? 'ai-log-item error' : 'ai-log-item'}>
                          <div className="ai-log-meta">
                            <span className="ai-log-time">{log.time}</span>
                            <span className="ai-log-provider">{log.provider === 'gemini' ? 'Gemini' : 'OpenAI'}</span>
                            <span className="ai-log-drug">{log.drugName}</span>
                            {log.isError && <span className="ai-log-badge error">오류</span>}
                          </div>
                          <div className="ai-log-section">
                            <span className="ai-log-label">요청</span>
                            <p className="ai-log-text">{log.request}</p>
                          </div>
                          <div className="ai-log-section">
                            <span className="ai-log-label">{log.isError ? '오류' : '응답'}</span>
                            <p className="ai-log-text response">{log.response}</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                  <div className="ai-followup-row">
                    <input
                      type="text"
                      className="ai-followup-input"
                      placeholder='"더 짧게 해줘", "지사제 추가해줘" 등 추가 지시...'
                      value={aiFollowUp}
                      onChange={(e) => setAiFollowUp(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); followUpAi(); } }}
                      disabled={isAiLoading}
                    />
                    <button type="button" className="ai-followup-btn" onClick={followUpAi} disabled={isAiLoading}>
                      {isAiLoading ? '⏳' : '➤'}
                    </button>
                  </div>
                </div>
              )}

              <label htmlFor="category">약 분류</label>
              <textarea id="category" className={getInputClassName('category', selectedTag.category)} name="category" value={selectedTag.category} onChange={handleInputChange} rows={2} />

              <div className="half-inputs">
                <div>
                  <label htmlFor="unit">단위</label>
                  <input id="unit" className={getInputClassName('unit', selectedTag.unit)} name="unit" value={selectedTag.unit} onChange={handleInputChange} />
                </div>
                <div>
                  <label htmlFor="price">가격</label>
                  <input id="price" className={getInputClassName('price', selectedTag.price)} name="price" value={selectedTag.price} onChange={handleInputChange} />
                </div>
              </div>
            </div>

            <button className="print-btn" type="button" onClick={handlePrint}>�️ 프린터로 출력하기</button>

            {/* ④ 디자인 설정 */}
            <div className={isDesignOpen ? 'design-box is-open' : 'design-box'}>
              <div className="design-box-header">
                <strong>🎨 디자인 설정</strong>
                <div className="design-header-right">
                  <button type="button" className="design-reset-btn" onClick={resetTagStyle}>초기화</button>
                  <button type="button" onClick={() => setIsDesignOpen((p) => !p)}>{isDesignOpen ? '접기' : '열기'}</button>
                </div>
              </div>
              {isDesignOpen && (
                <div className="design-box-content">
                  <div className="design-size-row">
                    <span className="design-label">가격표 크기</span>
                    <div className="size-preset-btns">
                      {SIZE_PRESETS.map((p) => (
                        <button key={p.value} type="button"
                          className={tagSizeMm === p.value && customWidthInput === '' && customHeightInput === '' ? 'size-preset-btn active' : 'size-preset-btn'}
                          onClick={() => updateTagSize(p.value)}>{p.label}</button>
                      ))}
                    </div>
                    <div className="design-custom-size-row">
                      <span className="design-label">직접 입력</span>
                      <div className="design-custom-size-inputs">
                        <label className="design-custom-size-label">
                          가로
                          <input type="number" className="design-custom-size-input" min={30} max={200} step={0.5}
                            placeholder={String(tagSizeMm)}
                            value={customWidthInput}
                            onChange={(e) => handleCustomWidthChange(e.target.value)} />
                          <small>mm</small>
                        </label>
                        <label className="design-custom-size-label">
                          세로
                          <input type="number" className="design-custom-size-input" min={20} max={200} step={0.5}
                            placeholder={tagHeightMm !== null ? String(tagHeightMm) : `${(tagSizeMm * 0.77).toFixed(1)}`}
                            value={customHeightInput}
                            onChange={(e) => handleCustomHeightChange(e.target.value)} />
                          <small>mm</small>
                        </label>
                        <label className="design-custom-size-label design-aspect-label">
                          <input type="checkbox" checked={keepAspectRatio} onChange={(e) => setKeepAspectRatio(e.target.checked)} />
                          비율 유지
                        </label>
                        <button type="button" className="design-custom-size-btn" onClick={applyCustomSize}>적용</button>
                        <button type="button" className={`design-custom-size-toggle${isSizeSliderOpen ? ' active' : ''}`} onClick={() => setIsSizeSliderOpen((p) => !p)} title="슬라이더">↕</button>
                      </div>
                    </div>
                    {isSizeSliderOpen && (
                      <div className="design-size-slider-box">
                        <div className="design-font-row">
                          <span className="design-label">가로</span>
                          <input type="range" min={30} max={200} step={0.5}
                            value={customWidthInput !== '' ? parseFloat(customWidthInput) : tagSizeMm}
                            onChange={(e) => { handleCustomWidthChange(e.target.value); }}
                          />
                          <span className="design-font-val">{customWidthInput !== '' ? parseFloat(customWidthInput).toFixed(1) : tagSizeMm}<small>mm</small></span>
                        </div>
                        <div className="design-font-row">
                          <span className="design-label">세로</span>
                          <input type="range" min={20} max={200} step={0.5}
                            value={customHeightInput !== '' ? parseFloat(customHeightInput) : (tagHeightMm !== null ? tagHeightMm : parseFloat((tagSizeMm * 0.77).toFixed(1)))}
                            onChange={(e) => { handleCustomHeightChange(e.target.value); }}
                          />
                          <span className="design-font-val">{customHeightInput !== '' ? parseFloat(customHeightInput).toFixed(1) : (tagHeightMm !== null ? tagHeightMm : (tagSizeMm * 0.77).toFixed(1))}<small>mm</small></span>
                        </div>
                        <button type="button" className="design-custom-size-btn" style={{ alignSelf: 'flex-end' }} onClick={applyCustomSize}>적용</button>
                      </div>
                    )}
                  </div>
                  <div className="design-font-grid">
                    <div className="design-font-row">
                      <span className="design-label">약명</span>
                      <input type="range" min={14} max={48} value={tagStyle.fontHeader} onChange={(e) => updateTagStyle('fontHeader', Number(e.target.value))} />
                      <span className="design-font-val">{tagStyle.fontHeader}<small>px</small></span>
                    </div>
                    <div className="design-font-row">
                      <span className="design-label">증상</span>
                      <input type="range" min={10} max={36} value={tagStyle.fontSymptoms} onChange={(e) => updateTagStyle('fontSymptoms', Number(e.target.value))} />
                      <span className="design-font-val">{tagStyle.fontSymptoms}<small>px</small></span>
                    </div>
                    <div className="design-font-row">
                      <span className="design-label">분류</span>
                      <input type="range" min={8} max={28} value={tagStyle.fontCategory} onChange={(e) => updateTagStyle('fontCategory', Number(e.target.value))} />
                      <span className="design-font-val">{tagStyle.fontCategory}<small>px</small></span>
                    </div>
                    <div className="design-font-row">
                      <span className="design-label">단위/가격</span>
                      <input type="range" min={8} max={32} value={tagStyle.fontFooter} onChange={(e) => updateTagStyle('fontFooter', Number(e.target.value))} />
                      <span className="design-font-val">{tagStyle.fontFooter}<small>px</small></span>
                    </div>
                  </div>
                  <div className="design-theme-section">
                    <span className="design-label">색상 테마</span>
                    <div className="design-theme-grid">
                      {COLOR_THEMES.map((theme) => {
                        const isActive = tagStyle.headerBg === theme.headerBg && tagStyle.categoryBg === theme.categoryBg;
                        return (
                          <button
                            key={theme.label}
                            type="button"
                            className={isActive ? 'theme-swatch active' : 'theme-swatch'}
                            title={theme.label}
                            onClick={() => {
                              const next = { ...tagStyle, ...theme };
                              setTagStyle(next);
                              window.localStorage.setItem(TAG_STYLE_STORAGE_KEY, JSON.stringify(next));
                            }}
                          >
                            <span className="swatch-header" style={{ background: theme.headerBg }} />
                            <span className="swatch-category" style={{ background: theme.categoryBg }} />
                            <span className="swatch-body" style={{ background: theme.bodyBg }} />
                            <span className="swatch-label">{theme.label}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  <div className="design-advanced">
                    <span className="design-label">세부 색상 직접 지정</span>
                    <div className="design-color-grid">
                      <div className="design-color-row">
                        <span className="design-label">헤더 배경</span>
                        <input type="color" value={tagStyle.headerBg} onChange={(e) => updateTagStyle('headerBg', e.target.value)} />
                        <span className="design-hex">{tagStyle.headerBg}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">헤더 글자</span>
                        <input type="color" value={tagStyle.headerText} onChange={(e) => updateTagStyle('headerText', e.target.value)} />
                        <span className="design-hex">{tagStyle.headerText}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">분류 배경</span>
                        <input type="color" value={tagStyle.categoryBg} onChange={(e) => updateTagStyle('categoryBg', e.target.value)} />
                        <span className="design-hex">{tagStyle.categoryBg}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">분류 글자</span>
                        <input type="color" value={tagStyle.categoryText} onChange={(e) => updateTagStyle('categoryText', e.target.value)} />
                        <span className="design-hex">{tagStyle.categoryText}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">본문 배경</span>
                        <input type="color" value={tagStyle.bodyBg} onChange={(e) => updateTagStyle('bodyBg', e.target.value)} />
                        <span className="design-hex">{tagStyle.bodyBg}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">본문 글자</span>
                        <input type="color" value={tagStyle.bodyText} onChange={(e) => updateTagStyle('bodyText', e.target.value)} />
                        <span className="design-hex">{tagStyle.bodyText}</span>
                      </div>
                      <div className="design-color-row">
                        <span className="design-label">테두리</span>
                        <input type="color" value={tagStyle.borderColor} onChange={(e) => updateTagStyle('borderColor', e.target.value)} />
                        <span className="design-hex">{tagStyle.borderColor}</span>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* ⑤ API 설정 */}
            <div className="api-key-row">
            <div className={`api-key-box${isApiKeyBoxOpen ? ' is-open' : ''}${needsServiceKeyInput ? ' unset' : ''}`}>
              <div className="api-key-header">
                <div>
                  <strong>공공데이터 API</strong>
                  <span className={needsServiceKeyInput ? 'api-key-status error' : 'api-key-status success'}>
                    {needsServiceKeyInput ? '미설정' : '사용 가능'}
                  </span>
                </div>
                <button type="button" onClick={() => setIsApiKeyBoxOpen((prev) => !prev)}>
                  {isApiKeyBoxOpen ? '접기' : '설정'}
                </button>
              </div>
              {isApiKeyBoxOpen && (
                <div className="api-key-content">
                  <a className="api-key-link" href="https://www.data.go.kr/data/15075057/openapi.do" target="_blank" rel="noreferrer">API 키 발급 받기</a>
                  <input type="password" placeholder="API 키를 입력한 뒤 저장 (Encoding)" value={serviceKeyInput} onChange={(e) => setServiceKeyInput(e.target.value)} />
                  <div className="api-key-actions">
                    <button type="button" onClick={saveServiceKey}>저장</button>
                    <button type="button" onClick={clearServiceKey}>삭제</button>
                  </div>
                </div>
              )}
            </div>
            <div className={`api-key-box${isAiKeyBoxOpen ? ' is-open' : ''}${!(aiProvider === 'gemini' ? savedGeminiKey : savedOpenAiKey) ? ' unset' : ''}`} style={{ flex: 1 }}>
              <div className="api-key-header">
                <div>
                  <strong>AI 설명 생성</strong>
                  <span className={(aiProvider === 'gemini' ? savedGeminiKey : savedOpenAiKey) ? 'api-key-status success' : 'api-key-status error'}>
                    {(aiProvider === 'gemini' ? savedGeminiKey : savedOpenAiKey) ? `${aiProvider === 'gemini' ? 'Gemini' : 'OpenAI'} 사용 가능` : '미설정'}
                  </span>
                </div>
                <button type="button" onClick={() => setIsAiKeyBoxOpen((prev) => !prev)}>
                  {isAiKeyBoxOpen ? '접기' : 'AI 설정'}
                </button>
              </div>
              {isAiKeyBoxOpen && (
                <div className="api-key-content">
                  <div className="ai-provider-toggle">
                    <button type="button" className={aiProvider === 'gemini' ? 'ai-provider-btn active gemini' : 'ai-provider-btn'} onClick={() => updateAiProvider('gemini')}>Gemini <small>무료</small></button>
                    <button type="button" className={aiProvider === 'openai' ? 'ai-provider-btn active openai' : 'ai-provider-btn'} onClick={() => updateAiProvider('openai')}>OpenAI <small>유료</small></button>
                  </div>
                  {aiProvider === 'gemini' ? (
                    <>
                      <a className="api-key-link gemini-link" href="https://aistudio.google.com/apikey" target="_blank" rel="noreferrer">Gemini API 키 발급 (무료)</a>
                      <input type="password" placeholder="AIza..." value={geminiKeyInput} onChange={(e) => setGeminiKeyInput(e.target.value)} />
                      <div className="api-key-actions">
                        <button type="button" onClick={saveGeminiKey}>저장</button>
                        <button type="button" onClick={clearGeminiKey}>삭제</button>
                      </div>
                    </>
                  ) : (
                    <>
                      <a className="api-key-link openai-link" href="https://platform.openai.com/api-keys" target="_blank" rel="noreferrer">OpenAI API 키 발급</a>
                      <input type="password" placeholder="sk-..." value={openAiKeyInput} onChange={(e) => setOpenAiKeyInput(e.target.value)} />
                      <div className="api-key-actions">
                        <button type="button" onClick={saveOpenAiKey}>저장</button>
                        <button type="button" onClick={clearOpenAiKey}>삭제</button>
                      </div>
                    </>
                  )}
                  <div className="ai-prompt-section">
                    <div className="ai-prompt-header">
                      <span className="design-label">AI 프롬프트</span>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <button type="button" className="ai-prompt-reset-btn" onClick={() => { setAiPrompt(DEFAULT_AI_PROMPT); window.localStorage.removeItem(AI_PROMPT_STORAGE_KEY); showToast('프롬프트를 초기화했습니다.', 'info'); }}>초기화</button>
                        <button type="button" className="ai-prompt-toggle-btn" onClick={() => setIsAiPromptOpen((p) => !p)}>{isAiPromptOpen ? '접기' : '편집'}</button>
                      </div>
                    </div>
                    {isAiPromptOpen && (
                      <>
                        <textarea className="ai-prompt-textarea" value={aiPrompt} rows={5} onChange={(e) => setAiPrompt(e.target.value)} placeholder={DEFAULT_AI_PROMPT} />
                        <button type="button" className="ai-prompt-save-btn" onClick={() => { window.localStorage.setItem(AI_PROMPT_STORAGE_KEY, aiPrompt); showToast('프롬프트를 저장했습니다.', 'success'); }}>프롬프트 저장</button>
                      </>
                    )}
                  </div>
                </div>
              )}
            </div>
            </div>

            <footer className="app-footer">
              <span>Made by Junsu Yoon</span>
              <a href="https://github.com/narumjungsangin" target="_blank" rel="noreferrer">GitHub @narumjungsangin</a>
              <span>문의사항: <a href="mailto:joonst26@gmail.com">joonst26@gmail.com</a></span>
            </footer>
          </>
        )}
      </section>

      <section className="preview-panel" aria-label="가격표 미리보기">
        <div className="price-tag-grid">
          {priceTags.map((tag) => (
            <div className={needsTagReview(tag) ? 'price-tag has-review' : 'price-tag'} key={tag.id}>
              <div className={getFieldClassName('tag-header', 'itemName', tag.itemName)}>{renderPreviewValue(tag.itemName, '약명 입력')}</div>
              <div className={getFieldClassName('tag-symptoms', 'symptoms', tag.symptoms)}>{renderPreviewValue(tag.symptoms, '주요 증상 입력')}</div>
              <div className={getFieldClassName('tag-category', 'category', tag.category)}>{renderPreviewValue(tag.category, '약 분류 입력')}</div>
              <div className={!tag.unit.trim() && !tag.price.trim() ? 'tag-footer is-empty' : 'tag-footer'}>
                <div className={getFieldClassName('tag-unit', 'unit', tag.unit)}>{renderPreviewValue(tag.unit, '단위')}</div>
                <div className={getFieldClassName('tag-price', 'price', tag.price)}>{renderPreviewPrice(tag.price)}</div>
              </div>
            </div>
          ))}
        </div>
      </section>
    </main>
  );
}

export default App;
