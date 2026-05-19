import { ChangeEvent, FormEvent, useEffect, useMemo, useState } from 'react';

type DrugInfo = {
  id: string;
  itemName: string;
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
    header?: {
      resultCode?: string;
      resultMsg?: string;
    };
    body?: {
      items?: EasyDrugApiItem[] | EasyDrugApiItem;
    };
  };
  header?: {
    resultCode?: string;
    resultMsg?: string;
  };
  body?: {
    items?: EasyDrugApiItem[] | EasyDrugApiItem;
  };
};

type EasyDrugApiPayload = {
  resultCode?: string;
  resultMsg?: string;
  items?: EasyDrugApiItem[] | EasyDrugApiItem;
};

type DrugInfoField = keyof Omit<DrugInfo, 'id'>;

type Toast = {
  id: string;
  message: string;
  type: 'success' | 'error' | 'info';
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

const REVIEW_VALUES: Record<DrugInfoField, string[]> = {
  itemName: ['새 가격표', ''],
  symptoms: ['주요 증상을 입력하세요', '증상 정보 없음', ''],
  category: ['약 분류', ''],
  unit: ['단위', ''],
  price: ['0', ''],
};

const stripHtml = (value: string) => value.replace(/<[^>]*>?/gm, '').replace(/\s+/g, ' ').trim();

const getShortItemName = (value: string) => value.split('(')[0].trim();

const needsFieldReview = (field: DrugInfoField, value: string) => {
  return REVIEW_VALUES[field].includes(value.trim());
};

const needsTagReview = (tag: DrugInfo) => {
  return (Object.keys(REVIEW_VALUES) as DrugInfoField[]).some((field) => needsFieldReview(field, tag[field]));
};

const getInputClassName = (field: DrugInfoField, value: string) => {
  return needsFieldReview(field, value) ? 'needs-review' : undefined;
};

const getDisplayValue = (value: string, fallback: string) => {
  return value.trim() || fallback;
};

const renderPreviewValue = (value: string, fallback: string) => {
  return value.trim() || <span className="print-hidden-placeholder">{fallback}</span>;
};

const renderPreviewPrice = (price: string) => {
  return price.trim() ? `${price}원` : <span className="print-hidden-placeholder">가격원</span>;
};

const getFieldClassName = (baseClassName: string, field: DrugInfoField, value: string) => {
  const classes = [baseClassName];

  if (needsFieldReview(field, value)) {
    classes.push('needs-review');
  }

  if (!value.trim()) {
    classes.push('is-empty');
  }

  return classes.join(' ');
};

const normalizeServiceKey = (serviceKey: string) => {
  const trimmedServiceKey = serviceKey.trim();
  return /%[0-9A-F]{2}/i.test(trimmedServiceKey) ? trimmedServiceKey : encodeURIComponent(trimmedServiceKey);
};

const createDrugSearchUrl = (serviceKey: string, itemName: string) => {
  const params = new URLSearchParams({
    itemName,
    type: 'json',
    pageNo: '1',
    numOfRows: '10',
  });

  return `https://apis.data.go.kr/1471000/DrbEasyDrugInfoService/getDrbEasyDrugList?ServiceKey=${normalizeServiceKey(serviceKey)}&${params.toString()}`;
};

const normalizeItems = (items: EasyDrugApiItem[] | EasyDrugApiItem | undefined): EasyDrugApiItem[] => {
  if (!items) {
    return [];
  }

  return Array.isArray(items) ? items : [items];
};

const getXmlText = (element: Element, tagName: string) => {
  return element.getElementsByTagName(tagName)[0]?.textContent?.trim() || undefined;
};

const parseXmlApiResponse = (responseText: string): EasyDrugApiPayload => {
  const xml = new DOMParser().parseFromString(responseText, 'application/xml');
  const parserError = xml.getElementsByTagName('parsererror')[0];

  if (parserError) {
    throw new Error('XML 응답을 해석하지 못했습니다.');
  }

  const items = Array.from(xml.getElementsByTagName('item')).map((item) => ({
    itemName: getXmlText(item, 'itemName'),
    efcyQesitm: getXmlText(item, 'efcyQesitm'),
    useMethodQesitm: getXmlText(item, 'useMethodQesitm'),
  }));

  return {
    resultCode: getXmlText(xml.documentElement, 'resultCode'),
    resultMsg: getXmlText(xml.documentElement, 'resultMsg'),
    items,
  };
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
  const trimmedResponseText = responseText.trim();

  if (trimmedResponseText.startsWith('<')) {
    return parseXmlApiResponse(trimmedResponseText);
  }

  try {
    return parseJsonApiResponse(trimmedResponseText);
  } catch {
    throw new Error(`API 응답을 해석하지 못했습니다. 응답: ${responseText.slice(0, 120)}`);
  }
};

function App() {
  const [searchQuery, setSearchQuery] = useState('');
  const [serviceKeyInput, setServiceKeyInput] = useState('');
  const [savedServiceKey, setSavedServiceKey] = useState('');
  const [searchResults, setSearchResults] = useState<EasyDrugApiItem[]>([]);
  const [priceTags, setPriceTags] = useState<DrugInfo[]>([DEFAULT_DRUG_INFO]);
  const [selectedTagId, setSelectedTagId] = useState(DEFAULT_DRUG_INFO.id);
  const [toast, setToast] = useState<Toast | null>(null);
  const [isSearching, setIsSearching] = useState(false);

  const envServiceKey = import.meta.env.VITE_PUBLIC_DATA_SERVICE_KEY as string | undefined;
  const serviceKey = savedServiceKey || envServiceKey;
  const selectedTag = priceTags.find((tag) => tag.id === selectedTagId) ?? priceTags[0];

  const canSearch = useMemo(() => searchQuery.trim().length > 0 && !isSearching, [isSearching, searchQuery]);

  useEffect(() => {
    const storedServiceKey = window.localStorage.getItem(SERVICE_KEY_STORAGE_KEY) ?? '';
    setSavedServiceKey(storedServiceKey);
    setServiceKeyInput(storedServiceKey);
  }, []);

  useEffect(() => {
    if (!toast) {
      return;
    }

    const timeoutId = window.setTimeout(() => setToast(null), 2600);
    return () => window.clearTimeout(timeoutId);
  }, [toast]);

  const showToast = (message: string, type: Toast['type'] = 'info') => {
    setToast({
      id: crypto.randomUUID(),
      message,
      type,
    });
  };

  const saveServiceKey = () => {
    const nextServiceKey = serviceKeyInput.trim();

    if (!nextServiceKey) {
      showToast('API 키를 입력해주세요.', 'info');
      return;
    }

    window.localStorage.setItem(SERVICE_KEY_STORAGE_KEY, nextServiceKey);
    setSavedServiceKey(nextServiceKey);
    showToast('API 키를 저장했습니다.', 'success');
  };

  const clearServiceKey = () => {
    window.localStorage.removeItem(SERVICE_KEY_STORAGE_KEY);
    setSavedServiceKey('');
    setServiceKeyInput('');
    showToast('저장된 API 키를 삭제했습니다.', 'info');
  };

  const applyDrugToSelectedTag = (drug: EasyDrugApiItem) => {
    const rawSymptoms = drug.efcyQesitm ? stripHtml(drug.efcyQesitm) : '증상 정보 없음';

    setPriceTags((prev) => prev.map((tag) => tag.id === selectedTag.id ? {
      ...tag,
      itemName: drug.itemName ? getShortItemName(drug.itemName) : tag.itemName,
      symptoms: rawSymptoms,
    } : tag));
    setSearchResults([]);
    showToast('데이터베이스에서 약 정보를 불러왔습니다.', 'success');
  };

  const searchDrug = async (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();

    if (!searchQuery.trim()) {
      showToast('약 이름을 입력해주세요.', 'info');
      return;
    }

    if (!serviceKey) {
      showToast('API 키를 먼저 입력하고 저장해주세요.', 'error');
      return;
    }

    setIsSearching(true);

    try {
      const response = await fetch(createDrugSearchUrl(serviceKey, searchQuery.trim()));
      const responseText = await response.text();

      if (!response.ok) {
        throw new Error(`API 응답 오류: ${response.status} ${responseText.slice(0, 120)}`);
      }

      const data = parseApiResponse(responseText);
      const resultCode = data.resultCode;
      const resultMsg = data.resultMsg;

      if (resultCode && resultCode !== '00') {
        throw new Error(`API 오류: ${resultMsg ?? resultCode}`);
      }

      const drugs = normalizeItems(data.items);

      if (drugs.length === 0) {
        showToast('검색 결과가 없습니다.', 'info');
        return;
      }

      if (drugs.length === 1) {
        applyDrugToSelectedTag(drugs[0]);
        return;
      }

      setSearchResults(drugs);
      showToast(`${drugs.length}개의 검색 결과가 있습니다. 적용할 약을 선택해주세요.`, 'info');
    } catch (error) {
      console.error('API 호출 에러:', error);
      showToast(error instanceof Error ? error.message : '데이터를 불러오는 중 문제가 발생했습니다.', 'error');
    } finally {
      setIsSearching(false);
    }
  };

  const handleInputChange = (event: ChangeEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    const { name, value } = event.target;
    setPriceTags((prev) => prev.map((tag) => tag.id === selectedTag.id ? { ...tag, [name]: value } : tag));
    setSearchResults([]);
  };

  const addPriceTag = () => {
    const newTag = {
      ...DEFAULT_DRUG_INFO,
      id: crypto.randomUUID(),
    };

    setPriceTags((prev) => [...prev, newTag]);
    setSelectedTagId(newTag.id);
    setSearchQuery('');
    setSearchResults([]);
  };

  const duplicatePriceTag = () => {
    const newTag = {
      ...selectedTag,
      id: crypto.randomUUID(),
    };

    setPriceTags((prev) => [...prev, newTag]);
    setSelectedTagId(newTag.id);
    setSearchResults([]);
  };

  const deletePriceTag = () => {
    if (priceTags.length === 1) {
      showToast('가격표는 최소 1개 이상 필요합니다.', 'info');
      return;
    }

    setPriceTags((prev) => {
      const nextTags = prev.filter((tag) => tag.id !== selectedTag.id);
      setSelectedTagId(nextTags[0].id);
      setSearchResults([]);
      return nextTags;
    });
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <main className="app-container">
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
        </div>

        <div className="api-key-box">
          <div className="api-key-header">
            <strong>공공데이터 API 키</strong>
            <span>{savedServiceKey || envServiceKey ? '사용 가능' : '미설정'}</span>
          </div>
          <input
            type="password"
            placeholder="API 키를 입력한 뒤 저장"
            value={serviceKeyInput}
            onChange={(event) => setServiceKeyInput(event.target.value)}
          />
          <div className="api-key-actions">
            <button type="button" onClick={saveServiceKey}>저장</button>
            <button type="button" onClick={clearServiceKey}>삭제</button>
          </div>
        </div>

        <div className="tag-list">
          <div className="tag-list-header">
            <strong>가격표 목록</strong>
            <span>{priceTags.length}개</span>
          </div>
          <div className="tag-list-items">
            {priceTags.map((tag, index) => (
              <button
                className={`${tag.id === selectedTag.id ? 'tag-list-item active' : 'tag-list-item'} ${needsTagReview(tag) ? 'needs-review' : ''}`}
                key={tag.id}
                type="button"
                onClick={() => setSelectedTagId(tag.id)}
              >
                <span>{index + 1}. {getDisplayValue(tag.itemName, '이름 미입력')}</span>
                <small>{needsTagReview(tag) ? '수정 필요' : `${tag.price}원`}</small>
              </button>
            ))}
          </div>
          <div className="tag-actions">
            <button type="button" onClick={addPriceTag}>새 가격표</button>
            <button type="button" onClick={duplicatePriceTag}>복제</button>
            <button type="button" onClick={deletePriceTag}>삭제</button>
          </div>
        </div>

        <form className="search-box" onSubmit={searchDrug}>
          <input
            type="text"
            placeholder="약 이름 검색"
            value={searchQuery}
            onChange={(event) => setSearchQuery(event.target.value)}
          />
          <button type="submit" disabled={!canSearch}>{isSearching ? '검색 중' : 'API 검색'}</button>
        </form>

        {searchResults.length > 0 && (
          <div className="search-results">
            <div className="search-results-header">
              <strong>검색 결과 선택</strong>
              <span>{searchResults.length}개</span>
            </div>
            {searchResults.map((drug) => (
              <button
                className="search-result-item"
                key={`${drug.itemName}-${drug.efcyQesitm}`}
                type="button"
                onClick={() => applyDrugToSelectedTag(drug)}
              >
                <strong>{getDisplayValue(drug.itemName ?? '', '제품명 없음')}</strong>
                <span>{getDisplayValue(stripHtml(drug.efcyQesitm ?? ''), '효능 정보 없음')}</span>
              </button>
            ))}
          </div>
        )}

        <div className="edit-form">
          <label htmlFor="itemName">약명</label>
          <textarea id="itemName" className={getInputClassName('itemName', selectedTag.itemName)} name="itemName" value={selectedTag.itemName} onChange={handleInputChange} rows={2} />

          <label htmlFor="symptoms">주요 증상</label>
          <textarea id="symptoms" className={getInputClassName('symptoms', selectedTag.symptoms)} name="symptoms" value={selectedTag.symptoms} onChange={handleInputChange} rows={3} />

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

        <button className="print-btn" type="button" onClick={handlePrint}>프린터로 출력하기</button>
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
