import { readFileSync } from 'node:fs';

const envText = readFileSync('.env', 'utf8');
const serviceKey = envText
  .split(/\r?\n/)
  .find((line) => line.trim().startsWith('VITE_PUBLIC_DATA_SERVICE_KEY='))
  ?.split('=')
  .slice(1)
  .join('=')
  .trim();

if (!serviceKey) {
  console.error('VITE_PUBLIC_DATA_SERVICE_KEY를 .env에서 찾지 못했습니다.');
  process.exit(1);
}

const decodeServiceKey = (key) => {
  try {
    return decodeURIComponent(key.trim());
  } catch {
    return key.trim();
  }
};

const params = new URLSearchParams({
  itemName: '타이레놀',
  type: 'json',
  pageNo: '1',
  numOfRows: '3',
});

console.log('API 키 길이:', serviceKey.length);
console.log('API 키 형식:', /%[0-9A-F]{2}/i.test(serviceKey) ? 'Encoding으로 보임' : 'Decoding으로 보임');
console.log('요청 테스트: 타이레놀');

const cases = [
  ['https + ServiceKey + 현재키', 'https', 'ServiceKey', serviceKey.trim()],
  ['https + serviceKey + 현재키', 'https', 'serviceKey', serviceKey.trim()],
  ['https + ServiceKey + 디코딩키', 'https', 'ServiceKey', decodeServiceKey(serviceKey)],
  ['https + serviceKey + 디코딩키', 'https', 'serviceKey', decodeServiceKey(serviceKey)],
  ['http + ServiceKey + 현재키', 'http', 'ServiceKey', serviceKey.trim()],
  ['http + serviceKey + 현재키', 'http', 'serviceKey', serviceKey.trim()],
];

for (const [label, scheme, keyName, keyValue] of cases) {
  const url = `${scheme}://apis.data.go.kr/1471000/DrbEasyDrugInfoService/getDrbEasyDrugList?${keyName}=${keyValue}&${params.toString()}`;
  const response = await fetch(url);
  const text = await response.text();

  console.log('\\n---', label, '---');
  console.log('HTTP 상태:', response.status, response.statusText);
  console.log('응답 앞부분:');
  console.log(text.slice(0, 500));
}
