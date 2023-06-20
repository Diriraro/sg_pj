const readline = require('readline');
const axios = require('axios');
const axiosCookieJarSupport = require('axios-cookiejar-support').default;
const tough = require('tough-cookie');
const cheerio = require('cheerio');
const FormData = require('form-data');
const XLSX = require('xlsx');

axiosCookieJarSupport(axios);
const cookieJar = new tough.CookieJar();

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// body를 받는 함수
function promptInput(question) {
  return new Promise((resolve) => {
      rl.question(question, (answer) => {
          resolve(answer);
      });
  });
}

async function fetchDataPOST(url, body1, body2, body3, headers) {
  // FormData 인스턴스 생성 및 body 설정
  const formData = new FormData();
  formData.append("body1", body1);
  formData.append("body2", body2);
  formData.append("body3", body3);
  
  try {
    const response = await axios.post(url, formData, {
      headers: { ...headers, ...formData.getHeaders() },
      jar: cookieJar,
      withCredentials: true
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with POST:", error);
    return "";
  }
}

async function fetchDataGET(url, headers) {
  try {
    const response = await axios.get(url, {
      headers,
      jar: cookieJar,
      withCredentials: true
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with GET:", error);
    return "";
  }
}

// HTML 파싱 및 원하는 데이터 추출
function parseData(html) {
  const $ = cheerio.load(html);
  const parsedData = [];

  // 원하는 데이터 추출(여기는 추출 코드를 추가하셔야 합니다.)
  // 예: $('div.className').each((index, element) => {
  //   parsedData.push($(element).text().trim());
  // });

  return parsed;
}

// 데이터를 엑셀에 저장
function saveToExcel(data, fileName) {
  const wb = XLSX.utils.book_new(); // 새 워크북 생성
  const ws = XLSX.utils.json_to_sheet(data); // JSON 데이터를 워크시트로 변환
  XLSX.utils.book_append_sheet(wb, ws, "Sheet 1"); // 워크북에 워크시트 추가
  XLSX.writeFile(wb, fileName); // 엑셀 파일로 저장
}

(async function main() {
  // url1, url2, url3를 내부적으로 지정합니다.
  const url1 = "https://example1.com";
  const url2 = "https://example2.com";
  const url3 = "https://example3.com";

  // 사용자 지정 헤더 정의
  const customHeaders = {
    'User-Agent': 'My-Custom-User-Agent',
    'X-Custom-Header': 'CustomHeaderValue'
  };

  // Body 값을 입력 받습니다.
  const body1 = await promptInput("Body1 값을 입력하세요: ");
  const body2 = await promptInput("Body2 값을 입력하세요: ");
  const body3 = await promptInput("Body3 값을 입력하세요: ");

  const data1 = await fetchDataGET(url1, customHeaders);
  const parsedData1 = parseData(data1);
  saveToExcel(parsedData1, 'output1.xlsx');

  const data2 = await fetchDataPOST(url2, body1, body2, body3, customHeaders);
  const parsedData2 = parseData(data2);
  saveToExcel(parsedData2, 'output2.xlsx');

  const data3 = await fetchDataGET(url3, customHeaders);
  const parsedData3 = parseData(data3);
  saveToExcel(parsedData3, 'output3.xlsx');

  // 사용자 키 입력을 대기하고 프로그램을 종료하는 코드를 추가합니다.
  rl.question("키를 입력하면 프로그램이 종료됩니다.", () => {
    rl.close();
  });
})();
