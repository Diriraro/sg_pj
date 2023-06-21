// Node.js의 readline, axios, axios-cookiejar-support, tough-cookie, cheerio, form-data, xlsx 모듈을 불러옵니다.
const readline = require('readline');
const axios = require('axios');
const axiosCookieJarSupport = require('axios-cookiejar-support').default;
const tough = require('tough-cookie');
const cheerio = require('cheerio');
const FormData = require('form-data');
const XLSX = require('xlsx');

// axios에 cookie jar 기능을 추가해줍니다.
axiosCookieJarSupport(axios);
// 쿠키를 저장할 수 있는 cookie jar를 생성합니다.
const cookieJar = new tough.CookieJar();

// 사용자로부터 입력을 받을 수 있는 readline 인터페이스를 생성합니다.
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// 사용자로부터 질문을 받아, 그에 대한 답을 Promise 형태로 반환하는 함수입니다.
function promptInput(question) {
  return new Promise((resolve) => {
    rl.question(question, (answer) => {
        resolve(answer);
    });
  });
}

// 주어진 url에 POST 요청을 보내, 응답을 받아옵니다. 요청 본문은 body 객체를 form data 형식으로 변환해 전달하며,
// 요청 헤더는 headers 객체를 사용합니다.
async function fetchDataPOST(url, body, headers) {
  const formData = new FormData();
  Object.keys(body).forEach(key => {
    formData.append(key, body[key]);
  });
  
  try {
    const response = await axios.post(url, formData, {
      headers: { ...headers, ...formData.getHeaders() },
      jar: cookieJar,
      withCredentials: true
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with POST:", error);
    return null;
  }
}

// 주어진 url에 GET 요청을 보내, 응답을 받아옵니다. 요청 헤더는 headers 객체를 사용합니다.
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
    return null;
  }
}

// HTML 문자열을 파싱해 원하는 데이터를 추출하는 함수입니다. 원하는 데이터의 형태에 따라 이 함수를 수정해야 합니다.
function parseData(html) {
  const $ = cheerio.load(html);
  const parsedData = [];

  // Modify this part to match the structure of the data you are trying to extract
  // 원하는 데이터 추출(여기는 추출 코드를 추가하셔야 합니다.)
  // 예: $('div.className').each((index, element) => {
  //   parsedData.push($(element).text().trim());
  // });


  return parsedData;
}

// 데이터를 엑셀 파일로 저장하는 함수입니다.
function saveToExcel(data, fileName) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Sheet 1");
  XLSX.writeFile(wb, fileName);
}

// 메인 함수: url에 대해 GET 요청과 POST 요청을 보내고, 그 결과를 파싱하고, 엑셀 파일로 저장합니다.
(async function main() {
  const url1 = "https://example1.com";
  const url2 = "https://example2.com";
  const url3 = "https://example3.com";

  const customHeaders = {
    'User-Agent': 'My-Custom-User-Agent',
    'X-Custom-Header': 'CustomHeaderValue'
  };

  const body = {
    body1: await promptInput("Body1 값을 입력하세요: "),
    body2: await promptInput("Body2 값을 입력하세요: "),
    body3: await promptInput("Body3 값을 입력하세요: "),
  };

  const data1 = await fetchDataGET(url1, customHeaders);
  if(data1) {
    const parsedData1 = parseData(data1);
    saveToExcel(parsedData1, 'output1.xlsx');
  }

  const data2 = await fetchDataPOST(url2, body, customHeaders);
  if(data2) {
    const parsedData2 = parseData(data2);
    saveToExcel(parsedData2, 'output2.xlsx');
  }

  const data3 = await fetchDataGET(url3, customHeaders);
  if(data3) {
    const parsedData3 = parseData(data3);
    saveToExcel(parsedData3, 'output3.xlsx');
  }

  rl.question("키를 입력하면 프로그램이 종료됩니다.", () => {
    rl.close();
  });
})();
