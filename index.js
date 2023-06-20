const readline = require('readline');
const axios = require('axios');
const cheerio = require('cheerio');
const FormData = require('form-data');
const XLSX = require('xlsx');

// fetchDataGET(url): URL에서 데이터를 가져오는 함수 (GET 요청)
async function fetchDataGET(url) {
  try {
    const response = await axios.get(url);
    return response.data;
  } catch (error) {
    console.error("Error fetching data with GET:", error);
    return "";
  }
}

// fetchDataPOST(url, formData): URL로 body값에 form-data를 포함한 POST 요청을 보내는 함수
async function fetchDataPOST(url, formData) {
  try {
    const response = await axios.post(url, formData, {
      headers: formData.getHeaders()
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with POST:", error);
    return "";
  }
}

// parseData(html): 웹 페이지 데이터를 파싱하는 함수
// HTML 문자열을 파싱하여 원하는 데이터를 추출한 후 배열 형태로 반환합니다.
function parseData(html) {
  const $ = cheerio.load(html);
  const dataList = [];

  // 원하는 데이터 추출(여기서는 h1 태그의 텍스트)
  $('h1').each((index, element) => { 
    dataList.push({ title: $(element).text() });
  });

  return dataList;
}

// saveToExcel(dataList, outputFilename): JSON 데이터를 엑셀 파일로 저장하는 함수
// dataList 파라미터에 주어진 배열 형태의 데이터를 엑셀 파일로 저장합니다.
function saveToExcel(dataList, outputFilename) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(dataList);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  XLSX.writeFile(workbook, outputFilename);
}

// Main function
async function main() {
  // 내부적으로 정의한 URL
  const url1 = "https://example1.com";
  const url2 = "https://example2.com";
  const url3 = "https://example3.com";
  const outputFilename = "output.xlsx";

  // 라인 리더 설정
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  
  // 사용자로부터 3개의 POST body 값을 입력 받기
  rl.question('Please enter the 1st value for POST body data: ', async function (postValue1) {
    rl.question('Please enter the 2nd value for POST body data: ', async function (postValue2) {
      rl.question('Please enter the 3rd value for POST body data: ', async function (postValue3) {
        rl.close();

        // POST body 데이터로 사용할 FormData 생성
        const formData = new FormData();
        formData.append('key1', postValue1);
        formData.append('key2', postValue2);
        formData.append('key3', postValue3);

        // GET/POST 요청을 원하는 순서대로 호출
        const data1 = await fetchDataGET(url1);
        const parsedData1 = parseData(data1);
        saveToExcel(parsedData1, 'output1.xlsx');
        
        const data2 = await fetchDataPOST(url2, formData);
        const parsedData2 = parseData(data2);
        saveToExcel(parsedData2, 'output2.xlsx');

        const data3 = await fetchDataGET(url3);
        const parsedData3 = parseData(data3);
        saveToExcel(parsedData3, 'output3.xlsx');
      });
    });
  });
}

main();
