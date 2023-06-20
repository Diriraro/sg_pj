const readline = require('readline');
const axios = require('axios');
const axiosCookieJarSupport = require('axios-cookiejar-support').default;
const tough = require('tough-cookie');
const cheerio = require('cheerio');
const FormData = require('form-data');
const XLSX = require('xlsx');

// 쿠키 관리를 위해 axios에 axiosCookieJarSupport 확장 기능 추가
axiosCookieJarSupport(axios);

// 새로운 쿠키 저장소 인스턴스 생성
const cookieJar = new tough.CookieJar();

// fetchDataGET(url, headers): 사용자 지정 헤더를 포함하고 쿠키를 관리하여 GET 요청을 보내는 함수
async function fetchDataGET(url, headers) {
  try {
    const response = await axios.get(url, {
      headers,
      jar: cookieJar, // 쿠키 저장소 설정
      withCredentials: true // 쿠키를 요청에 포함
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with GET:", error);
    return "";
  }
}

// fetchDataPOST(url, formData, headers): 사용자 지정 헤더를 포함하고 쿠키를 관리하여 FormData를 포함한 POST 요청을 보내는 함수
async function fetchDataPOST(url, formData, headers) {
  try {
    const response = await axios.post(url, formData, {
      headers: { ...headers, ...formData.getHeaders() }, // 사용자 지정 헤더와 FormData 헤더를 병합
      jar: cookieJar, // 쿠키 저장소 설정
      withCredentials: true // 쿠키를 요청에 포함
    });
    return response.data;
  } catch (error) {
    console.error("Error fetching data with POST:", error);
    return "";
  }
}

// 기존 parseData와 saveToExcel 함수는 동일하게 유지됩니다.

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
  
            // 사용자 지정 헤더 정의
            const customHeaders = {
                'User-Agent': 'My-Custom-User-Agent',
                'X-Custom-Header': 'CustomHeaderValue'
            };

            // 기존 main 함수는 동일하게 유지되지만 fetchDataGET과 fetchDataPOST 호출에 사용자 지정 헤더 추가
            const data1 = await fetchDataGET(url1, customHeaders);
            const parsedData1 = parseData(data1);
            saveToExcel(parsedData1, 'output1.xlsx');

            const data2 = await fetchDataPOST(url2, formData, customHeaders);
            const parsedData2 = parseData(data2);
            saveToExcel(parsedData2, 'output2.xlsx');

            const data3 = await fetchDataGET(url3, customHeaders);
            const parsedData3 = parseData(data3);
            saveToExcel(parsedData3, 'output3.xlsx');
        });
      });
    });

    rl.question("키를 입력하면 프로그램이 종료됩니다.", () => {
        rl.close();
    });
  }

main();
