// Node.js의 readline, axios, axios-cookiejar-support, tough-cookie, cheerio, form-data, xlsx 모듈을 불러옵니다.
const readline = require('readline');
const axios = require('axios');
const tough = require('tough-cookie');
const axiosCookieJarSupport = require('axios-cookiejar-support').default;
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

async function httpConnect(actionId, method, url, bodyText, retryCount = 0) {
    var resultData = "";
    var fullUrl = url.indexOf("://") >= 0 ? url : this.hostURL + url;
  
    try {
      for (var i = 0; i <= retryCount; i++) {
        console.error(`[${i + 1}] 번째 실행 : ${fullUrl}`);
  
        this.userErrorMessage = "";
  
        const config = {
          url: fullUrl,
          method: method,
          jar: cookieJar,
          withCredentials: true,
        };

        const customHeadersGet = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Origin': 'https://prm.iniwedding.com'
        };
    
        const customHeadersPost = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Origin': 'https://prm.iniwedding.com',
            'Content-Type': 'application/x-www-form-urlencoded'
        };
  
        if (method.toLowerCase() === 'post' && bodyText) {
          config.data = bodyText;
          config.header = customHeadersPost;
        } else {
            config.header = customHeadersGet;
        }
  
        try {
          const response = await axios(config);
  
          if (response.status === 200) {
            this.setCookieString(response.headers['set-cookie']);
            resultData = response.data;
          } else if (response.status === 302 && response.headers.location) {
            return await this.httpConnect(`${actionId}_3`, "GET", response.headers.location, "", 2);
          } else {
            this.userError = `${actionId}_1::${response.status}`;
            return false;
          }
        } catch (error) {
          if (i < retryCount) {
            await new Promise(resolve => setTimeout(resolve, 500));
            continue;
          }
          this.userError = `${actionId}_2::${error.response.status}`;
          return false;
        }
  
        if (!resultData) {
          resultData = Buffer.from(response.data).toString('hex');
        }
  
        break;
      }
    } catch (error) {
      console.error(`[${actionId}] 실패 - 재시도 횟수 초과`);
      throw error;
    }
  
    return resultData;
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

function checkBody(body) {

  if(body.srcType == "0"){
    body.srcType = "RMON";
  } else if (body.srcType == "1") {
    body.srcType = "WMON";
  } else {
    return false;
  }
  if(body.stDate.length != 10 || body.edDate.length != 10) {
    return false;
  }
  return body;
}

// 메인 함수: url에 대해 GET 요청과 POST 요청을 보내고, 그 결과를 파싱하고, 엑셀 파일로 저장합니다.
(async function main() {
  const baseURL = "https://prm.iniwedding.com";
  const url3 = "https://example3.com";

  const customHeadersGet = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Origin': 'https://prm.iniwedding.com'
  };

  const customHeadersPost = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Origin': 'https://prm.iniwedding.com',
    'Content-Type': 'application/x-www-form-urlencoded'
  };

  let body = {
    srcType: await promptInput("리허설 검색은 0, 예식일 검색은 1 중 하나를 입력하세요: "),
    stDate: await promptInput("검색 시작할 날짜를 yyyy-MM-dd 형식으로 입력하세요: "),
    edDate: await promptInput("검색 끝낼 날짜를 yyyy-MM-dd 형식으로 입력하세요: ")
  };
  
  body = checkBody(body);
  
  if (!body){
    rl.question("잘못된 값이 입력되어 프로그램이 종료됩니다. 아무키나 누르세요.", () => {
      rl.close();
      process.exit();
    });
  }

  const data1 = await httpConnect('main', 'GET',baseURL,'');
  // if(data1) {
  //   const parsedData1 = parseData(data1);
  //   saveToExcel(parsedData1, 'output1.xlsx');
  // }
  console.log(data1);
  console.log(cookieJar);
  const loginBody = {
    mb_id: "bus151a00009",
    mb_password: "xkdla"
  } 
//   const data2 = await fetchDataPOST(baseURL + "/bbs/login_check.php?", loginBody, customHeadersPost);
//   console.log(data2);


  // if(data2) {
  //   const parsedData2 = parseData(data2);
  //   saveToExcel(parsedData2, 'output2.xlsx');
  // }

  // const data3 = await fetchDataGET(url3, customHeaders);
  // if(data3) {
  //   const parsedData3 = parseData(data3);
  //   saveToExcel(parsedData3, 'output3.xlsx');
  // }

  rl.question("키를 입력하면 프로그램이 종료됩니다.", () => {
    rl.close();
  });
})();
