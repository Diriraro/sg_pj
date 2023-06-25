// Node.js의 readline, axios, axios-cookiejar-support, tough-cookie, cheerio, form-data, xlsx 모듈을 불러옵니다.
const readline = require('readline');
const axios = require('axios');
const tough = require('tough-cookie');
const axiosCookieJarSupport = require('axios-cookiejar-support').default;
const cheerio = require('cheerio');
const FormData = require('form-data');
const XLSX = require('xlsx');
const credentials = require('./credentials.js');
require("./String.js");

let tempData = {};
tempData.message = "";

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
    let response; 
    // var fullUrl = url.indexOf("://") >= 0 ? url : this.hostURL + url;
    // console.log(url);

    try {
      for (var i = 0; i <= retryCount; i++) {
        console.error(`${actionId}:: [${i + 1}] 번째 실행 : ${url}`);
  
        const config = {
          url: url,
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
          config.headers = customHeadersPost;
        } else {
            config.headers = customHeadersGet;
        }

        try {
            response = await axios(config);
  
          if (response.status === 200) {
            resultData = response.data;
          } else if ((response.status === 302 || response.status === 301) && response.headers.location) {
            return await httpConnect(`${actionId}_3`, "GET", response.headers.location, "");
          } else {
            tempData = response;
            console.log(`${actionId} response status :: ${response.status}`);
            console.log(`${actionId} response url :: ${response.config.url}`);
            return false;
          }
        } catch (error) {
          if (i < retryCount) {
            await new Promise(resolve => setTimeout(resolve, 500));
            continue;
          }
          console.log(`${actionId} :: ${error.message}`);
          if(error.response.status) console.log( `${actionId}_2::${error.response.status}`);
          tempData = error;
          logOut();
          return false;
        }
  
        break;
      }
    } catch (error) {
      console.error(`[${actionId}] 실패 - 재시도 횟수 초과`);
      console.error(error.message);
      tempData = error;
      logOut();
      return false;
    }
  
    return resultData;
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

function logOut(){
  httpConnect('MAINPAGE', 'GET', 'https://prm.iniwedding.com/bbs/logout.php','');
}

function errorCatch(err){

  console.log('=======================================')
  console.log(err.message)
  console.log('=======================================')
  
  rl.question("에러발생. 프로그램이 종료됩니다. 아무키나 누르세요.", () => {
    rl.close();
    process.exit();
  });
}

function checkBody(body) {
  if(body.srcType == "0"){
    body.srcType = "RMON";
  } else if (body.srcType == "1") {
    body.srcType = "WMON";
  } else {
    return false;
  }

  if (body.stDate.length == 0 || body.edDate.length == 0) {
    console.log('날짜가 입력되지 않음');
    const now = new Date();
    const stDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7).toISOString().slice(0,10);
    const edDate = new Date(now.getFullYear(), now.getMonth(), now.getDate()).toISOString().slice(0,10);
    console.log(`시작날짜 ::${stDate} / 종료날짜 ::${edDate}로 검색`);
    body.stDate = stDate;
    body.edDate = edDate;
  }else if(body.stDate.length != 10 || body.edDate.length != 10) {
    console.log('시작날짜/끝낼날짜가 빈값, 또는 10자리의 값이 아닙니다. 에러발생.')
    return false;
  } 
  return body;
}

// 메인 함수: url에 대해 GET 요청과 POST 요청을 보내고, 그 결과를 파싱하고, 엑셀 파일로 저장합니다.
(async function main() {
  const baseURL = "https://prm.iniwedding.com";

  let body = {
    srcType: await promptInput("리허설(촬영용) 검색은 0, 예식일 검색은 1 중 하나를 입력하세요: "),
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

  resultData = await httpConnect('MAINPAGE', 'GET', baseURL,'');
  if (resultData === false) errorCatch(tempData);
  
  const loginBody = `mb_id=${credentials.id}&mb_password=${credentials.password}`;
  
  resultData = await httpConnect('LOGIN', 'POST', baseURL + "/bbs/login_check.php?", loginBody);
  if (resultData === false) errorCatch(tempData);
  
  if (resultData.indexOf("location.replace('/home')") == -1) {
    tempData.message = '로그인 실패 / 코드수정필요';
    errorCatch(tempData);
  }
  
  resultData = await httpConnect('Login2', 'GET', baseURL + '/home','');
  if (resultData === false) errorCatch(tempData);
  let chkPage = resultData.grap('<title>', '</title>');
  if(chkPage != '아이니웨딩 PRM') {
    tempData.message = '로그인 실패 2 / 코드수정필요';
    errorCatch(tempData);
  }

  resultData = await httpConnect('Order1', 'GET', baseURL + '/Order/OrderList.php','');
  if (resultData === false) errorCatch(tempData);
  chkPage = resultData.grap('<title>', '</title>');
  if(chkPage != '발주현황') {
    tempData.message = '발주현황페이지 진입실패';
    errorCatch(tempData);
  }

  let isPagingEnd = false;
  let paging = 1;
  let codeArr = [];
  let parsingArr = [];
  while(true) {
    let postOrderBody = "";
    postOrderBody += "button_flag="
    postOrderBody += "&sort=CP_PlacingDateTime"
    postOrderBody += "&pages=" + paging;
    postOrderBody += "&ContractPlacing_Code="
    postOrderBody += "&OrderType=O"
    postOrderBody += "&dateFrmName="
    postOrderBody += "&idxno="
    postOrderBody += "&ContractName="
    postOrderBody += "&ShContractPlacing_Code="
    postOrderBody += "&CP_GoodName="
    postOrderBody += "&SearchMon=" + body.srcType;
    postOrderBody += "&SDAY=" + body.stDate;
    postOrderBody += "&EDAY=" + body.edDate;

    resultData = await httpConnect(`Order2_${paging}`, 'POST', baseURL + "/Order/OrderList.php", postOrderBody);
    if (resultData === false) errorCatch(tempData);
    chkPage = resultData.grap('<title>', '</title>');
    if(chkPage != '발주현황') {
      tempData.message = '발주현황페이지 진입실패';
      errorCatch(tempData);
    }
    let isData = resultData.grap(`<tr class='ConteTR' style="height:32px;">`, '</tr>');
    if (isData == "" || !isData) break;

    let trCnt = 0;
    while(true) {
      let trData = resultData.grap(`<tr class='ConteTR' style="height:32px;">`, '</tr>', trCnt);
      
      if(trCnt == 0 && (trData == "" || !trData)){
        isPagingEnd = true;
        break;
      } 

      let chkIsOk = trData.grap("<td class='ConteTD_End_C'>", '</td>').removeHtmlTagAll().trim();
      if (chkIsOk == "확인") {
        isPagingEnd = true;
        break;
      }

      let dataCd = trData.grap("<td class='ConteTD_C'>", '</td>', 1);
      codeArr.push(dataCd);
      trCnt++;
    }

    if(isPagingEnd) break;
    paging++;
  }

  for (let cnt = 0; cnt < codeArr.length; cnt++) {
    let itemCd = codeArr[cnt];
    let postOrderBody = "";
    postOrderBody += "button_flag="
    postOrderBody += "&sort=CP_PlacingDateTime"
    postOrderBody += "&pages=1";
    postOrderBody += "&ContractPlacing_Code=" + itemCd
    postOrderBody += "&OrderType=O"
    postOrderBody += "&dateFrmName="
    postOrderBody += "&idxno="
    postOrderBody += "&ContractName="
    postOrderBody += "&ShContractPlacing_Code="
    postOrderBody += "&CP_GoodName="
    postOrderBody += "&SearchMon=" + body.srcType;
    postOrderBody += "&SDAY=" + body.stDate;
    postOrderBody += "&EDAY=" + body.edDate;

    resultData = await httpConnect(`Contract_${cnt}`, 'POST', baseURL + "/Order/ContractOptionFax_Fixed.php", postOrderBody);
    if (resultData === false) errorCatch(tempData);
    if(resultData.replaceEntities().indexOf('발   주   서') == -1) {
      tempData.message = '발주서 진입 실패';
      errorCatch(tempData);
    }

    parsingArr.push(resultData.replaceEntities());
  }

  /**
   * 1. 발주현황 입력값 추가 > post통신
   *  1-1. 입력받은 데이터 체크 > 해당 값으로 검색
   * 2. 페이징 확인 힘듬 > while문으로 데이터 수집
   *  2-1. 각 td 마다 전부 '발주서 인쇄' 통신으로 추가 데이터 수집 필요
   *  2-2. 수집 하다가 마지막 td가 '확인' 일 경우 반복 중지
   *  2-3. 다음 페이지 table이 없을경우 반복 중지
   * -- OK
   * 3. 가져온 데이터 파싱
   * 3-1. 입력값에 따라 다른 값 파싱 필요
   * 4. 엑셀파일 생성
   */

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
