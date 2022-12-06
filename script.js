//store all balances and updated date in global variable
const BALANCES = PropertiesService.getScriptProperties();
PropertiesService.getScriptProperties().setProperty("firstRun", "1");
const firstRun = PropertiesService.getScriptProperties().getProperty('firstRun');
const EPS = 1e-14

function uncheck(address) {
  if (BALANCES.getProperty(address).indexOf("&") >= 0) return false;
  return true;
}

function atEdit(e) {
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const sheet = SpreadsheetApp.getActiveSheet();
  const name = sheet.getSheetName();

  if (name == 'Balance') {
    if (e.value == "Calculate") {
      const address = sheet.getRange(row, 1).getValue();
      if (col == 2) { // total balance
        const totalBalance = getTotalBalanceWithDebank(address);
        sheet.getRange(row, 3).setValue(totalBalance);
      }
      if (col == 4) { // total balance in NFT
        const totalNFTBalance = getTotalNFTBalanceWithDebank(address);
        sheet.getRange(row, 6).setValue(totalNFTBalance);
      }

      sheet.getRange(row, col).setValue("Display");
    }
  }

  if (name == 'TAMA token') {
    if (e.value == 'Calculate') {
      
      const address = sheet.getRange(row, 1).getValue();
      const tamaAmount = getTamaAmount(address, true);
      
      if (tamaAmount > EPS) {
        sheet.getRange(row, 3).setValue("Yes");
        sheet.getRange(row, 3).setBackground("green"); 
        sheet.getRange(row, 4).setValue(tamaAmount);
      } else {
        sheet.getRange(row, 3).setValue("No");
        sheet.getRange(row, 3).setBackground("red"); 
        sheet.getRange(row, 4).setValue(0);
      }
      sheet.getRange(row, col).setValue("Display");
    }
  }

  if (name == 'D2T token') {
    if (e.value == 'Calculate') {
      const address = sheet.getRange(row, 1).getValue();
      const d2tAmount = getD2TAmount(address, true);
      
      if (d2tAmount > EPS) {
        sheet.getRange(row, 3).setValue("Yes");
        sheet.getRange(row, 3).setBackground("green"); 
        sheet.getRange(row, 4).setValue(d2tAmount);
      } else {
        sheet.getRange(row, 3).setValue("No");
        sheet.getRange(row, 3).setBackground("red"); 
        sheet.getRange(row, 4).setValue(0);
      }
      sheet.getRange(row, col).setValue("Display");
    }
  }

  if (name == 'IMPT token') {
    if (e.value == 'Calculate') {
      const address = sheet.getRange(row, 1).getValue();
      const d2tAmount = getImptAmount(address, true);
      
      if (d2tAmount > EPS) {
        sheet.getRange(row, 3).setValue("Yes");
        sheet.getRange(row, 3).setBackground("green"); 
        sheet.getRange(row, 4).setValue(d2tAmount);
      } else {
        sheet.getRange(row, 3).setValue("No");
        sheet.getRange(row, 3).setBackground("red"); 
        sheet.getRange(row, 4).setValue(0);
      }
      sheet.getRange(row, col).setValue("Display");
    }
  }
}

const walletBalance = {}
// make url with apiKey, walletAddress, networkNames
function makeURL(apiKEY, walletAddress, networkNames) {
    let baseURL = `https://api.zapper.fi/v2/balances?addresses%5B%5D=${walletAddress}&api_key=${apiKEY}&bundled=false`;
    //since the networkNames is array of addresses, so we should iterate all of items to get the whole balance.
    networkNames.forEach(networkName => {
        baseURL += `&networks%5B%5D=${networkName}`
    })
    return baseURL;
}

//calculate the total balance from the response data array.
function calulateBalance(dataArray) {
  if (dataArray === "reload") {
    return "reload";
  }
    const result = {};
    let address;
    dataArray.forEach(data => {
        address = data.addresses[0];
        if (result[address] === undefined) {
            result[address] = 0;
        }
        data.totals.forEach(item => result[address] += item["balanceUSD"]);
    })
    return result[address];
}

//since the response data has the special type (not object notation, string type), we should parse the data mannually
function parseZapperResponse(data) {
    
    //since each event is seperated by '\n\n', we split the string by '\n\n'.
    const events = data.split("\n\n");

    const result = [];
    for (let i = 0; i < events.length - 2; i++) {
        const event = events[i];
        const eventData = event.slice(event.indexOf("data:") + 5);
        const object = JSON.parse(eventData.toString());
        result.push(object);
    }

    return result;
}

//fetch data from the url.
//if we fail in the first trial, we iterate the several times (maxTries) until we get the data.
function fetchData(url, maxTries) {
  let tries = 0, data, res;
  do {
    tries++;
    if (tries > 1) {
      Utilities.sleep(5000);
    }
    res = UrlFetchApp.fetch(url);
    data = res.getContentText();
  } while (!data && (tries < maxTries));

  if (!data) return "reload";
  return data;
}

function pullDataWithZapper(walletAddress, networkNames) {
  const apiKEY = "9b62ab8a-35c2-4ec8-b12a-9dfc93975f6d";  // API key from Zapper

  //we should use multiple networkNames for fetching.
  const url = makeURL(apiKEY, walletAddress, networkNames);
  
  //get response from the data.
  //if we fail in the first trial, we iterate several times (in our case 5 trial)
  try {
    const response = UrlFetchApp.fetch(url);
    const data = response.getContentText();
    // const data = fetchData(url, 5);
    return parseZapperResponse(data);
  } catch(err) {
    return "reload";
  }

}

// get total balance of a single wallet.
function getSinlgeWalletBalance(walletAddress) {
  if (!walletAddress) return 'reload';

  //all available names supported by zapper.
  const networkNames = [
      "ethereum",
      "polygon",
      "optimism",
      "gnosis",
      "binance-smart-chain",
      "fantom",
      "avalanche",
      "arbitrum",
      "celo",
      "harmony",
      "moonriver",
      "bitcoin",
      "cronos",
      "aurora",
      "evmos"];

  //get response array of object from walletaddres and networknames using zapper api
  const dataArray = pullDataWithZapper(walletAddress, networkNames);

  //get total balance from the response array.
  const totalBalance = calulateBalance(dataArray);

  return totalBalance;
}

// we can run this function to get all total balanceses.
function runTotalBalance() {

  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet(); // get active spreadsheet
  const targetSheet = activeSpreadSheet.getSheetByName('Sheet1'); //the tab where the data is going
  
  const values = targetSheet.getDataRange().getValues(); // get all active values : get wallet addresse.

  for (let row = 1; row<values.length; row++) {
    const walletAddress = values[row][0]; //get wallet address in each row.
    if (!walletAddress) continue;
    const balance = getSinlgeWalletBalance(walletAddress); // get the total balance of the wallet in USD.
    targetSheet.getRange(row, 3).setValue(balance);
    Utilities.sleep(3000);
  }
}

function getCurrentSheet(name) {
  //get active spread sheet
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  //get target sheet by its' name
  const sheet = spreadSheet.getSheetByName(name);
  return sheet;
}

//we get the row index from the address when we click the checkbox button.
function getRowIndexFromAddress(walletAddress) {
  const spreadsheet = SpreadsheetApp.getActive();
  const tf = spreadsheet.createTextFinder(walletAddress);
  const all = tf.findAll();
  
  return all[0].getRow();

}

function initAddress() {
  const sheet = getCurrentSheet('Sheet1');
  const initialValue  = sheet.getDataRange().getValues(); // get all active values : get wallet addresse.
  for (let row = 1; row < initialValue.length; row ++) {
    const walletAddress = initialValue[row][0];
    if (!walletAddress) continue;
    const balance = initialValue[row][2];
    BALANCES.setProperty(walletAddress, balance);
  }
//  Logger.log("alalalalalala----------------");
}

//get total balance checked by user.
function getTotalBalance(address, checked) {
  // Utilities.sleep(50);
  
  if (checked) {  //if the check box is selected, calculate the total balance
    const balance = getSinlgeWalletBalance(address);

    if (balance == 'reload') { // if the request time limits error occured, return the previous value
      
      return BALANCES.getProperty(address).split("&")[0];
    } else { // if we succeed, set the upgraded time into current time and return the calculated value.
      const date = new Date();
      BALANCES.setProperty(address, balance + "&" + date.toUTCString());
      return BALANCES.getProperty(address).split("&")[0];
    }
  } else { // if the check box is unchecked, return the current balance.
    return BALANCES.getProperty(address).split("&")[0];
  }
}


function getSingleWalletNftBalance(address) {
  address = "0x05b3b97c67ffcf6b441ab8b08896329674acf46b";
  if (!address) return 'reload';
  
  const apiKEY = "9b62ab8a-35c2-4ec8-b12a-9dfc93975f6d";  // API key from Zapper
  const url = `https://api.zapper.fi/v2/nft/balances/net-worth?addresses%5B%5D=${address}&api_key=${apiKEY}`;

  try {
    const data = UrlFetchApp.fetch(url).getContentText();
    const obj = JSON.parse(data);

    return obj[address];
  } catch (err) {
    return 'reload';
  }
}

function getTotalNftBalance(address, checked) {
  if (checked) {
    const nftBalance = getSingleWalletNftBalance(address);
    if (nftBalance == 'reload') {
      return BALANCES.getProperty("NFT"+address).split("&")[0];
    } else {
      const date = new Date();
      BALANCES.setProperty("NFT"+address, nftBalance + "&" + date.toUTCString());
      return nftBalance;
    }
  
  } else {
    return BALANCES.getProperty("NFT"+address).split("&")[0];
  }
}

function setTime(address) {
  if (BALANCES.getProperty(address).indexOf("&") >= 0) return BALANCES.getProperty(address).split("&")[1];
  return " ";
}

function setNftTime(address) {
  if (BALANCES.getProperty("NFT"+address).indexOf("&") >= 0) return BALANCES.getProperty("NFT"+address).split("&")[1];
  return " ";
}

//////////////////Debank API////////////////////////////
const tokenTAMA = "0x12b6893cE26Ea6341919FE289212ef77e51688c8";
const tokenD2T = "0x081071ddB7D5aF61Eb545C362ee78CC46c4bAd8f";
const tokenIMPT = "0x73Da6c5f616EB9d8A3B3a4bb920365eDEf8AF4E4";

function getAllTokenList(wallet) {
  const apiKey = "72f21c6d6d7e1d6a130e5cd201daa62349284f86";
  const url_balance = `https://pro-openapi.debank.com/v1/user/all_token_list?id=${wallet}`;
  const params = {
    'muteHttpExceptions': true,
    'headers': {
      'accept': 'application/json',
      'AccessKey': apiKey
    }
  };

  try {
    const result = JSON.parse(UrlFetchApp.fetch(url_balance, params).getContentText()); 
    return result;
  } 
  catch (err) {
    Browser.msgBox(err);
    return 'reload';
  }
}

function getTokenAmount(tokenList, tokenAddress) {
  for (const token of tokenList) {
    
    if (String(token.id).toLocaleLowerCase() == String(tokenAddress).toLocaleLowerCase()) {
      return token.amount;
    }
  }
  return 0;
}

function getTamaTokenAmount(walletAddress) {
  const tokenList = getAllTokenList(walletAddress);
  if (tokenList == 'reload') return 'reload';

  const amount = getTokenAmount(tokenList, tokenTAMA);
  return amount;
}

function getTamaAmount(address, checked) {
  if (checked) {
    const amount = getTamaTokenAmount(address);
    
    if (amount == 'reload') {
      try {
        return BALANCES.getProperty("TAMA" + address).split("&")[0];
      } catch {
        return 0;
      }
    }
    const date = new Date();
    BALANCES.setProperty("TAMA" + address, amount + "&" + date.toUTCString());
    return BALANCES.getProperty("TAMA" + address).split("&")[0];
  } else {
    try {
      const result = BALANCES.getProperty("TAMA" + address).split("&")[0];
      return result;
    }
    catch(err) {
      return 0
    }
  }
}

function getHistoryList(walletAddress, startTime, endTime) {
  
  const apiKey = "72f21c6d6d7e1d6a130e5cd201daa62349284f86";

  let historyArray = [];

  while (startTime > endTime) {

    const url_balance = `https://pro-openapi.debank.com/v1/user/history_list?id=${walletAddress}&chain_id=eth&start_time=${startTime}`;

    const params = {
      'muteHttpExceptions': true,
      'headers': {
        'accept': 'application/json',
        'AccessKey': apiKey
      }
    };

    try {
      const result = JSON.parse(UrlFetchApp.fetch(url_balance, params).getContentText());
      const histLen = result["history_list"].length;
      // Browser.msgBox(url_balance + " histLen " + histLen);
      const histArray = result["history_list"];
      // Browser.msgBox(url_balance + " histArray " + histArray.length);
      historyArray = [...historyArray, ...histArray];
      startTime = histArray[histLen - 1]["time_at"];
    } 
    catch (err) {
      // Browser.msgBox(err);
      break;
    }
  }
  return historyArray;
}

function getBlockNumber(txnHash) {
  const apiKeyEtherScan = "WYIR324M6JDCAGH79VJH6H966Y1KQRN73Q";
  const url = `https://api.etherscan.io/api?module=proxy&action=eth_getTransactionByHash&txhash=${txnHash}&apiKey=${apiKeyEtherScan}`;
  try {
    const result = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    return result["result"]["blockNumber"];
  } catch(err) {
    return 0;
  }
}

function getTokenNumber(blockNumber, tokenName) {
  blockNumber = Number.parseInt(blockNumber, 16);
  let topic0;
  if (tokenName == "D2T") {
    topic0 = "0x62e796e00a8e66154d78da76daae129635b4795a6e1b889f2caa6c5cea22ac68";
  } else {
    topic0 = "0x4d8aead3491b7eba4b5c7a65fc17e493b9e63f9e433522fc5f6a85a168fc9d36";
  }
  
  const apiKeyEtherScan = "WYIR324M6JDCAGH79VJH6H966Y1KQRN73Q";
  const url = `https://api.etherscan.io/api?module=logs&action=getLogs&fromBlock=${blockNumber}&toBlock=${blockNumber}&topic0=${topic0}&page=1&offset=1000&apikey=${apiKeyEtherScan}`;
  try {
    const result = JSON.parse(UrlFetchApp.fetch(url).getContentText());
    console.log(result["result"][0]["data"]);
    if (tokenName == "D2T") {
      return result["result"][0]["topics"][2];
    } else {
      const mid = result["result"][0]["data"].slice(0, 66);
      console.log(mid);
      return mid;
    }
  } catch(err) {
    return 0;
  }
}

function getD2TTokenAmount(walletAddress) {
  const endTime = 1666196807;
  const startTime = Math.floor(new Date().valueOf()/1000);
  const contractD2TAddress = "0x6448d7a20ece8c57212ad52b362b5c9b4feac27d";
  try {
    const historyList = getHistoryList(walletAddress, startTime, endTime);
    const hashList = historyList.filter(item => item["other_addr"].toLocaleLowerCase() == contractD2TAddress.toLocaleLowerCase());
    
    let result = 0;
    // Browser.msgBox(hashList.length);
    for (let i = 0; i<hashList.length; i++) {
      const blockNumber = getBlockNumber(hashList[i]["id"]);
      // Browser.msgBox(blockNumber);
      const tokenNumber = getTokenNumber(blockNumber, "D2T");
      // Browser.msgBox(tokenNumber);
      // Browser.msgBox(hashList[i]["cate_id"]);
      // console.log(hashList[i], Number.parseInt(tokenNumber, 16));
      result += Number.parseInt(tokenNumber, 16);

    }
    // Browser.msgBox(result);
    return result;
  } catch (err) {
    return 0;
  }
}

function getD2TAmount(address, checked) {
  if (checked) {
    const amount = getD2TTokenAmount(address);
    if (amount == 'reload') {
      try {
        return BALANCES.getProperty("D2T" + address).split("&")[0];
      } catch {
        return 0;
      }
    }
    const date = new Date();
    BALANCES.setProperty("D2T" + address, amount + "&" + date.toUTCString());
    return BALANCES.getProperty("D2T" + address).split("&")[0];
  } else {
    try {
      const result = BALANCES.getProperty("D2T" + address).split("&")[0];
      return result;
    }
    catch(err) {
      return 0
    }
  }
}

function getImptTokenAmount(walletAddress) {
  const endTime = 1664546843;
  const startTime = Math.floor(new Date().valueOf()/1000);
  const contractImptAddress = "0xf2e391f11cd1609679d03a1ac965b1d0432a7007";
  try {
    const historyList = getHistoryList(walletAddress, startTime, endTime);
    const hashList = historyList.filter(item => item["other_addr"].toLocaleLowerCase() == contractImptAddress.toLocaleLowerCase());
    console.log(hashList);
    let result = 0;
    // Browser.msgBox(hashList.length);
    for (let i = 0; i<hashList.length; i++) {
      const blockNumber = getBlockNumber(hashList[i]["id"]);
      // Browser.msgBox(blockNumber);
      const tokenNumber = getTokenNumber(blockNumber, "IMPT");

      console.log("tokenNumber", tokenNumber);
      // Browser.msgBox(tokenNumber);
      // Browser.msgBox(hashList[i]["cate_id"]);
      // console.log(hashList[i], Number.parseInt(tokenNumber, 16));
      result += Number.parseInt(tokenNumber, 16);

    }
    // Browser.msgBox(result);
    return result;
  } catch (err) {
    return 0;
  }

}

function getImptAmount(address, checked) {
    if (checked) {
    const amount = getImptTokenAmount(address);
    console.log(amount);
    if (amount == 'reload') {
      try {
        return BALANCES.getProperty("IMPT" + address).split("&")[0];
      } catch {
        return 0;
      }
    }
    const date = new Date();
    BALANCES.setProperty("IMPT" + address, amount + "&" + date.toUTCString());
    return BALANCES.getProperty("IMPT" + address).split("&")[0];
  } else {
    try {
      const result = BALANCES.getProperty("IMPT" + address).split("&")[0];
      return result;
    }
    catch(err) {
      return 0
    }
  }
}

function displayPrompt(message) {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(message);
  return result.getResponseText();
}

function calculateWholeTama() {

  let startAddressId = 0;
  let endAddressId = 2;
  startAddressId= displayPrompt("Please input the starting index of address");
  endAddressId = displayPrompt("Please input the ending index of address");

  const activeSpreadSheet = SpreadsheetApp.getActiveSpreadsheet(); // get active spreadsheet
  const targetSheet = activeSpreadSheet.getSheetByName('TAMA token'); //the tab where the data is going
  
  const values = targetSheet.getDataRange().getValues(); // get all active values : get wallet addresse.

  
  for (let row = Math.max(startAddressId, 2); row<=Math.min(values.length, endAddressId); row++) {
    const walletAddress = values[row-1][0]; //get wallet address in each row.
    if (!walletAddress) continue;
    const balance = getTamaTokenAmount(walletAddress); 
    
    if (balance < EPS) {
      targetSheet.getRange(row, 3).setValue("No");
      targetSheet.getRange(row, 3).setBackground("red"); 
      targetSheet.getRange(row, 4).setValue(0);
    } else {
      targetSheet.getRange(row, 3).setValue("Yes");
      targetSheet.getRange(row, 3).setBackground("Green"); 
      targetSheet.getRange(row, 4).setValue(balance);
    }
    targetSheet.getRange(row, 2).setValue("Display");
    Utilities.sleep(10);
  }
}


function calculateWholeD2T() {

  let startAddressId = 0;
  let endAddressId = 2;
  startAddressId= displayPrompt("Please input the starting index of address");
  endAddressId = displayPrompt("Please input the ending index of address");


  const targetSheet = SpreadsheetApp.getActiveSheet();
  
  const values = targetSheet.getDataRange().getValues(); // get all active values : get wallet addresse.

  for (let row = Math.max(startAddressId, 2); row<=Math.min(values.length, endAddressId); row++) {
    const walletAddress = values[row-1][0]; //get wallet address in each row.
    if (!walletAddress) continue;
    const balance = getD2TTokenAmount(walletAddress); 
    
    if (balance < EPS) {
      targetSheet.getRange(row, 3).setValue("No");
      targetSheet.getRange(row, 3).setBackground("red"); 
      targetSheet.getRange(row, 4).setValue(0);
    } else {
      targetSheet.getRange(row, 3).setValue("Yes");
      targetSheet.getRange(row, 3).setBackground("Green"); 
      targetSheet.getRange(row, 4).setValue(balance);
    }
    targetSheet.getRange(row, 2).setValue("Display");
    Utilities.sleep(10);
  }
}

function calculateWholeImpt() {
  let startAddressId = 0;
  let endAddressId = 2;
  startAddressId= displayPrompt("Please input the starting index of address");
  endAddressId = displayPrompt("Please input the ending index of address");

  const targetSheet = SpreadsheetApp.getActiveSheet();
  
  const values = targetSheet.getDataRange().getValues(); // get all active values : get wallet addresse.

  for (let row = Math.max(startAddressId, 2); row<=Math.min(values.length, endAddressId); row++) {
    const walletAddress = values[row-1][0]; //get wallet address in each row.
    if (!walletAddress) continue;
    const balance = getImptTokenAmount(walletAddress); 
    
    if (balance < EPS) {
      targetSheet.getRange(row, 3).setValue("No");
      targetSheet.getRange(row, 3).setBackground("red"); 
      targetSheet.getRange(row, 4).setValue(0);
    } else {
      targetSheet.getRange(row, 3).setValue("Yes");
      targetSheet.getRange(row, 3).setBackground("Green"); 
      targetSheet.getRange(row, 4).setValue(balance);
    }
    targetSheet.getRange(row, 2).setValue("Display");
    Utilities.sleep(10);
  }
}

function getTotalBalanceWithDebank(walletAddress) {
  const apiKey = "72f21c6d6d7e1d6a130e5cd201daa62349284f86";
  const url_balance = `https://pro-openapi.debank.com/v1/user/total_balance?id=${walletAddress}`;
  const params = {
    'muteHttpExceptions': true,
    'headers': {
      'accept': 'application/json',
      'AccessKey': apiKey
    }
  };

  try {
    const result = JSON.parse(UrlFetchApp.fetch(url_balance, params).getContentText()).total_usd_value; 
    return result;
  } 
  catch (err) {
    Browser.msgBox(err);
    return '0';
  }
}

function getTotalBalanceWhole() {
  let startAddressId = 0;
  let endAddressId = 2;
  startAddressId= displayPrompt("Please input the starting index of address");
  endAddressId = displayPrompt("Please input the ending index of address");


  const targetSheet = SpreadsheetApp.getActiveSheet();
  
  const values = targetSheet.getDataRange().getValues(); // get all active values : get wallet addresse.

  for (let row = Math.max(startAddressId, 2); row<=Math.min(values.length, endAddressId); row++) {
    const walletAddress = values[row-1][0]; //get wallet address in each row.
    if (!walletAddress) continue;
    const balance = getTotalBalanceWithDebank(walletAddress); 
    targetSheet.getRange(row, 3).setValue(balance);
    targetSheet.getRange(row, 2).setValue("Display");
    Utilities.sleep(10);
  }
}
  