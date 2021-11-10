enum Func {
    All = "all",
    chart = "chart",
    OTCchart = "OTCchart",
    time = "time",
    stockId = "stockId", 
    stockName = "stockName",
    OpenPrice = "OpenPrice",
    ClosingPrice = "ClosingPrice",
    volume = "volume"

}

type stock = {
    closingPrice: number,
    high: number,
    low: number,
    openPrice: number,
    changeRange: string,
    date: string,
    foreignInvestors: number,
    investmentTrust: number,
    stockId: number,
    stockName: string,
    volume: number,
    yesterdayClose: number,
}
//console.log("hello world");
type DoGet = GoogleAppsScript.Events.DoGet & {
    parameter:{
       func:Func;
    }
}


function onOpen(){
    console.log("test open")
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("script").addItem("Test","test").addToUi();


    // var response = UrlFetchApp.fetch('https://google.com/');
    // Logger.log(response.getContentText());

    // const fetch = require('isomorphic-fetch')

    // fetch('http://example.com')
    //   .then(res => res.text())
    //   .then(text => console.log(text))

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("stockData");

    function getData() {
        //NOTE google Script只能用他們自己的UrlFetchApp 目前已不使用此方法

        //   var response = UrlFetchApp.fetch('https://google.com/');
        //   Logger.log(response.getContentText());
        // const cors = 'https://cors-anywhere.herokuapp.com/'; // use cors-anywhere to fetch api data
        // const url = 'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date=20210526&type=ALL'; // origin api url
        // sheet.getRange("A7").setValue("2222"); 
        // let res = await fetch(`${cors}${url}`);
        // let text = await res.json();
        // Logger.log(text.fields7[0]);
        // console.log("text.fields7[0]");
        // sheet.getRange("A7").setValue(text.fields7[0]); 
      
    }
    getData();
}

//依日期建立新表格
function createsheet(newDate){
    let ss = SpreadsheetApp.getActiveSpreadsheet();
 
    ss.insertSheet("stock" + newDate);

    ss.getRange(`A1`).setValue(`date`); 
    ss.getRange(`B1`).setValue(`stockId`); 
    ss.getRange(`C1`).setValue(`stockName`);
    ss.getRange(`D1`).setValue(`yesterdayClose`);  
    ss.getRange(`E1`).setValue(`openPrice`); 
    ss.getRange(`F1`).setValue(`closePrice`); 
    ss.getRange(`G1`).setValue(`high`); 
    ss.getRange(`H1`).setValue(`low`); 
    ss.getRange(`I1`).setValue(`changeRange`); 
    ss.getRange(`J1`).setValue(`volume`); 
    ss.getRange(`K1`).setValue(`foreignInvestors`); 
    ss.getRange(`L1`).setValue(`investmentTrust`); 
    console.log("create")
    
}

function point(){
    var TSCpoint = UrlFetchApp.fetch(`https://www.twse.com.tw/exchangeReport/MI_5MINS_INDEX?response=json&date=20211020`);
    //Logger.log(responseTSC.getContentText());
    const TSCdataPoint = JSON.parse(TSCpoint.getContentText())


    const modifiedStockPoint = TSCdataPoint.data
        .map(v => {
            return {
                time: "" + v[0],
                point: v[1],
            }
        });
    
    console.log(modifiedStockPoint)
}

function point2(){
    var OTCpoint = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/iNdex_info/minute_index/1MIN_result.php?l=zh-tw&_=1636538447144`);
    
    const OTCdataPoint = JSON.parse(OTCpoint.getContentText())


    const modifiedStockPoint2 = OTCdataPoint.aaData
        .map(v => {
            return {
                time: "" + v[0],
                point: v[21],
            }
        });
    
    console.log(modifiedStockPoint2)
}



function getCSV(){
    let OTCpoint = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/iNdex_info/minute_index/1MIN_download.php?l=zh-tw&d=110/11/10&s=0,asc,0`);
    //Logger.log(responseTSC.getContentText());
    //const OTCdataPoint = JSON.parse(OTCpoint.getContentText())
    let csvTest = OTCpoint.getContentText("BIG5");
    let csv = Utilities.parseCsv(csvTest)
    console.log(csv)
}

function getHTML(){
    //NOTE 取得HTML整個網頁的資料 用正規表達式儲存資料

    let OTCpoint = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/iNdex_info/minute_index/1MIN_print.php?l=zh-tw&d=110/11/10&l&s=0,asc,0`);
    let Test = OTCpoint.getContentText();
    console.log(Test)
}


function test(){
    const myDate = new Date()
    const year = myDate.getFullYear().toString();
    const month = (myDate.getMonth() > 9) ? (myDate.getMonth() + 1): "0" + (myDate.getMonth() + 1);
    const date = (myDate.getDate() > 9) ? myDate.getDate() + 0: "0" + (myDate.getDate() - 0);
    const time = year + month + date;
    const time2 = 20210124;
    console.log(time);
    //取得上市櫃資料

    var responseTSC = UrlFetchApp.fetch(`https://www.twse.com.tw/exchangeReport/MI_INDEX?response=JSON&date=${time}&type=ALLBUT0999`);
    //Logger.log(responseTSC.getContentText());
    const data = JSON.parse(responseTSC.getContentText())


    const modifiedStockTSC = data.data9
        .filter(([code]) => code.length === 4)
        .map(v => {
            return {
                stockId: "" + v[0],
                stockName: v[1],
                volume: Math.round(+v[2].replace(/,/g,"") / 1000),
                openPrice: v[5],
                high: v[6],
                low: v[7],
                closePrice: v[8],
                //changeRange: testformat(v[9],v[10]),
                changeRange: v[9].match(/[+-]/g) === null ? v[10] : v[9].match(/[+-]/g) + v[10],
                date: data.date,
            }
        });

    function testformat(x,y){
        if(x.match(/[+-]/g) === null){
            return y
        }
        else return x.match(/[+-]/g) + y
    }

     console.log(modifiedStockTSC);

    var responseOTC = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&o=json&d=110/${month}/${date}&s=0,asc,0`);
    //Logger.log(responseOTC.getContentText());
    const data2 = JSON.parse(responseOTC.getContentText())

    const modifiedStockOTC = data2.aaData
    .filter(([code]) => code.length === 4)
    .map(v => {
        return {
            stockId: "" + v[0],
            stockName: v[1],
            volume: Math.round(+v[8].replace(/,/g,"") / 1000),
            openPrice: v[4],
            high: v[5],
            low: v[6],
            closePrice: v[2],
            changeRange: v[3],
            date: data.date,
        }
    });
    console.log(modifiedStockOTC);
    const modifiedStockALL = modifiedStockTSC.concat(modifiedStockOTC);
    // console.log(modifiedStockALL);
    // const data = JSON.parse(response.getContentText()) as {
    //     stat:string
    //     data9:string[][];
    // };


    if(data.stat.includes("重新查詢")){
        console.log("查詢的日期沒有資料")
        return;
    } 
    //console.log(data);

    //建立資料表到stockData
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let yourNewSheet = ss.getSheetByName("stock" + time);
    //if (yourNewSheet) return
    if (yourNewSheet === null) {

        createsheet(time);

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("stock" + time);
        //ss.insertSheet("123", 1 );
        //console.log(sheet);

        Logger.log(data.fields7[0]);

        //如果沒資料應該被中斷
        //填入資料到資料庫
        modifiedStockALL.forEach((v,i) => {
            //TODO format 把stockid轉成string (text)
            sheet.getRange(`A${i + 2}`).setValue(v.date); 
            sheet.getRange(`B${i + 2}`).setNumberFormat("@").setValue(v.stockId); 
            sheet.getRange(`C${i + 2}`).setValue(v.stockName);
            sheet.getRange(`D${i + 2}`).setValue("0");  
            sheet.getRange(`E${i + 2}`).setValue(v.openPrice); 
            sheet.getRange(`F${i + 2}`).setValue(v.closePrice); 
            sheet.getRange(`G${i + 2}`).setValue(v.high); 
            sheet.getRange(`H${i + 2}`).setValue(v.low); 
            sheet.getRange(`I${i + 2}`).setValue(v.changeRange); 
            sheet.getRange(`J${i + 2}`).setValue(v.volume); 
            sheet.getRange(`K${i + 2}`).setValue("0"); 
            sheet.getRange(`L${i + 2}`).setValue("0"); 
        });

    }
}

function buysell(){
    const myDate = new Date()
    const year = myDate.getFullYear().toString();
    const month = (myDate.getMonth() > 9) ? (myDate.getMonth() + 1): "0" + (myDate.getMonth() + 1);
    const date = (myDate.getDate() > 9) ? myDate.getDate() - 0: "0" + (myDate.getDate() - 0);
    const time = year + month + date;
    //取得上市櫃買賣超
    console.log(time);
    //上市
    let buysellTSC = UrlFetchApp.fetch(`https://www.twse.com.tw/fund/T86?response=JSON&date=${time}&selectType=ALLBUT0999`);
    //console.log(month,date)
    const data = JSON.parse(buysellTSC.getContentText());

    const modifiedBuysellTSC = data.data
        .filter(([code]) => code.length === 4)
        .map(v => {
            return {
                stockId: v[0],
                stockName: v[1].replace(/ /g,""),
                foreignInvestors: Math.round(+v[4].replace(/,/g,"") / 1000),
                investmentTrust: Math.round(+v[10].replace(/,/g,"") / 1000),
                date: data.date,
            }
        });

    //console.log(modifiedBuysellTSC);

    let buysellOTC = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&se=EW&t=D&d=110/${month}/${date}&s=0,asc`);
    //上櫃
    const data2 = JSON.parse(buysellOTC.getContentText());
    
    const modifiedBuysellOTC = data2.aaData
        .filter(([code]) => code.length === 4)
        .map(v => {
            return {
                stockId: v[0],
                stockName: v[1].replace(/ /g,""),
                foreignInvestors: Math.round(+v[10].replace(/,/g,"") / 1000),
                investmentTrust: Math.round(+v[13].replace(/,/g,"") / 1000),
                date: data.date,
            }
        });

    //console.log(modifiedBuysellOTC);
    const buySellAll = modifiedBuysellTSC.concat(modifiedBuysellOTC);
        //建立資料表到buysell


        let stockDataSheet:any[][];
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("stock" + time); 

        if(sheet !== null){

       
        stockDataSheet = sheet.getRange("A1:L1").getValues();
        
        const keys = stockDataSheet[0];
        stockDataSheet = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
        console.log(stockDataSheet);
        console.log(buySellAll);
 
        let bsDataResult = stockDataSheet.map(v =>{
            return keys.reduce((pre, k, i) =>{
                pre[k] = v[i];
                return pre;
            },{})
        })

        console.log(bsDataResult);

        function mergeArrays(arrays, prop) {
            const merged = {};
        
            arrays.forEach(arr => {
                arr.forEach(item => {
                    merged[item[prop]] = Object.assign({}, merged[item[prop]], item);
                });
            });
            
            //console.log(Object.values(merged).map(v => Object.values(v)));
            //console.log(Object.values(merged));
            //console.log(merged);
            return Object.values(merged);
        }

        console.log(mergeArrays([bsDataResult, buySellAll], 'stockId'));

        let result = mergeArrays([bsDataResult, buySellAll], 'stockId') as any;
        console.log(result);

        result.forEach((v,i) => {
            //TODO format 把stockid轉成string (text)
            sheet.getRange(`A${i + 2}`).setValue(v.date); 
            sheet.getRange(`B${i + 2}`).setNumberFormat("@").setValue(v.stockId); 
            sheet.getRange(`C${i + 2}`).setValue(v.stockName);
            sheet.getRange(`D${i + 2}`).setValue("0");  
            sheet.getRange(`E${i + 2}`).setValue(v.openPrice); 
            sheet.getRange(`F${i + 2}`).setValue(v.closePrice); 
            sheet.getRange(`G${i + 2}`).setValue(v.high); 
            sheet.getRange(`H${i + 2}`).setValue(v.low); 
            sheet.getRange(`I${i + 2}`).setValue(v.changeRange); 
            sheet.getRange(`J${i + 2}`).setValue(v.volume); 
            sheet.getRange(`K${i + 2}`).setValue(v.foreignInvestors); 
            sheet.getRange(`L${i + 2}`).setValue(v.investmentTrust); 
        });

        // const ss = SpreadsheetApp.getActiveSpreadsheet();
        // const sheet = ss.getSheetByName("buysell");
       
        // buySellAll.forEach((v,i) => {
    
        //     sheet.getRange(`A${i + 2}`).setValue(v.date); 
        //     sheet.getRange(`B${i + 2}`).setValue(v.stockId); 
        //     sheet.getRange(`C${i + 2}`).setValue(v.stockName);
        //     sheet.getRange(`D${i + 2}`).setValue(v.foreignInvestors);  
        //     sheet.getRange(`E${i + 2}`).setValue(v.investmentTrust); 

        // });
        }
    
}



function doPost(e) {
    console.log("test doPost")
    var param = e.parameter;
    var name = param.name;
    var age = param.age;
  
    var replyMsg = '你的名字是：' + name + '，年紀：' + age + '歲。';

    return ContentService.createTextOutput(replyMsg);
  
  }
    function testArray(){

        
        let data:any[][];
        let bsData:any[][];
        let result;
        let title;
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("stockData");
        const bsSheet = ss.getSheetByName("buysell");

        bsData = bsSheet.getRange("A1:E1").getValues();
        data = sheet.getRange("A1:L1").getValues();
        
        const keys = data[0];
        const bskeys = bsData[0];
        data = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
        bsData = bsSheet.getRange(2,1,bsSheet.getLastRow(),bsSheet.getLastColumn()).getValues();


        let dataResult = data.map(v =>{
            return keys.reduce((pre, k, i) =>{

                pre[k] = v[i];
                return pre;
            },{})
        })
        let bsDataResult = bsData.map(v =>{
            return bskeys.reduce((pre, k, i) =>{
                pre[k] = v[i];
                return pre;
            },{})
        })

        function mergeArrays(arrays, prop) {
            const merged = {};
        
            arrays.forEach(arr => {
                arr.forEach(item => {
                    merged[item[prop]] = Object.assign({}, merged[item[prop]], item);
                });
            });
            
            //console.log(Object.values(merged).map(v => Object.values(v)));
            //console.log(Object.values(merged));
            //console.log(merged);
            return Object.values(merged);
        }
        //NOTE 測試已接通
        
        console.log(mergeArrays([dataResult, bsDataResult], 'stockId'));

    }

    function doGet({parameter}:DoGet){
    //console.log(e);
    const { func } = parameter; 
    //let data: GoogleAppsScript.Spreadsheet.Range[][]
    // NOTE 時間不能跑
    const myDate = new Date()
    const year = myDate.getFullYear().toString();
    const month = (myDate.getMonth() > 8) ? (myDate.getMonth() + 1): "0" + (myDate.getMonth() + 1);
    const date = (myDate.getDate() > 9) ? myDate.getDate() - 0: "0" + (myDate.getDate() - 0);
    const time = year + month + date;
    console.log(time);

    let data:any[][];
    let bsData:any[][];
    // for (const v in {func}) {
    //     if(v!=="func")
    //         {
    //             console.error("不存在的參數!!!");
    //             return ContentService.createTextOutput("不存在的參數!!!").setMimeType(ContentService.MimeType.TEXT);
    //         }
    // }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("stock" + time);
    const bsSheet = ss.getSheetByName("buysell");
    let result;
    let title;
    switch(func){
        case Func.All:
            //丟資料;
            

            title = sheet.getRange("A1:L1").getNotes();

            data = sheet.getRange("A1:L1").getValues();
            bsData = bsSheet.getRange("A1:E1").getValues();
            const keys = data[0];
            const bskeys = bsData[0];
            // const yesterdayCloseIndex = keys.indexOf("yesterdayClose");
            // const ClosingPriceIndex = keys.indexOf("ClosingPrice");

            //data = sheet.getRange("A2:L10").getValues();
            data = sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
            bsData = bsSheet.getRange(2,1,bsSheet.getLastRow(),bsSheet.getLastColumn()).getValues();
          
            let bsDataResult = bsData.map(v =>{
                return bskeys.reduce((pre, k, i) =>{
                    pre[k] = v[i];
                    return pre;
                },{})
            })
    
            let dataResult = data.map(v =>{
                return keys.reduce((pre, k, i) =>{
                    // 不用計算了
                    // if(k==="changeRange"){
                    //     //pre[k] = (v[ClosingPriceIndex]- v[yesterdayCloseIndex]).toFixed(2) + "(" + Math.abs(((1 - (v[ClosingPriceIndex] / v[yesterdayCloseIndex])) * 100)).toFixed(2) + "%)"; 
                    //     //改證明欄位在哪個位子
                    //     //pre[k] = pre[i] + 1; 
                    //     //return v[i];
                    //     pre[k] = `${(v[ClosingPriceIndex] - v[yesterdayCloseIndex]).toFixed(2)} ( ${Math.abs(((1 - (v[ClosingPriceIndex] / v[yesterdayCloseIndex])) * 100)).toFixed(2)} %)`;
                    // }

                    // else{
                    //     pre[k] = v[i];
                    // }
                    

                    pre[k] = v[i];


                    return pre;
                },{})
            })

            // 合併物件，以id比對 合併外資買賣超
            function mergeArrays(arrays, prop) {
                const merged = {};
            
                arrays.forEach(arr => {
                    arr.forEach(item => {
                        merged[item[prop]] = Object.assign({}, merged[item[prop]], item);
                    });
                });
            
                return Object.values(merged);
            }

            //const mergedArray = {...dataResult, ...bsDataResult};
            //const uniqueData = {...mergedArray.reduce((map, obj) => map.set(obj.stockId, obj), new Map()).values()};
            //let obj = Object.assign(bsDataResult, dataResult);

            //result = mergeArrays([dataResult, bsDataResult], 'stockId');
            result = dataResult;
           //note doGET 無法console取得資料查看
            
            break;

        case Func.chart:
            result = "This is stockId.";
            //data = sheet.getRange("B2:B6").getValues();
            https://www.twse.com.tw/exchangeReport/MI_5MINS_INDEX?response=json&date=${time}
            var TSCpoint = UrlFetchApp.fetch(`https://www.twse.com.tw/exchangeReport/MI_5MINS_INDEX?response=json&date=20211020`);
            //Logger.log(responseTSC.getContentText());
            const TSCdataPoint = JSON.parse(TSCpoint.getContentText())
        
        
            const modifiedStockPoint = TSCdataPoint.data
                .map(v => {
                    return {
                        time: "" + v[0],
                        point: v[1].replace(/,/g,""),
                    }
                });

            //result = data.toString();
            return ContentService.createTextOutput(JSON.stringify(modifiedStockPoint)).setMimeType(ContentService.MimeType.JSON)
            break;

        case Func.OTCchart:
          
            var OTCpoint = UrlFetchApp.fetch(`https://www.tpex.org.tw/web/stock/iNdex_info/minute_index/1MIN_result.php?l=zh-tw&_=1636538447144`);
    
            const OTCdataPoint = JSON.parse(OTCpoint.getContentText())
        
        
            const modifiedStockPoint2 = OTCdataPoint.aaData
                .map(v => {
                    return {
                        time: "" + v[0],
                        point: v[21],
                    }
                });

            return ContentService.createTextOutput(JSON.stringify(modifiedStockPoint2)).setMimeType(ContentService.MimeType.JSON)

        case Func.time:
            const myDate = new Date()
            const year = myDate.getFullYear().toString();
            const month = (myDate.getMonth() > 8) ? (myDate.getMonth() + 1): "0" + (myDate.getMonth() + 1);
            const date = (myDate.getDate() > 9) ? myDate.getDate() - 0: "0" + (myDate.getDate() - 0);
            const hour = (myDate.getHours());
            const time = year + month + date + " " + hour;
            return ContentService.createTextOutput(JSON.stringify({time})).setMimeType(ContentService.MimeType.JSON);
            break;
        case Func.stockId:
            //result = "This is stockId.";
            data = sheet.getRange("B2:B6").getValues();
            result = data.toString();
            break;
        case Func.stockName:
            //result = "This is stockName";
            data = sheet.getRange("C2:C6").getValues();
            result = data.toString();
            break;
        case Func.OpenPrice:
            //result = "This is OpenPrice.";
            data = sheet.getRange("E2:E6").getValues();
            result = data.toString();
            break;
        case Func.ClosingPrice:
            //result = "This is ClosingPrice";
            data = sheet.getRange("F2:F6").getValues();
            result = data.toString();
            break;
        case Func.volume:
            //result = "This is volume.";
            data = sheet.getRange("J2:J6").getValues();
            result = data.toString();
            break;    
                        
        default:
            console.error("這不存在的參數");
            //return ContentService.createTextOutput("這不存在的參數").setMimeType(ContentService.MimeType.TEXT);
           
    }
    //throw new Error("多出不必要的參數內容")
    
    //const {user, e} = parameter;
    //SpreadsheetApp.flush();
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON)

    //可以丟兩個不同資料嗎
}

