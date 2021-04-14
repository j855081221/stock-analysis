enum Func {
    stockId = "stockId", 
    stockName = "stockName",
    OpenPrice = "OpenPrice",
    ClosingPrice = "ClosingPrice",
    volume = "volume"

}
//console.log("hello world");
type DoGet = GoogleAppsScript.Events.DoGet & {
    parameter:{
       func:Func;
    }
}

function doGet({parameter}:DoGet){
    //console.log(e);
    const { func } = parameter; 
    //let data: GoogleAppsScript.Spreadsheet.Range[][]
    let data
    // for (const v in {func}) {
    //     if(v!=="func")
    //         {
    //             console.error("不存在的參數!!!");
    //             return ContentService.createTextOutput("不存在的參數!!!").setMimeType(ContentService.MimeType.TEXT);
    //         }
    // }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("stockData");
    let result: string;

    
    switch(func){
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
            console.error("不存在的參數");
            return ContentService.createTextOutput("不存在的參數").setMimeType(ContentService.MimeType.TEXT);
    }
    //throw new Error("多出不必要的參數內容")
    
    //const {user, e} = parameter;
    //SpreadsheetApp.flush();
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON)
}