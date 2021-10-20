enum Func {
    Test = "test",
    Stock = "stock",
    Chart = "chart",
}

type Stock = {
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

type DoGet = GoogleAppsScript.Events.DoGet & {
    // parameter: {
    //     func: Func,
    //     time?: string,
    // },
    parameter: {
        func: Func.Test;
    } | {
        func: Func.Stock | Func.Chart;
        time?: string,
    },
}
