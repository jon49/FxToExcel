namespace UpdateXL

open FXCore
open NetOffice.ExcelApi
open ExcelExtensions
open UpdateXL

module Financial =
    
    let internal updateTables (xlApp:NetOffice.ExcelApi.Application) wkb fxData =
        Tables.TransactionTable.create wkb fxData
        Tables.CategoryTable.create wkb fxData
        Tables.CategoryTypeTable.create wkb fxData
        wkb.RefreshPivotTables()
        xlApp.StatusBar <- printf "Refreshed XML data and pivot tables - %A." System.DateTime.Now
        
    /// <summary> Update Excel data from financial application file. </summary>
    /// <param name="wkb"> Workbook to import financial data. </param>
    /// <param name="sourceFinancialFile"> Financial data file name (full). </param>
    /// <param name="alwaysUpdate"> When true will always perform update regardless of last updated information. </param>
    /// <param name="lastUpdated"> Last time Excel workbook was updated. </param>
    /// <param name="dateToImport"> Import only values above this date. </param>
    let Data (wkb : Workbook) sourceFinancialFile alwaysUpdate lastUpdated dateToImport =
        let update = Utils.shouldUpdate lastUpdated sourceFinancialFile
        if update || alwaysUpdate then
            let fxData = Read.FinancialFile(sourceFinancialFile).ImportData dateToImport
            let xlApp = (wkb.Parent) :?> NetOffice.ExcelApi.Application
            match fxData with
            | Some(fxData) ->
                updateTables xlApp wkb fxData |> ignore
            | None -> 
                xlApp.StatusBar <- printf "Update did not occur. No data returned from financial file."
                None |> ignore
        
            
            
            
        

