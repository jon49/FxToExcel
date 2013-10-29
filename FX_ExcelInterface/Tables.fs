namespace UpdateXL.Tables

open NetOffice.ExcelApi
open SpreadSharp
open ExcelExtensions
open UpdateXL

module internal TransactionTable =
    
    /// <summary> Export transaction data to Excel </summary>
    /// <param name="wkb"> Target workbook to update. </param>
    /// <param name="fxData"> Data to update workbook with. </param>
    let create (wkb : Workbook) (fxData:FXCore.Structure.Transactions) =
        let wks = wkb.CreateWorksheet(XLNames.Worksheets.Data)
        let headerNames = fxData.Transactions |> Records.fieldsNames |> Array.map (fun o -> string o)
        //Export transaction data to Excel (need to do headers and data separate to avoid erasing table name and range names)
        let rTransactions = wks.Range("A1") |> Utils.tableToRange XLNames.TableNames.Transactions headerNames (fxData.Transactions |> Records.fieldsArray)
        //Create dynamic range names
        if (Utils.xlAppVersion < 12) then 
            Utils.createDynamicRangeNames (rTransactions.CurrentRegion) (XLNames.TableNames.Transactions + ".") headerNames // |> ignore
        //Format date column to default local settings "m/d/yy" see http://bytes.com/topic/access/answers/284725-date-formats-when-exporting-excel-via-vb
        rTransactions 
        |> XlRange.resize (rTransactions.Rows.Count - 1) 1 
        |> XlRange.offset 1 (Array.findIndex (fun header -> header = XLNames.HeaderNames.Date) headerNames)
        |> XlRange.numberFormat XLNames.Formulas.DateFormat
        
module internal CategoryTable =
    
    ///Export category data to Excel
    let create (wkb:Workbook) (fxData:FXCore.Structure.Transactions) =
        let wks = wkb.CreateWorksheet(XLNames.Worksheets.Categories)
        let headerNames = [| XLNames.HeaderNames.Category; XLNames.HeaderNames.Type |]
        // Get the Category and Type in 2D sorted array
        let categoryData = 
            fxData.Transactions 
            |> Seq.distinctBy (fun r -> r.Category) 
            |> Seq.sortBy (fun r -> r.Category_Type, r.Category)
            |> Seq.map (fun r -> [| box r.Category; box r.Category_Type |])
            |> array2D
        let rCategories = wks.Range("A1") |> Utils.tableToRange XLNames.TableNames.Categories headerNames categoryData
        Utils.createDynamicRangeNames rCategories (XLNames.TableNames.Categories + ".") headerNames |> ignore
        

module internal CategoryTypeTable =
    
    ///Export category type data to Excel
    let create (wkb:Workbook) (fxData:FXCore.Structure.Transactions) =
        let wks = wkb.CreateWorksheet(XLNames.Worksheets.CategoryTpes)
        let headerNames = [| XLNames.HeaderNames.CategoryType |]
        let categoryTypes =
            fxData.Transactions
            |> Seq.distinctBy (fun r -> r.Category_Type)
            |> Seq.sort
            |> Seq.map (fun r -> [| box r.Category_Type |])
            |> array2D
        let rCategoryTypes = wks.Range("A1") |> Utils.tableToRange XLNames.TableNames.CategoryTypes headerNames categoryTypes
        if (Utils.xlAppVersion < 12) then 
            Utils.createDynamicRangeNames rCategoryTypes (XLNames.TableNames.CategoryTypes + ".") headerNames |> ignore
        

