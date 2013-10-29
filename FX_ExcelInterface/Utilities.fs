namespace UpdateXL

open System.IO
open NetOffice.ExcelApi
open SpreadSharp
open ExcelExtensions

module internal Utils =

    /// <summary>Determine if the last time updated is newer or older than updated than financial file. </summary>
    /// <param name="lastUpdated">Last time Excel was updated with financial information.</param>
    /// <param name="sourceFinancialFile">Full name of financial source file.</param>
    /// <returns>True: When financial source file is newer than last time Excel was updated. 
    /// False: Otherwise</returns>
    let shouldUpdate lastUpdated sourceFinancialFile =
        let file = new FileInfo(sourceFinancialFile)
        if file.Exists then 
            let lastWriteTime = file.LastWriteTime
            (lastWriteTime > lastUpdated)
        else
            false

    ///Create a dynamic single column named range
    let createRangeName rangeName (range : Range) = range.CreateDynamicRangeName(rangeName, 1, 1)
        
    ///Create multiple dynamic single column named ranges
    let createDynamicRangeNames (range : Range) tablePrefix (headerNames : string []) =
        let columnCount = range.Columns.Count
        let offsetRange offset = range |> XlRange.resize 1 1 |> XlRange.offset 1 offset
        let ranges = [|0 .. (columnCount - 1)|] |> Array.map offsetRange
        Array.map2 
            (fun n r -> createRangeName (tablePrefix + n) r) headerNames ranges
            |> ignore

    ///copy over table to range by headers then by data as to not overwrite current table
    let tableToRange tableName (headers : string[]) data (range : Range) =
        let headers2D = Array2D.init 1 headers.Length (fun i j -> headers.[j])
        let range' = 
            range 
            |> XlRange.resize 1 1 |> XlRange.setValue headers2D
            |> XlRange.offset 1 0 |> XlRange.setValue data
            |> XlRange.currentRegion
        range'.CreateTable(tableName) |> ignore
        range'

    /// Get the version number of Excel
    let xlAppVersion = (XlApp.getActiveApp()).Version |> float |> int
