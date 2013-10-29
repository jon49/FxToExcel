// Learn more about F# at http://fsharp.net

namespace KMyMoney

open System.IO
open FSharp.Data
open FXCore.Utilities
open FXCore.Structure

    module internal GetData =

        type KMyMoney = XmlProvider<"KMyMoneySampleFile.xml"> 
        
        let kMyMoney (sourceData:string) = KMyMoney.Parse(sourceData)

        let getAccountNames sourceData = 
            (kMyMoney sourceData).Accounts.GetAccounts()
            |> Array.map (fun acc -> acc.Id, acc.Name)
            |> Map.ofArray

        let getPayeeNames sourceData =
            (kMyMoney sourceData).Payees.GetPayees()
            |> Array.map (fun p -> p.Id, p.Name)
            |> Map.ofArray

        let getAccountTypes sourceData =
            let typeName (account:KMyMoney.DomainTypes.Account) =
                if account.Parentaccount.Length > 0 then (account.Parentaccount.Substring(account.Parentaccount.IndexOf("::") + 2)) else account.Name
            (kMyMoney sourceData).Accounts.GetAccounts()
            |> Array.filter (fun a -> Option.isSome  a.Subaccounts)
            |> Array.map (fun a -> (typeName a), (Option.get (a.Subaccounts)).GetSubaccounts())
            |> Array.map (fun x -> 
                let name = fst x
                let accounts = snd x
                accounts |> Array.map (fun y -> y.Id, name))
            |> Array.concat |> Map.ofArray
                
        let splitTransaction (accNames:Map<string,string>) (accTypes:Map<string, string>) (payeeNames:Map<string,string>) (transaction:KMyMoney.DomainTypes.Transaction) =
            let split = transaction.Splits.GetSplits()
            {Id = transaction.Id; Commodity = transaction.Commodity; Date = transaction.Postdate
            ; Account = accNames.TryFind(split.[0].Account) |? "Unknown"; Amount = amount split.[0].Value
            ; Payee = payeeNames.TryFind(split.[1].Payee) |? "Unknown"; Category = accNames.TryFind(split.[1].Account) |? "Unknown"
            ; Number = split.[0].Number; Memo = split.[0].Memo; Shares = amount split.[0].Shares
            ; Category_Type = accTypes.TryFind(split.[1].Account) |? "Unknown"}
        
        let groupTransactions split (dte:System.DateTime) toCombineTransactions =
            let combinedTransactions = 
                let amounts = [ (fun (x:Transaction) -> x.Amount) ]
                toCombineTransactions
                |> Array.map split
                |> Seq.groupBy (fun trnx -> trnx.Account, trnx.Category) 
                |> Seq.map (fun (key, values) ->
                    (key, List.sum [for a in amounts -> values |> Seq.sumBy a]))
            let blankConsolidatedTransaction account category amt = 
                {Id = ""; Commodity = ""; Date = dte.AddDays(-1.); Account = account; Amount = amt
                ; Payee = ""; Category = category; Number = ""; Memo = ""
                ; Shares = 0.; Category_Type = ""}
            combinedTransactions
            |> Seq.map (fun ((acc, cat), total) ->
                blankConsolidatedTransaction acc cat total)
            |> Seq.toArray

        let getTransactions sourceData dte = 
            let split = splitTransaction (getAccountNames sourceData) (getAccountTypes sourceData) (getPayeeNames sourceData)
            let kData = kMyMoney sourceData
            let listAbove, toCombineTransactions = 
                Array.partition (fun (t:KMyMoney.DomainTypes.Transaction) -> t.Postdate >= dte)  (kData.Transactions.GetTransactions())
            let listAbove' = listAbove |> Array.map split
            Array.append (groupTransactions split dte toCombineTransactions) listAbove'
        
        let importData sourceFile date =
            let sourceData = decompressFileAndRead sourceFile
            let transactions = (getTransactions sourceData date) |> List.ofArray
            {Transactions = transactions}

    module internal Read =
        
        open GetData
        open FXCore.Structure
        
        ///Import data from KMyMoney XML file as 2D object array.
        let ImportData sourceFile date =
            importData sourceFile date