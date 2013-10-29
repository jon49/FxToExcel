namespace FXCore

open System.IO

module internal Read' =

    let gnuExtension = """.gnucash"""
    let kmyExtension = """.kmy"""

    type SourceType = 
        | KMyMoney of string
        | GNUCash of string
        | ErrorMessage of string

module Read =
    
    type FinancialFile(sourceFile : string) = 
        member internal this.FileLocation = sourceFile
        member internal this.Type =
            let fileExtension = (new FileInfo(sourceFile)).Extension.ToLower()
            match fileExtension with
            | f when f = Read'.kmyExtension -> Read'.SourceType.KMyMoney sourceFile
            | f when f = Read'.gnuExtension -> Read'.SourceType.GNUCash sourceFile
            | _ -> Read'.SourceType.ErrorMessage "Unrecognised file type!"
        member this.ImportData date =
            match this.Type with
            | Read'.GNUCash(gnuExtension) -> None
            | Read'.KMyMoney(kmyExtension) -> Some(KMyMoney.Read.ImportData (this.FileLocation) date)
            | _ -> None


