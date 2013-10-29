namespace XLNames

module Worksheets =
    
    let Data = "Transactions"
    let Categories = "Categories"
    let CategoryTpes = "CategoryTpes"

module TableNames =

    let Transactions = "T.Transactions"
    let Categories = "T.Categories"
    let CategoryTypes = "T.CategoryTypes"

module RangeNames =

    let Categories_Category = TableNames.Categories + ".Category"
    let Transaction_Category = TableNames.Transactions + ".Category"

module HeaderNames =

    let Date = "Date"
    let Category = "Category"
    let Type = "Type"
    let CategoryType = "Category Type"

module Formulas =

    let DateFormat = "m/d/yy"

//module Columns =
//
//    let TransactionColumnLetter = "A"
//    let CategoryColumnLetter = "A"
//    let TypeColumnLetter = "A"
//    let CategoryTypeColumnLetter = "A"
