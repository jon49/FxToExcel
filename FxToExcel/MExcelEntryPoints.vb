Imports Excel = NetOffice.ExcelApi
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI

<ComVisible(True)> _
Public Class CExcelEntryPoints

    Inherits ExcelRibbon

    Public Sub ImportTransactions(ByVal control As IRibbonControl)
        'Dim xlApp = New Excel.Application(Nothing, control.Context.Application)
        UpdateTransactions(True)
    End Sub

End Class

Public Module MExcelEntryPoints

    Public Sub ImportTransactions()
        'Dim xlApp = New Excel.Application(Nothing, control.Context.Application)
        UpdateTransactions(True)
    End Sub
    
End Module
