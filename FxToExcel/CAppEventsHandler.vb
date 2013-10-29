Option Explicit On

Imports ExcelDna.Integration
Imports NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports System.Linq

Public Class CAppEventsHandler

    Implements IExcelAddIn

    Private WithEvents XLApp As Application
    Private mbCloseNewWorkbook = True
    Private mbClosed As Boolean = False

    Public Sub Start() Implements IExcelAddIn.AutoOpen

        NetOffice.Factory.Initialize()

        XLApp = New Application(Nothing, ExcelDnaUtil.Application)

        Dim xlCalc = XlCalculation.xlCalculationAutomatic

        Try
            InitializeGlobals(XLApp, xlCalc)
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Open Error")
        Finally
            gXLApp.EnableEvents = True
            gXLApp.Calculation = xlCalc
            gXLApp.DisplayAlerts = True
        End Try

    End Sub

    Public Sub Close() Implements IExcelAddIn.AutoClose
        If Not IsNothing(gXLApp) Then
            If gXLApp.Workbooks.Count = 0 Then gXLApp.IgnoreRemoteRequests = False
        End If
    End Sub

    Private Sub XLAPP_WorkbookClose(ByVal Wb As Workbook, ByRef cancel As Boolean) Handles XLApp.WorkbookBeforeCloseEvent

        If Not IsNothing(gwkbMain) AndAlso gwkbMain.FullName = Wb.FullName Then
            gXLApp.IgnoreRemoteRequests = False
            UpdateAddIn()
            gXLApp.Dispose()
            gXLApp = Nothing
        End If

    End Sub

    Private Sub XLApp_WorkbookNew(ByVal Wb As Workbook) Handles XLApp.NewWorkbookEvent

        If Not IsNothing(gXLApp) AndAlso Not IsNothing(gwkbMain) Then
            If IsNothing(Wb.Path) OrElse Wb.Path.Length = 0 Then
                Wb.Close()
                Wb = Nothing
            End If
        End If

    End Sub


End Class
