Imports NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Module MGlobals

    Public gbDebug = False

    Public gXLApp As Application
    Public gwkbMain As Workbook
    Public gxlCalc As XlCalculation

    Public Function InitializeGlobals(ByVal xlApp As Application, ByRef xlCalc As XlCalculation) As Boolean

        gXLApp = xlApp
        gXLApp.DisplayAlerts = False
        gXLApp.IgnoreRemoteRequests = True

        If InitializeWorkbook() Then
            gxlCalc = xlCalc
            gwkbMain = gXLApp.ActiveWorkbook
            gXLApp.Visible = True
            UpdateTransactions(False)
            Return Not IsNothing(gwkbMain)
        End If

        Return False

    End Function

    ''' <summary>
    ''' Open the associated workbook.
    ''' </summary>
    Private Function InitializeWorkbook() As Boolean

        Dim clsSettings As New CSettings

        Try

            Dim sFileFullName = clsSettings.WorkbookLocation
            If Not My.Computer.FileSystem.FileExists(sFileFullName) Then
                'See if this is the first time that this file will be located in this directory.  If so, create new workbook.
                Dim swkbLocation = GetUserWkbLocationFromUser(sFileFullName)
                If swkbLocation.Length = 0 Then Return False
                clsSettings.WorkbookLocation = swkbLocation
                clsSettings.Save()
            End If
            'Open workbook if it isn't opened already.
            If IsNothing(gXLApp.Workbooks.Where(Function(wkb) wkb.FullName = clsSettings.WorkbookLocation).FirstOrDefault) Then
                gXLApp.Workbooks.Open(clsSettings.WorkbookLocation, True, False, Type.Missing, Type.Missing, Type.Missing, Type.Missing _
                                      , Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
            End If

            Return True

        Catch ex As Exception
            My.Application.Log.WriteException(ex)
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Failed to open workbook.")
        Finally
            clsSettings = Nothing
        End Try

        Return False

    End Function

End Module
