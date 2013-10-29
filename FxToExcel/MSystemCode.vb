Imports System.IO
Imports ExcelDna.Integration.CustomUI
Imports NetOffice.ExcelApi
Imports ExcelExtensions
Imports UpdateXL

Friend Module MSystemCode

    ''' <summary>
    ''' Gets the financial Excel workbook location and creates shortcut if desired.
    ''' </summary>
    ''' <param name="sFileFullName">Full name of workbook.</param>
    ''' <returns>File name and location.</returns>
    Public Function GetUserWkbLocationFromUser(ByVal sFileFullName As String) As String

        Dim sFileName = My.Computer.FileSystem.GetName(sFileFullName)

        Dim b2007AndAbove As Boolean = False

        Try
            'Make sure form was filled correctly.
            If sFileName.Length > 0 Then
                'Determine is extension is on file.
                If CInt(gXLApp.Version) > 11 Then b2007AndAbove = True
                If Not sFileName.ToUpper.Contains(".XLS") Then
                    If b2007AndAbove Then
                        sFileFullName = sFileFullName & ".xlsx"
                    Else
                        sFileFullName = sFileFullName & ".xls"
                    End If
                End If
                'If file doesn't exist then create workbook from template
                If Not My.Computer.FileSystem.FileExists(sFileFullName) Then
                    gXLApp.DisplayAlerts = False
                    gwkbMain = gXLApp.CreateWorkbook(sFileFullName)
                End If
                Return sFileFullName
            End If
        Catch
            MsgBox("Error while installing new excel workbook!", MsgBoxStyle.Critical)
        Finally
            gXLApp.DisplayAlerts = True
        End Try

        Return ""

    End Function

    ''' <summary>
    ''' Update the add-in to the current newest files.
    ''' </summary>
    ''' <remarks>Jon Nyman 2013-10-21</remarks>
    Friend Sub UpdateAddIn()

        If My.Computer.FileSystem.FileExists(CSettings.WyUpdateFullName) Then
            Dim hiddenFiles = getHiddenFileNames(CSettings.ApplicationPath)
            If Not hiddenFiles.Contains(".SyncID") Then
                'If there is an update available then start WyUpdate form.
                Dim p = New ProcessStartInfo(CSettings.WyUpdateFullName, "/quickcheck /noerr")
                p.WindowStyle = ProcessWindowStyle.Hidden
                Process.Start(p)
            End If
        End If

    End Sub

    ''' <summary>
    ''' Update financial transaction information to Excel.
    ''' </summary>
    ''' <param name="AlwaysUpdate">Update regardless if currently updated.</param>
    ''' <remarks>Jon Nyman 2013-10-21</remarks>
    Friend Sub UpdateTransactions(ByVal AlwaysUpdate As Boolean)

        Dim clsSettings = New CSettings

        Try
            Financial.Data(gwkbMain, clsSettings.FxFullName, AlwaysUpdate, clsSettings.ExcelLastUpdated, clsSettings.FxStartDate)
            clsSettings.ExcelLastUpdated = Now
            clsSettings.Save()
        Catch ex As Exception
            MsgBox("There was an error which occurred while updating. " & Environment.NewLine & ex.Message, MsgBoxStyle.Exclamation, "Update Error")
        Finally
            clsSettings = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' Return number of hidden files in folder
    ''' </summary>
    ''' <param name="path">Path of directory.</param>
    ''' <returns>List of hidden files.</returns>
    Private Function getHiddenFileNames(path As String) As List(Of String)

        Return _
            (New DirectoryInfo(path)).GetFiles() _
            .Where(Function(file) (FileAttributes.Hidden = file.Attributes)) _
            .Select(Function(f) f.Name).ToList

    End Function

End Module

