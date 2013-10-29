Friend Module Utilities

    ''' <summary>
    ''' Get's file name from the user
    ''' </summary>
    ''' <param name="Title">Title on dialogue box.</param>
    ''' <param name="Filter">Filter used in dialogue box.
    ''' E.g., 
    ''' Excel Workbook (*.xls*)|*.xls*
    ''' txt files (*.txt)|*.txt|All files (*.*)|*.*</param>
    ''' <param name="InitialDirectory">Load directory.
    ''' E.g.,
    ''' My.Computer.FileSystem.SpecialDirectories.MyDocuments.ToString</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function GetFileLocationFromUser(ByVal Title As String, ByVal Filter As String _
                                            , ByVal CheckFileExists As Boolean, ByVal CheckPathExists As Boolean _
                                            , ByVal InitialDirectory As String) As String

        Using FileDialog1 As New System.Windows.Forms.OpenFileDialog
            FileDialog1.Title = Title
            FileDialog1.InitialDirectory = InitialDirectory
            FileDialog1.Filter = Filter
            FileDialog1.CheckFileExists = CheckFileExists
            FileDialog1.CheckPathExists = CheckPathExists
            If FileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                Return FileDialog1.FileName
            Else
                Return ""
            End If
        End Using

    End Function

End Module
