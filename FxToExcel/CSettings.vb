
Imports System.Configuration
Imports System.Reflection
Imports NetOffice.ExcelApi
Imports ExcelDna.Integration
Imports System.Linq

Class CSettings
    Inherits ApplicationSettingsBase

    Private Const mCsRegistryAddress As String = "HKEY_CURRENT_USER\Software\Spreadsheet Budget\AppSettings"
    Private Const mCsWkbLocationFromRegistry As String = "FXToExcelWorkbookLocation"
    Private Const mCsFxLocationFromRegistry = "FXToExcel_FxLocation"

#Region " Methods "

    Private Function GetFileName(ByVal sSettingName As String, ByVal sRegistryName As String, ByVal Title As String, ByVal Filter As String _
                                , ByVal CheckFileExists As Boolean, ByVal CheckPathExists As Boolean, ByVal InitialDirectory As String) As String

        Dim sFileFullName = CStr(Me(sSettingName))

        If sFileFullName.Length = 0 OrElse Not My.Computer.FileSystem.FileExists(sFileFullName) Then
            sFileFullName = CStr(My.Computer.Registry.GetValue(mCsRegistryAddress, sRegistryName, ""))
        End If

        If Not My.Computer.FileSystem.FileExists(sFileFullName) Then
            Dim sFileNewFile = GetFileLocationFromUser(Title, Filter, CheckFileExists, CheckPathExists, InitialDirectory)
            SetFileName(sFileNewFile, sSettingName, sRegistryName)
            Return sFileNewFile
        End If

        Return sFileFullName

    End Function

    Private Sub SetFileName(ByVal sFullFileName As String, ByVal sSettingName As String, ByVal sRegistryName As String)
        Me(sSettingName) = sFullFileName
        Try
            My.Computer.Registry.SetValue(mCsRegistryAddress, sRegistryName, sFullFileName)
        Catch ex As Exception
            MsgBox(ex)
        End Try
    End Sub

#End Region

#Region " User Settings "

    <UserScopedSettingAttribute(), DefaultSettingValue("1/1/1900")> _
    Property FxStartDate As Date
        Get
            Return CDate(Me("FxStartDate"))
        End Get
        Set(value As Date)
            Me("FxStartDate") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("")> _
    Public Property WorkbookLocation() As String
        Get
            Return GetFileName("WorkbookLocation", mCsWkbLocationFromRegistry _
                               , "Create Excel Workbook Name and Location" _
                               , "Excel Workbook (*.xls*)|*.xls*" _
                               , False, False _
                               , My.Computer.FileSystem.SpecialDirectories.MyDocuments.ToString)
        End Get
        Set(ByVal value As String)
            SetFileName(value, "WorkbookLocation", mCsWkbLocationFromRegistry)
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValue("")> _
    Property FxFullName As String
        Get
            Return GetFileName("FXFullName", mCsFxLocationFromRegistry _
                               , "Financial File Location" _
                               , "KMyMoney (*.kmy)|*.kmy" _
                               , True, True _
                               , My.Computer.FileSystem.SpecialDirectories.MyDocuments.ToString)
        End Get
        Set(value As String)
            SetFileName(value, "FXFullName", mCsFxLocationFromRegistry)
        End Set
    End Property


#End Region

#Region " App Settings "

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("1/1/1900")> _
    Public Property AppLastUpdated() As Date
        Get
            Return CDate(Me("AppLastUpdated"))
        End Get
        Set(value As Date)
            Me("AppLastUpdated") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("1/1/1900")> _
    Public Property ExcelLastUpdated() As Date
        Get
            Return CDate(Me("ExcelLastUpdated"))
        End Get
        Set(value As Date)
            Me("ExcelLastUpdated") = value
        End Set
    End Property

    Public Shared ReadOnly ApplicationPath = System.AppDomain.CurrentDomain.BaseDirectory

    Public Shared ReadOnly WyUpdateFullName = System.IO.Path.Combine(ApplicationPath, "wyUpdate.exe")

#End Region



End Class
