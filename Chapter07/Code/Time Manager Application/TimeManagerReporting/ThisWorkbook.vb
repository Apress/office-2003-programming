Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports MSForms = Microsoft.Vbe.Interop.Forms

' Office integration attribute. Identifies the startup class for the workbook. Do not modify.
<Assembly: System.ComponentModel.DescriptionAttribute("OfficeStartupClass, Version=1.0, Class=TimeManagerReporting.OfficeCodeBehind")>

Public Class OfficeCodeBehind 
#Region "Declarations"
  Friend WithEvents ThisWorkbook As Excel.Workbook
  Friend WithEvents ThisApplication As Excel.Application
  Friend WithEvents TimeManagerSettingsMenuButton As Office.CommandBarButton

  Private WithEvents btnGetEmpReport As MSForms.CommandButton
  Private comboEmployeeNames As MSForms.ComboBox
  Private txtDateRangeStart As MSForms.TextBox
  Private txtDateRangeEnd As MSForms.TextBox

  Private TimeManager As New TMW.TimeManagerWeb
  Private UserArray() As TMW.tmUser

  Private CurrentUser As TMW.tmUser
#End Region


#Region "Generated initialization code"

  ' Default constructor.
  Public Sub New()
  End Sub

  ' Required procedure. Do not modify.
  Public Sub _Startup(ByVal application As Object, ByVal workbook As Object)
    ThisApplication = CType(application, Excel.Application)
    ThisWorkbook = CType(workbook, Excel.Workbook)


  End Sub

  ' Required procedure. Do not modify.
  Public Sub _Shutdown()
    ThisApplication = Nothing
    ThisWorkbook = Nothing
  End Sub

  ' Returns the control with the specified name on ThisWorkbook's active worksheet.
  Overloads Function FindControl(ByVal name As String) As Object
    Return FindControl(name, CType(ThisWorkbook.ActiveSheet, Excel.Worksheet))
  End Function

  ' Returns the control with the specified name on the specified worksheet.
  Overloads Function FindControl(ByVal name As String, ByVal sheet As Excel.Worksheet) As Object
    Dim theObject As Excel.OLEObject
    Try
      theObject = CType(sheet.OLEObjects(name), Excel.OLEObject)
      Return theObject.Object
    Catch Ex As Exception
      ' Returns Nothing if the control is not found.
    End Try
    Return Nothing
  End Function
#End Region


#Region "Custom Methods"
  Public Sub DoLogin()
    '-----------------------------------------------------------------------
    '   Displays a login form allowing the user to enter a username and
    '   password.  If the user clicks the cancel button, the workbook will
    '   close.  If the user clicks the login button, the information 
    '   provided will be checked against the database using the web service
    '   If the login attempt fails, the login will be displayed again.  If
    '   it succeeds, the subroutine will exit and the CurrentUser variable 
    '   will contain a reference to the current user.
    '-----------------------------------------------------------------------

    Dim loginForm As New TimeManagerLogin

    loginForm.ShowDialog()

    If loginForm.Cancelled = True Then
      ThisWorkbook.Close()
    Else
      CurrentUser = TimeManager.Login(loginForm.userId, loginForm.password)
      If CurrentUser Is Nothing Then
        MsgBox("The credential you supplied were invalid.  Please try again")
        loginForm = Nothing
        DoLogin()
      Else
        If Not CurrentUser.admin Then
          MsgBox("You are not authorized to use this reporting tool.", MsgBoxStyle.Exclamation, "Unauthorized")
          ThisWorkbook.Close()
        End If
      End If
    End If

  End Sub
  Public Sub SetupControlReferences()
    btnGetEmpReport = CType(Me.FindControl("btnGetEmpReport"), _
      MSForms.CommandButton)
    comboEmployeeNames = CType(Me.FindControl("comboEmployeeNames"), _
      MSForms.ComboBox)
    txtDateRangeStart = CType(Me.FindControl("txtDateRangeStart"), _
      MSForms.TextBox)
    txtDateRangeEnd = CType(Me.FindControl("txtDateRangeEnd"), _
      MSForms.TextBox)
  End Sub
  Public Sub SetupEmployeeDropDownList()
    UserArray = TimeManager.GetAllUsers()


    If Not UserArray Is Nothing Then
      comboEmployeeNames.Clear()
      For Each UserObj As TMW.tmUser In UserArray
        comboEmployeeNames.AddItem(UserObj.nameLast & ", " & _
          UserObj.nameFirst)
      Next
    End If

  End Sub


#End Region


#Region "Event Methods"

  Private Sub btnGetEmpReport_Click() Handles btnGetEmpReport.Click
    Dim startDate As Date = Nothing
    Dim endDate As Date = Nothing

    '---------------------------------------------------
    If Me.txtDateRangeStart.Text <> "" Then
      If Not IsDate(Me.txtDateRangeStart.Text) Then
        MsgBox("You must specify a valid date for the starting date.")
        Exit Sub
      Else
        startDate = CDate(Me.txtDateRangeStart.Text)
      End If
    End If

    '---------------------------------------------------
    If Me.txtDateRangeEnd.Text <> "" Then
      If Not IsDate(Me.txtDateRangeEnd.Text) Then
        MsgBox("You must specify a valid date for the ending date.")
        Exit Sub
      Else
        endDate = CDate(Me.txtDateRangeEnd.Text)
      End If
    End If

    Dim HourArray() As TMW.tmHours = TimeManager.GetProjectReportInfoByEmployee _
      (UserArray(Me.comboEmployeeNames.ListIndex).userId, startDate, endDate)

    'Clear cells
    Dim WS As Excel.Worksheet = ThisWorkbook.Worksheets(1)
    WS.Range("A4", "E2048").Clear()

    If Not HourArray Is Nothing Then
      For index As Integer = 4 To 4 + HourArray.Length - 1
        WS.Range("A" & index).Value = Format(HourArray(index - 4).startDate, "MM/dd/yyyy") & _
          " to " & Format(HourArray(index - 4).endDate, "MM/dd/yyyy")
        WS.Range("B" & index).Value = HourArray(index - 4).ProjectName
        WS.Range("C" & index).Value = HourArray(index - 4).hours
        WS.Range("D" & index).Value = HourArray(index - 4).description
      Next
    End If

  End Sub
  Private Sub ThisWorkbook_Open() Handles ThisWorkbook.Open

    DoLogin()
    CreateSettingsMenuItem()
    SetupControlReferences()
    SetupEmployeeDropDownList()


  End Sub


  Private Sub ThisWorkbook_BeforeClose(ByRef Cancel As Boolean) _
    Handles ThisWorkbook.BeforeClose
    Cancel = False
  End Sub
#End Region


#Region "Code Not Covered in Book"
  Public Sub CreateSettingsMenuItem()
    Const toolsMenuId As Integer = 30007
    Dim CB As Office.CommandBar = _
      ThisApplication.CommandBars("Worksheet Menu Bar")

    Dim ToolsMenu As Office.CommandBarPopup = CB.FindControl(id:=toolsMenuId)

    TimeManagerSettingsMenuButton = _
      ToolsMenu.Controls.Add(Type:=Office.MsoControlType.msoControlButton, _
      Temporary:=True)
    TimeManagerSettingsMenuButton.Caption = "Time Manager Reports"
  End Sub

  Private Sub TimeManagerSettingsMenuButton_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As _
    Boolean) Handles TimeManagerSettingsMenuButton.Click

    Dim SettingsDialog As New TimeManagerSettings
    SettingsDialog.ShowDialog()
    SetupEmployeeDropDownList()
  End Sub
#End Region

End Class
