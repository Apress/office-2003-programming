Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports MSForms = Microsoft.Vbe.Interop.Forms

' Office integration attribute. Identifies the startup class for the workbook. Do not modify.
<Assembly: System.ComponentModel.DescriptionAttribute("OfficeStartupClass, Version=1.0, Class=TimeManagerEntry.OfficeCodeBehind")>

Public Class OfficeCodeBehind 

  Friend WithEvents ThisWorkbook As Excel.Workbook
  Friend WithEvents ThisApplication As Excel.Application
  Friend CurrentUser As TMW.tmUser = Nothing
  Friend TimeManagerWeb As New TMW.TimeManagerWeb

  'Private HiddenItems As New ArrayList
  Private InAutoSave As Boolean = False

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

  Public Sub AutoSave()
    InAutoSave = True
    ThisWorkbook.Save()
    InAutoSave = False
  End Sub
  Public Sub CreateProjectList()
    Dim ProjectList As Object() = TimeManagerWeb.GetAllProjects()
    Dim WS As Excel.Worksheet = ThisWorkbook.Worksheets(2)
    Dim Index As Integer

    'Clear our the current project contents
    Try
      WS.Range("A1", "A255").Clear()
    Catch ex As Exception
      Dim x As String = ex.Message
    End Try

    If Not ProjectList Is Nothing Then
      For Index = 1 To ProjectList.Length
        WS.Range("A" & Index).Value = CStr(ProjectList(Index - 1))
      Next
      ThisWorkbook.Names.Item("ProjectList").Value = "=Projects!$A$1:$A$" & _
        ProjectList.Length
    Else
      ThisWorkbook.Names.Item("ProjectList").Value = "=Projects!$A$1:$A$1"
    End If

  End Sub

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
      CurrentUser = TimeManagerWeb.Login(loginForm.userId, loginForm.password)
      If CurrentUser Is Nothing Then
        MsgBox("The credential you supplied were invalid.  Please try again")
        loginForm = Nothing
        DoLogin()
      End If
    End If

  End Sub

  Public Function SaveContentsToWeb(ByVal StartDate As DateTime) As Boolean
    '-----------------------------------------------------------------------
    '   INFO
    '-----------------------------------------------------------------------

    Dim WS As Excel.Worksheet = ThisWorkbook.Worksheets.Item(1)
    Dim index As Integer = 4
    Dim done As Boolean
    Dim obj As TMW.tmHours
    Dim objCol As New ArrayList
    Dim ErrorFlag As Boolean
    Dim endDate As Date

    While Not done
      obj = GetHoursFromWorksheet(index, StartDate, ErrorFlag)
      If obj Is Nothing Then
        done = True
      Else
        endDate = obj.endDate
        objCol.Add(obj)
      End If
      index += 1
    End While

    If objCol.Count > 0 Then
      Return TimeManagerWeb.SaveHourArrayList(CurrentUser.userId, _
        StartDate, endDate, objCol.ToArray)
    Else
      Return Not ErrorFlag
    End If

  End Function

  Public Function GetHoursFromWorksheet(ByVal index As Integer, _
    ByVal StartDate As DateTime, ByRef ErrorFlag As Boolean) As TMW.tmHours

    Dim hoursObj As New TMW.tmHours
    Dim ws As Excel.Worksheet = ThisWorkbook.Worksheets.Item(1)

    hoursObj.userId = CurrentUser.userId
    hoursObj.startDate = CDate(Format(StartDate, "MM/dd/yyyy") & " 12:00 AM")
    hoursObj.endDate = CDate(Format(StartDate.AddDays(7), "MM/dd/yyyy") & _
      " 11:59:59 PM")

    If CStr(ws.Range("A" & index).Value) = String.Empty Then
      Return Nothing
    Else
      hoursObj.ProjectName = ws.Range("A" & index).Value
    End If

    If CStr(ws.Range("B" & index).Value) = String.Empty Then
      MsgBox("You must enter a description of the work you completed for " & _
        "the project entitled " & hoursObj.ProjectName, _
        MsgBoxStyle.Exclamation Or MsgBoxStyle.OKCancel, "Error")

      ws.Range("B" & index).Select()
      ErrorFlag = True
      Return Nothing
    Else
      hoursObj.description = ws.Range("B" & index).Value
    End If

    If CStr(ws.Range("C" & index).Value) = String.Empty Then
      MsgBox("You must enter the number of hours you worked on the " & _
      "project entitled " & hoursObj.ProjectName, MsgBoxStyle.Exclamation Or _
         MsgBoxStyle.OKCancel, "Error")
      ws.Range("C" & index).Select()
      ErrorFlag = True
      Return Nothing
    Else
      Try
        hoursObj.hours = CSng(ws.Range("C" & index).Value)
      Catch ex As Exception
        MsgBox("You must enter a numeric value for the number of hours " & _
          "you worked on the project entitled " & hoursObj.ProjectName, _
          MsgBoxStyle.Exclamation Or MsgBoxStyle.OKCancel, "Error")
        ws.Range("C" & index).Select()
        ErrorFlag = True
        Return Nothing
      End Try
    End If

    Return hoursObj

  End Function


  Private Sub ThisWorkbook_Open() Handles ThisWorkbook.Open
    '-----------------------------------------------------------------------
    '   The following code will be executed when the workbook is opened
    '-----------------------------------------------------------------------

    DoLogin()
    CreateProjectList()
    AutoSave()

  End Sub

 
  Private Sub ThisWorkbook_BeforeSave(ByVal SaveAsUI As Boolean, _
    ByRef Cancel As Boolean) Handles ThisWorkbook.BeforeSave

    If InAutoSave Then Return

    Dim frmSaveDialog As New SaveDialog
    Dim WS As Excel.Worksheet

    If frmSaveDialog.ShowDialog = DialogResult.OK Then
      If frmSaveDialog.SaveToWeb Then
        If SaveContentsToWeb(frmSaveDialog.StartDay) Then
          If frmSaveDialog.ClearContents Then
            WS = ThisWorkbook.Worksheets(1)
            For index As Integer = 4 To 4 + 255
              WS.Range("A" & index).Value = ""
              WS.Range("B" & index).Value = ""
              WS.Range("C" & index).Value = ""
            Next
          End If
          MsgBox("Your information was successfully saved to the " & _
            "Time Management system.", MsgBoxStyle.Information, _
            "Successfully Saved")
        Else
          MsgBox("Your information was NOT uploaded to the Time" & _
            "Management system.", MsgBoxStyle.Exclamation, _
            "Error Uploading Data")
        End If
      End If
    Else
      MsgBox("None of your changes have been saved", _
        MsgBoxStyle.Information Or MsgBoxStyle.OKOnly, "Save Cancelled")
      Cancel = True
    End If

  End Sub

End Class
