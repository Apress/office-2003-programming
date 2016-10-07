'NOT IN USE!!


Imports OL = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Core

Public Class BravoMenu

#Region "Declarations"
  '====================================================
  'For the Main Bravo Menu

  Private cbMenuBar As CommandBar
  Private cbbBravoMenu As CommandBarPopup
  Private cbbTemplates As CommandBarPopup

  Private WithEvents cbbCreateEmailFromTemplate As CommandBarButton
  Private WithEvents cbbSettings As CommandBarButton
  Private WithEvents cbbGoToTemplatesFolder As CommandBarButton
  Private WithEvents itmTemplates As OL.Items

  'Events
  Public Event SettingsButton_Click(ByVal ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
  Public Event GoToTemplatesFolder_Click(ByVal ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
  Public Event CreateEmailFromTemplate_Click(ByVal ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
  Public Event TemplatesFolder_ItemAdd(ByVal Item As Object)
  Public Event TemplatesFolder_ItemChange(ByVal Item As Object)
  Public Event TemplatesFolder_ItemRemove()
#End Region


#Region "Methods"
  Public Sub New()
    InitializeMenu()
  End Sub

  Private Function InitializeMenu()
    Dim cbCommandBars As CommandBars
    'Dim cbMenuBar As CommandBar

    cbCommandBars = appOutlook.Application.ActiveExplorer.CommandBars
    cbMenuBar = cbCommandBars.Item("Menu Bar")

    cbbBravoMenu = cbMenuBar.FindControl(tag:="Bravo Tools")

    If cbbBravoMenu Is Nothing Then
      cbbBravoMenu = cbMenuBar.Controls.Add( _
        Type:=MsoControlType.msoControlPopup, _
        Before:=6, Temporary:=True)


      With cbbBravoMenu
        .Caption = "&Bravo Tools"
        .Tag = "Bravo Tools Menu"
        .OnAction = "!<EmailTemplates.Connect>"
        .Visible = True

        '===  ===  ===  ===  === 
        'Create Templates Button
        cbbTemplates = cbbBravoMenu.Controls.Add( _
        Type:=MsoControlType.msoControlPopup, _
         Temporary:=True)
        With cbbTemplates
          .Caption = "Email Templates"
          .Tag = "Email Templates"
          .OnAction = "!<EmailTemplates.Connect>"
          .Visible = True
        End With

        'GoTo Button
        cbbGoToTemplatesFolder = .Controls.Add( _
          Type:=MsoControlType.msoControlButton, _
          Temporary:=True)
        With cbbGoToTemplatesFolder
          .Caption = "GoTo Templates Folder"
          .Style = MsoButtonStyle.msoButtonIconAndCaption
          .Tag = "Navigate to the Email Templates Folder"
          .OnAction = "!<EmailTemplates.Connect>"
          .FaceId = 1589
          .Visible = True
        End With


        'Settings Button
        cbbSettings = .Controls.Add( _
          Type:=MsoControlType.msoControlButton, _
          Temporary:=True)
        With cbbSettings
          .Caption = "User Settings..."
          .Style = MsoButtonStyle.msoButtonIconAndCaption
          .Tag = "Change the Document Generator User Settings."
          .OnAction = "!<EmailTemplates.Connect>"

          .BeginGroup = True
          .Visible = True
        End With

      End With

      CreateTemplatesMenu()


    End If

  End Function

  Public Sub CreateTemplatesMenu()
    Dim fldTemplates As OL.MAPIFolder
    Try
      Dim itm As Object
      Dim i As Integer

      i = 1
      fldTemplates = appOutlook.CurrentNamespace.GetFolderFromID _
        (UserSettings.EntryID, UserSettings.StoreID)

      If fldTemplates.Items.Count > 0 Then
        itmTemplates = fldTemplates.Items
        If cbbTemplates.Controls.Count > 0 Then
          Dim iCount As Integer
          For iCount = 1 To cbbTemplates.Controls.Count
            'The collection reshuffles, must always delete item #1 
            'or will an 'Invalid Index error'
            cbbTemplates.Controls(1).Delete()
          Next
        End If

        For Each itm In itmTemplates
          If TypeOf itm Is OL.PostItem Or TypeOf itm Is OL.MailItem Then
            AddTemplateCommandButton(itm.Categories, itm.Subject)
            i += 1
          End If
        Next itm

      End If

      'store reference to the folder for use later on
      appOutlook.TemplatesFolder = fldTemplates

    Catch ex As Exception
      If Err.Number = 91 Then
        MsgBox("The Email Templates Engine's settings file does not exist." & _
         vbCrLf & vbCrLf & _
         "Please set your settings now.", MsgBoxStyle.Information, _
         "Settings Not Found")
      Else
        MsgBox(ex.Message, "CreateTemplatesMenu")
      End If

    End Try

  End Sub

  Private Function AddTemplateCommandButton(ByVal _
    NewCategoryButtonName As String, ByVal NewButtonName As String) _
    As CommandBarButton
    'Dim cbbNew As CommandBarButton
    Dim cbpRoot As CommandBarPopup


    cbpRoot = GetCategoryMenu(NewCategoryButtonName, cbbTemplates)
    cbbCreateEmailFromTemplate = cbpRoot.Controls.Add _
      (Type:=MsoControlType.msoControlButton, Temporary:=True)

    With cbbCreateEmailFromTemplate
      .Caption = NewButtonName
      .Style = MsoButtonStyle.msoButtonIconAndCaption
      'These must be the same in order to respond to the same event.
      .Tag = "Email Template"
      .FaceId = 1757
      .OnAction = "!<EmailTemplates.Connect>"
      .Visible = True
      .Parameter = NewButtonName

    End With

    Return cbbCreateEmailFromTemplate
  End Function

  Private Shared Function GetCategoryMenu(ByVal MenuNameToFind As String, _
    ByVal MenuControl As CommandBarPopup) As CommandBarPopup
    'Looks for a menu of passed Category name.
    'If it does not exist, it is created.
    Dim ctl As CommandBarPopup
    Dim ctlSought As CommandBarPopup
    Dim fExists As Boolean

    fExists = False

    'First check that controls exist on the CommandBar.  
    'If not, we know we should create one
    If MenuControl.Controls.Count > 0 Then
      'loop through all countrols to determine if the control exists
      For Each ctl In MenuControl.Controls
        If ctl.Tag = MenuNameToFind Then
          fExists = True
          ctlSought = ctl
        End If
      Next


      'If the control exists, use it.
      If fExists Then
        Return ctlSought

      Else
        ctlSought = MenuControl.Controls.Add _
          (Type:=MsoControlType.msoControlPopup, Temporary:=True)

        With ctlSought
          .Caption = MenuNameToFind
          .Tag = MenuNameToFind
          .OnAction = "!<EmailTemplates.Connect>"
          .Visible = True
        End With

        Return ctlSought

      End If

    Else
      ctlSought = MenuControl.CommandBar.Controls.Add _
      (Type:=MsoControlType.msoControlPopup, Temporary:=True)


      With ctlSought
        .Caption = MenuNameToFind
        .Tag = MenuNameToFind
        .OnAction = "!<EmailTemplates.Connect>"
        .Visible = True
      End With

      Return ctlSought
    End If

  End Function


#End Region



#Region "Events"

  Private Sub cbbSettings_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef _
    CancelDefault As Boolean) Handles cbbSettings.Click

    RaiseEvent SettingsButton_Click(Ctrl, CancelDefault)
  End Sub
  Private Sub cbbGoToTemplatesFolder_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, _
    ByRef CancelDefault As Boolean) Handles cbbGoToTemplatesFolder.Click

    RaiseEvent GoToTemplatesFolder_Click(Ctrl, CancelDefault)
  End Sub

  Private Sub cbbCreateEmailFromTemplate_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, _
    ByRef CancelDefault As Boolean) _
    Handles cbbCreateEmailFromTemplate.Click

    RaiseEvent CreateEmailFromTemplate_Click(Ctrl, CancelDefault)
  End Sub
  Private Sub itmTemplates_ItemAdd(ByVal Item As Object) _
      Handles itmTemplates.ItemAdd

    RaiseEvent TemplatesFolder_ItemAdd(Item)
  End Sub

  Private Sub itmTemplates_ItemChange(ByVal Item As Object) _
    Handles itmTemplates.ItemChange

    RaiseEvent TemplatesFolder_ItemChange(Item)
  End Sub

  Private Sub itmTemplates_ItemRemove() Handles itmTemplates.ItemRemove
    RaiseEvent TemplatesFolder_ItemRemove()
  End Sub




#End Region

End Class