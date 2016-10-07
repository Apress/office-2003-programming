Imports Microsoft.Office.Core
Imports Extensibility
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the PPTGEN2003Setup project 
' by right clicking the project in the Solution Explorer, then choosing install.
#End Region

<GuidAttribute("DA9F48BB-E025-4EA3-B7FA-ACBE8C2290B4"), ProgIdAttribute("PresentationGenerator.Connect")> _
Public Class Connect
  Implements Extensibility.IDTExtensibility2

  Dim appPowerPoint As Object
  Dim addInInstance As Object
  Dim cbbBravoMenu As CommandBarPopup 'Root Command Bar for the following buttons
  Dim WithEvents cbbNewPPTFromTemplate As CommandBarButton  'Create New Presentation Button
  Dim WithEvents cbbSaveBravoPPT As CommandBarButton  'Save Presentation Button
  Dim WithEvents cbbGetUpdates As CommandBarButton  'Get updates from SQL Database record
  Dim WithEvents cbbSettings As CommandBarButton  'Edit User Setttings button

#Region "IDTExtensibility2 Subs"

  Public Sub OnConnection(ByVal application As Object, ByVal connectMode As Extensibility.ext_ConnectMode, _
    ByVal addInInst As Object, ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnConnection

    appPowerPoint = application
    AppPPT.Setup(application)
    UserSettings.LoadSettings(System.Windows.Forms.Application.StartupPath)
    addInInstance = addInInst


    ''' If you aren't in startup, manually call OnStartupComplete.
    ''If (connectMode <> Extensibility.ext_ConnectMode.ext_cm_Startup) Then
    ''  Call OnStartupComplete(custom)
    ''End If
  End Sub


  Public Sub OnStartupComplete(ByRef custom As System.Array) Implements _
    Extensibility.IDTExtensibility2.OnStartupComplete

    Dim cbCommandBars As CommandBars
    Dim cbMenuBar As CommandBar
    Dim iToolsMenuPosition As Integer

    cbCommandBars = appPowerPoint.commandbars
    cbMenuBar = cbCommandBars.Item("Menu Bar")
    iToolsMenuPosition = cbMenuBar.Controls("Tools").Index
    iToolsMenuPosition = iToolsMenuPosition + 1

    If Not DoesMenuExist(cbMenuBar, "Bravo Tools") Then 'TODO:  Update Chapter 2 to reflect call to DoesMenuExist
      cbbBravoMenu = cbMenuBar.Controls.Add(Type:=MsoControlType.msoControlPopup, _
                    before:=iToolsMenuPosition)

      cbbBravoMenu.Tag = "Bravo Tools" 'TODO: update in book.

      With cbbBravoMenu
        .Caption = "&Bravo Tools"

        '=========Create New PPT Button============
        cbbNewPPTFromTemplate = .Controls.Add(MsoControlType.msoControlButton) 'Todo: Update this line in the book
        With cbbNewPPTFromTemplate
          .Caption = "Ne&w PPT..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Generate New Presentation from Template"
          .OnAction = "!<PresentationGenerator.Connect>"
          .Visible = True
        End With

        '=========Create cbbSaveBravoPPT Button============
        cbbSaveBravoPPT = .Controls.Add(MsoControlType.msoControlButton) 'Todo: Update this line in the book
        With cbbSaveBravoPPT
          .Caption = "&Save Presentation..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Save Bravo Presentation to Bravo Folder"
          .OnAction = "!<PresentationGenerator.Connect>"
          .Visible = True
        End With

        '=========Create Update Database Button============
        cbbGetUpdates = .Controls.Add(MsoControlType.msoControlButton) 'Todo: Update this line in the book
        With cbbGetUpdates
          .Caption = "&Update Local Templates"
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Get the latest templates from Headquarters."
          .OnAction = "!<PresentationGenerator.Connect>"
          .Visible = True
        End With

        '=========Create Settings Button============
        cbbSettings = .Controls.Add(MsoControlType.msoControlButton) 'Todo: Update this line in the book
        With cbbSettings
          .Caption = "Se&ttings..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Change Add-In Settings."
          .OnAction = "!<PresentationGenerator.Connect>"
          .Visible = True
        End With

      End With

    End If

    cbMenuBar = Nothing
    cbCommandBars = Nothing

  End Sub


  Public Sub OnBeginShutdown(ByRef custom As System.Array) _
    Implements Extensibility.IDTExtensibility2.OnBeginShutdown

    cbbNewPPTFromTemplate.Delete()
    cbbGetUpdates.Delete()
    cbbSettings.Delete()
    cbbBravoMenu.Delete()
    AppPPT.ShutDown()


    cbbGetUpdates = Nothing
    cbbSettings = Nothing
    cbbNewPPTFromTemplate = Nothing

  End Sub


  Public Sub OnDisconnection(ByVal RemoveMode As Extensibility.ext_DisconnectMode, _
    ByRef custom As System.Array) _
    Implements Extensibility.IDTExtensibility2.OnDisconnection


    'If RemoveMode <> Extensibility.ext_DisconnectMode.ext_dm_HostShutdown Then
    '  Call OnBeginShutdown(custom)
    'End If

    appPowerPoint = Nothing
  End Sub


  Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
    '
  End Sub

#End Region

#Region "CustomMethods"
  Private Function DoesMenuExist(ByVal Menu As CommandBar, ByVal Tag As String) As Boolean
    'TODO:  Update Chapter 2 to reflect call to DoesMenuExist
    Dim cbc As CommandBarControl

    cbc = Menu.FindControl(Tag:=Tag)
    If Not cbc Is Nothing Then
      Return True
    Else
      Return False
    End If

  End Function

#End Region


#Region "Command Bar Buttons' Code"

  Private Sub cbbNewPPTFromTemplate_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) _
    Handles cbbNewPPTFromTemplate.Click

    Dim frmList As New ListPresentations
    frmList.Show()
  End Sub

  Private Sub cbbGetUpdates_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) _
    Handles cbbGetUpdates.Click

    Dim frmDBF As New GetUpdates
    frmDBF.Show()
  End Sub

  Private Sub cbbSettings_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) _
    Handles cbbSettings.Click

    Dim frmSettings As New Settings
    frmSettings.Show()
  End Sub

  Private Sub cbbSaveBravoPPT_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, _
    ByRef CancelDefault As Boolean) _
    Handles cbbSaveBravoPPT.Click

    Dim sfdFile As New SaveFileDialog

    With sfdFile
      .Title = "Save Bravo Presentation"
      'Change the FileDialog’s initial folder to the one 
      'specified in the UserSettings
      .InitialDirectory = UserSettings.SaveFolder
      .Filter = "PowerPoint files (*.ppt)|*.ppt|All files (*.*)|*.*"
      'Set .ppt as default type (the first filter)
      .FilterIndex = 1
      'notify user if file exists
      .OverwritePrompt = True
      'Restore the Initial Directory settings back to the user's default.
      .RestoreDirectory = True
      .InitialDirectory = UserSettings.SaveFolder()

      If .ShowDialog() = DialogResult.OK Then
        Dim bExists As Boolean = .CheckFileExists()
        If Not bExists Then
          'User Powerpoint's save function to save file to location just specified
          AppPPT.App.ActivePresentation.SaveAs(.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation, -2)
        End If
      End If

    End With
  End Sub

#End Region

End Class
