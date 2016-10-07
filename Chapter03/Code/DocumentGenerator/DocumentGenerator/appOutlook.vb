Option Explicit On 

Imports OL = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Core
Imports System.Windows.Forms


Public Class appOutlook
  Private Shared appOL As OL.Application
  Private Shared fldClients As OL.MAPIFolder
  Private Shared cbbBravoMenu As CommandBarPopup
  Private Shared WithEvents cbbCreateDocs As CommandBarButton
  Private Shared WithEvents cbbSettings As CommandBarButton
  Private Shared WithEvents cbbGetUpdates As CommandBarButton

  Friend Shared Sub Setup(ByVal oApp As OL.Application, _
    ByVal EntryID As String, ByVal StoreID As String)

    appOL = oApp
    fldClients = GetFolder(EntryID, StoreID)
    SetupBravoMenu()
  End Sub

  Friend Shared Sub ShutDown()
    appOL.Quit()
  End Sub

  Private Shared Function SetupBravoMenu()
    Dim cbCommandBars As CommandBars
    Dim cbMenuBar As CommandBar

    ' Outlook has the CommandBars collection on the Explorer object.
    cbCommandBars = appOL.ActiveExplorer.CommandBars
    cbMenuBar = cbCommandBars.Item("Menu Bar")

    ' In case the button was not deleted, use the exiting one.
    cbbBravoMenu = cbMenuBar.FindControl(tag:="Bravo Tools")
    If cbbBravoMenu Is Nothing Then
      cbbBravoMenu = cbMenuBar.Controls.Add( _
        Type:=MsoControlType.msoControlPopup, _
        Before:=6, Temporary:=True)

      With cbbBravoMenu
        .Caption = "&Bravo Tools"
        .Tag = "Bravo Tool"
        .TooltipText = "Bravo Crop Tools Menu"
        .OnAction = "!<DocumentGenerator.Connect>"
        .Visible = True
        'Create Documents Button
        cbbCreateDocs = .Controls.Add( _
          Type:=MsoControlType.msoControlButton, _
          Temporary:=True)
        With cbbCreateDocs
          .Caption = "Create Project Documents..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Create a set of Project Documents to send to a client."
          .OnAction = "!<DocumentGenerator.Connect>"
          .Visible = True
        End With
        'Settings Button
        cbbSettings = .Controls.Add( _
          Type:=MsoControlType.msoControlButton, _
          Temporary:=True)
        With cbbSettings
          .Caption = "User Settings..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Change the Document Generator User Settings."
          .OnAction = "!<DocumentGenerator.Connect>"
          .Visible = True
        End With
        'Download Updates Button
        cbbGetUpdates = .Controls.Add( _
          Type:=MsoControlType.msoControlButton, _
          Temporary:=True)
        With cbbGetUpdates
          .Caption = "Download Templates..."
          .Style = MsoButtonStyle.msoButtonCaption
          .Tag = "Download additional Document Templates.."
          .OnAction = "!<DocumentGenerator.Connect>"
          .Visible = True
        End With
      End With
    End If
  End Function

  Private Shared Function GetFolder(ByVal EID As String, _
    ByVal SID As String) As OL.MAPIFolder

    Dim nsMAPI As OL.NameSpace
    Dim fld As OL.MAPIFolder
    Try
      nsMAPI = appOL.Application.GetNamespace("MAPI")
      fld = nsMAPI.GetFolderFromID(EID, SID)

      Return fld
    Catch ex As Exception

    End Try
  End Function

  Friend Shared Function PickContactsFolder() As String()
    Dim nsMAPI As OL.NameSpace
    Dim fld As OL.MAPIFolder
    Dim str(2) As String
    nsMAPI = appOL.Application.GetNamespace("MAPI")

    Do

      fld = nsMAPI.PickFolder
      If fld Is Nothing Then
        Exit Function
      End If

      If fld.DefaultItemType <> OL.OlItemType.olContactItem Then
        MsgBox("Please pick a folder containing Contact Items.")
      Else
        str(0) = fld.FolderPath
        str(1) = fld.EntryID
        str(2) = fld.StoreID
        Return str
      End If
    Loop While fld.DefaultItemType <> OL.OlItemType.olContactItem
  End Function

  Friend Shared Function ListContacts(ByVal ctl As ComboBox)
    Dim fld As OL.MAPIFolder
    Dim itms As OL.Items
    Dim itm As OL.ContactItem
    Dim strName As String
    Dim i As Integer


    itms = fldClients.Items
    itms.Sort("[LastName]", False)

    For i = 1 To itms.Count
      itm = itms(i)
      strName = itm.FirstName & " " & itm.LastName
      ctl.Items.Add(strName)

    Next i
  End Function

  Friend Shared Function GetContact(ByVal ContactName As String) _
    As OL.ContactItem

    Dim itms As OL.Items
    Dim itm As OL.ContactItem
    Try
      itms = fldClients.Items
      itm = itms(ContactName)

      Return itm
    Catch ex As Exception

    End Try


  End Function

  Friend Shared Sub CreateEmail(ByVal FileNames() As String, _
    ByVal ClientName As String)

    Dim mi As OL.MailItem
    Dim ci As OL.ContactItem
    Dim i As Integer
    Dim TotalCount As Integer
    TotalCount = FileNames.GetUpperBound(0)

    ci = GetContact(ClientName)
    mi = appOL.CreateItem(OL.OlItemType.olMailItem)
    mi.To = ci.Email1Address
    mi.Subject = "Event Documents for " & ci.CompanyName

    For i = 0 To TotalCount - 1
      mi.Attachments.Add(FileNames(i))
    Next

    mi.Display()

  End Sub


#Region "CommandBarButton Code"

  Private Shared Sub cbbCreateDocs_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) _
    Handles cbbCreateDocs.Click


    Try
      Dim frmDocuments As New Documents
      frmDocuments.Show()
    Catch ex As Exception

    End Try
  End Sub

  Private Shared Sub cbbSettings_Click(ByVal Ctrl As _
    Microsoft.Office.Core.CommandBarButton, _
    ByRef CancelDefault As Boolean) Handles cbbSettings.Click

    Dim frmSettings As New Settings
    frmSettings.Show()

  End Sub

  Private Shared Sub cbbGetUpdates_Click(ByVal Ctrl _
    As Microsoft.Office.Core.CommandBarButton, _
    ByRef CancelDefault As Boolean) Handles cbbGetUpdates.Click

    Dim frm As New GetUpdates
    frm.ShowDialog()
  End Sub
#End Region



End Class
