Option Explicit On 

Imports OL = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Core
Imports System.Text.RegularExpressions


Public Class appOutlook

  Private Shared appOL As OL.Application
  Private Shared ns As OL.NameSpace

  Private Shared fldTemplates As OL.MAPIFolder

  Private Shared WithEvents usData As New UserSettings
  Private Shared WithEvents bmMainMenu As BravoMenu

#Region "Property Procedures"
  Friend Shared ReadOnly Property Application() As OL.Application
    Get
      Return appOL
    End Get

  End Property

  Friend Shared ReadOnly Property CurrentNamespace() As OL.NameSpace
    Get
      Return ns
    End Get
  End Property


  Friend Shared Property TemplatesFolder() As OL.MAPIFolder
    Get
      Return fldTemplates
    End Get
    Set(ByVal Value As OL.MAPIFolder)
      fldTemplates = Value
    End Set
  End Property


#End Region


#Region "Class Methods"


  Friend Shared Sub Setup(ByVal oApp As OL.Application)
    Try
      appOL = oApp
      ns = appOL.GetNamespace("MAPI")
      bmMainMenu = New BravoMenu
      fldTemplates = ns.GetFolderFromID(UserSettings.EntryID, _
        UserSettings.StoreID)

    Catch ex As Exception
      Select Case ex.Message
        Case "The operation cannot be performed because the" & _
          "object has been deleted."
          bmMainMenu = New BravoMenu

          MsgBox("You have not specified the location for your " & _
            "Bravo Corp Email Templates." & vbCrLf & vbCrLf & _
            "Please choose a location using the Settings form under the " & _
            "Bravo Tools Menu.", MsgBoxStyle.Critical, _
              "Missing Email Templates Folder")
        Case Else

      End Select
    End Try

  End Sub


  Friend Shared Sub ShutDown()
    appOL.Quit()
  End Sub


  Friend Shared Function PickOutlookFolder() As String()

    Dim fld As OL.MAPIFolder
    Dim str(2) As String

    Do

      fld = ns.PickFolder

      If fld Is Nothing Then
        Exit Function

      Else
        If fld.DefaultItemType = OL.OlItemType.olMailItem _
          Or OL.OlItemType.olPostItem Then

          str(0) = fld.FolderPath
          str(1) = fld.EntryID
          str(2) = fld.StoreID
          Return str
        Else
          MsgBox("Please pick a folder containing Mail or Post items.")

        End If
      End If

    Loop While fld.DefaultItemType <> OL.OlItemType.olMailItem Or OL.OlItemType.olPostItem


  End Function


  Private Shared Function LookupContact(ByVal FullName As String) _
    As OL.ContactItem

    Dim ci As OL.ContactItem
    Dim fldContacts As OL.MAPIFolder

    fldContacts = ns.GetDefaultFolder _
      (OL.OlDefaultFolders.olFolderContacts)

    ci = fldContacts.Items.Find("[FullName] = '" & FullName & "'")
    If Not TypeName(ci) = "Nothing" Then
      Return ci

    Else
      MsgBox("Contact not found.")
    End If

    Return ci
  End Function


  Private Shared Sub NavigateToTemplatesFolder()
    appOL.ActiveExplorer.CurrentFolder = fldTemplates

  End Sub


  Private Shared Function ScanAndFillTags(ByVal TemplateBody As String, _
    ByVal UserName As String)
    'Scan the string for tags, match up to properties of a ContactItem.
    'The UserName string is used to find the ci.
    Dim strNewBody As String
    strNewBody = TemplateBody

    Dim ci As OL.ContactItem
    ci = LookupContact(UserName)
    'Create the RegEx object using the ETE InfoTag pattern
    Dim re As New Regex("(?<tag><!::(?<value>.*)::>)", RegexOptions.IgnoreCase)
    Dim m As Match
    Dim gValue As Group
    Dim gTag As Group
    Dim strPropValue As String
    Try
      'Scan the email template body for matches
      For Each m In re.Matches(strNewBody)
        'For each match, extract the Value portion of the pattern
        'in order to match with ContactItem's property of the 
        'same name
        gValue = m.Groups("value")
        'retrieve corresponding value from the CI object
        strPropValue = ci.ItemProperties(gValue.Value).Value
        'Replace the entire InfoTag with the new value
        strNewBody = re.Replace(strNewBody, m.Groups("tag").Value, strPropValue)
      Next

      Return strNewBody
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try

  End Function


  Private Shared Function GetCurrentSelectionEmail() As String
    'Looksup the ContactItem for the currently selected MailItem
    'Uses the From email address to find them in the default Contacts folder.

    Dim miSelected As OL.MailItem
    'This will get the first selected item (in case there is more than one).
    Try

      miSelected = appOL.ActiveExplorer.Selection(1)

      Return miSelected.SenderEmailAddress
    Catch ex As Exception
      Return ""
    End Try


  End Function


  Private Shared Function CreateEmail(ByVal EmailText As String, _
    ByVal Subject As String)   ', ByVal Contact As OL.ContactItem
    'This is the last event in the workflow.
    '1. Use the text as the body of the email
    '2. Use the contact item to fill the TO field, etc...  This may need to just be a string
    Dim mi As OL.MailItem
    mi = appOL.CreateItem(OL.OlItemType.olMailItem)
    With mi

      .Body = EmailText.ToString
      .To = GetCurrentSelectionEmail()
      .Subject = Subject
      .Display()
    End With

  End Function


#End Region



#Region "Class Events"

  Private Shared Sub usData_AfterSettingsChange(ByVal EntryID As String, ByVal StoreID As String) Handles usData.AfterSettingsChange
    Try
      If appOL.Explorers.Count > 0 Then
        bmMainMenu.CreateTemplatesMenu()
      End If
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try

  End Sub
  Private Shared Sub bmMainMenu_TemplatesFolder_ItemAdd(ByVal Item _
    As Object) Handles bmMainMenu.TemplatesFolder_ItemAdd

    bmMainMenu.CreateTemplatesMenu()
  End Sub

  Private Shared Sub bmMainMenu_TemplatesFolder_ItemChange(ByVal Item _
    As Object) Handles bmMainMenu.TemplatesFolder_ItemChange

    bmMainMenu.CreateTemplatesMenu()
  End Sub

  Private Shared Sub bmMainMenu_TemplatesFolder_ItemRemove() _
    Handles bmMainMenu.TemplatesFolder_ItemRemove

    bmMainMenu.CreateTemplatesMenu()
  End Sub
  Private Shared Sub bmMainMenu_CreateEmailFromTemplate_Click(ByVal ctrl _
    As Microsoft.Office.Core.CommandBarButton, _
      ByRef CancelDefault As Boolean) _
      Handles bmMainMenu.CreateEmailFromTemplate_Click

    Dim mi As OL.MailItem
    Dim pi As OL.PostItem

    Try
      'Something funky here trying to get the Ol.OlItemType to show.  
      'Have to backspace and then it appears.
      mi = ns.Application.CreateItem(OL.OlItemType.olMailItem)
      pi = fldTemplates.Items(ctrl.Parameter)

      If Not pi Is Nothing Then
        Dim str As String = ScanAndFillTags(pi.Body, UserSettings.UserName)
        CreateEmail(str, pi.Subject)

      End If

    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try
  End Sub
  Private Shared Sub bmMainMenu_GoToTemplatesFolder_Click(ByVal _
    ctrl As Microsoft.Office.Core.CommandBarButton, _
      ByRef CancelDefault As Boolean) Handles _
      bmMainMenu.GoToTemplatesFolder_Click

    'fldTemplates.Display()
    'or
    NavigateToTemplatesFolder()
  End Sub
  Private Shared Sub bmMainMenu_SettingsButton_Click(ByVal ctrl _
    As Microsoft.Office.Core.CommandBarButton, _
      ByRef CancelDefault As Boolean) _
      Handles bmMainMenu.SettingsButton_Click

    Dim frmSettings As New Settings
    frmSettings.Show()
  End Sub


#End Region

End Class
