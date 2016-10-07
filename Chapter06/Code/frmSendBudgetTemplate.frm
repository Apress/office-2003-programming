VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSendBudgetTemplate 
   Caption         =   "Distribute Budget Templates"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   OleObjectBlob   =   "frmSendBudgetTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSendBudgetTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub GetContacts()
On Error GoTo eHandler

'We are going to use late-binding b/c who knows what version of OL the user may have
  Dim olApp As New Outlook.Application
  Dim oNS As Outlook.Namespace
  Dim oContactFolder As Outlook.MAPIFolder
  Dim oContact As Outlook.ContactItem
  
  Set oNS = olApp.GetNamespace("MAPI")
  Set oContactFolder = oNS.GetDefaultFolder(olFolderContacts)
  
  For Each oContact In oContactFolder.Items
    'We only want contact items
    If oContact.MessageClass = "IPM.Contact" Then
      'We only want contact items with email addresses
      If Len(oContact.Email1Address) > 0 Then
        lstContacts.AddItem oContact.LastNameAndFirstName, oContact.Email1Address
      End If
    End If
   Next oContact
  
  Exit Sub
eHandler:
  MsgBox Err.Description
  
  
End Sub


