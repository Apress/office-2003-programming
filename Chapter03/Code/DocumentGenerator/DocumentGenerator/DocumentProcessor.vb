Option Explicit On 

Imports W = Microsoft.Office.Interop.Word
Imports OL = Microsoft.Office.Interop.Outlook
Imports Frm = System.Windows.Forms


Public Class DocumentProcessor
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call
    Try
      m_appW = GetObject(, "Word.Application")
      m_WordIsRunning = True
    Catch ex As Exception
      If Err.Number = 429 Then
        m_appW = CreateObject("Word.Application")
        m_WordIsRunning = False
      Else
        Throw New System.Exception("Microsoft Word Automation error.")
      End If

    End Try

  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If

    If m_WordIsRunning Then
      m_Doc.Close()
    Else
      m_Doc.Close()
      m_appW.Quit()
    End If

    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents pnlTitle As System.Windows.Forms.Panel
  Friend WithEvents lblTitle As System.Windows.Forms.Label
  Friend WithEvents cbOK As System.Windows.Forms.Button
  Friend WithEvents cbCancel As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DocumentProcessor))
    Me.pnlTitle = New System.Windows.Forms.Panel
    Me.lblTitle = New System.Windows.Forms.Label
    Me.cbOK = New System.Windows.Forms.Button
    Me.cbCancel = New System.Windows.Forms.Button
    Me.pnlTitle.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnlTitle
    '
    Me.pnlTitle.BackColor = System.Drawing.Color.White
    Me.pnlTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.pnlTitle.Controls.Add(Me.lblTitle)
    Me.pnlTitle.Location = New System.Drawing.Point(0, 0)
    Me.pnlTitle.Name = "pnlTitle"
    Me.pnlTitle.Size = New System.Drawing.Size(496, 50)
    Me.pnlTitle.TabIndex = 0
    '
    'lblTitle
    '
    Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblTitle.Location = New System.Drawing.Point(16, 16)
    Me.lblTitle.Name = "lblTitle"
    Me.lblTitle.Size = New System.Drawing.Size(352, 16)
    Me.lblTitle.TabIndex = 0
    Me.lblTitle.Text = "Bookmarks within selected document"
    '
    'cbOK
    '
    Me.cbOK.Location = New System.Drawing.Point(280, 384)
    Me.cbOK.Name = "cbOK"
    Me.cbOK.TabIndex = 1
    Me.cbOK.Text = "&OK"
    '
    'cbCancel
    '
    Me.cbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cbCancel.Location = New System.Drawing.Point(360, 384)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.TabIndex = 2
    Me.cbCancel.Text = "C&ancel"
    '
    'DocumentProcessor
    '
    Me.AcceptButton = Me.cbOK
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.CancelButton = Me.cbCancel
    Me.ClientSize = New System.Drawing.Size(442, 422)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.cbOK)
    Me.Controls.Add(Me.pnlTitle)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "DocumentProcessor"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Document Processor"
    Me.pnlTitle.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

#Region "Variable Declarations"
  Private m_appW As W.Application
  Private m_Doc As W.Document                   'The Word Document
  Private m_OpenPath As String                  'Folder path where Template file is located
  Private m_SavePath As String                  'Folder path to save files
  Private m_FileName As String                  'Name of Template File
  Private m_SavedPathAndFileName As String      'Name and Path of the Saved file
  Private m_Contact As OL.ContactItem           '
  Private m_Documents As System.Windows.Forms.ListBox
  Private m_WordIsRunning As Boolean
  Private m_DocSaved As Boolean
#End Region

#Region "Methods"

  Public Function OpenDoc() As Boolean
    Try
      With m_appW
        .DisplayAlerts = W.WdAlertLevel.wdAlertsNone
        .Documents.Open(m_OpenPath & "\" & m_FileName)
        m_Doc = .ActiveDocument
        .DisplayAlerts = W.WdAlertLevel.wdAlertsAll
      End With
      Return True
    Catch ex As Exception
      Return False

    End Try
  End Function

  Public Sub PopulateForm()
    Dim bm As W.Bookmark
    Dim i As Integer
    Dim iTop As Integer
    Dim itms As OL.ItemProperties
    Dim itm As OL.ItemProperty
    Dim val As String

    Try
      iTop = pnlTitle.Top + 100
      itms = m_Contact.ItemProperties

      For Each bm In m_Doc.Bookmarks
        Dim lbl As New Windows.Forms.Label
        Dim txt As New Windows.Forms.TextBox

        With lbl
          .Text = bm.Name
          .Left = 10
          .Width = 150
          .Top = iTop
          .Visible = True
        End With

        If PropertyExists(m_Contact, bm.Name) Then
          itm = itms(bm.Name)
          val = itm.Value
        Else
          val = ""
        End If

        With txt
          .Name = bm.Name
          .Text = val
          .Width = 250
          .Left = lbl.Width + 10
          .Top = iTop
          .Visible = True
        End With

        Me.Controls.Add(txt)
        Me.Controls.Add(lbl)

        iTop = iTop + 22
      Next


    Catch ex As Exception
      Throw New System.Exception("An exception has occurred.")
    End Try
  End Sub

  Private Function PropertyExists(ByVal Contact As OL.ContactItem, _
    ByVal PropertyName As String) As Boolean

    Dim prop As OL.ItemProperty

    prop = Contact.ItemProperties(PropertyName)

    If Not prop Is Nothing Then
      Return True
    Else
      Return False
    End If

  End Function

  Private Sub UpdateBookMarks()
    Dim ctl As Frm.Control

    For Each ctl In Me.Controls
      If TypeOf (ctl) Is Frm.TextBox Then
        SetBookMark(ctl)
      End If
    Next

  End Sub

  Private Sub SetBookMark(ByVal txt As Frm.TextBox)
    Dim rng As W.Range
    Dim str As String
    str = txt.Name.ToString

    If m_Doc.Bookmarks.Exists(str) Then
      rng = m_Doc.Bookmarks(str).Range
      rng.InsertBefore(txt.Text)
      'Uncomment this line to redefine the bookmark using the current range
      ''rng.Bookmarks.Add(str)
    End If

  End Sub

  Private Function SaveDoc() As String
    Try
      m_appW.ActiveDocument.Fields.Update()
      m_appW.ActiveDocument.SaveAs(m_SavePath & "\" & _
        m_Contact.CompanyName & " - " & m_FileName)

      m_DocSaved = True
      m_SavedPathAndFileName = m_Contact.CompanyName & " - " & m_FileName

    Catch ex As Exception
      m_DocSaved = False
      MsgBox(ex.GetBaseException)
    End Try

  End Function

#End Region

#Region "Events"

  Private Sub cbOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbOK.Click

    UpdateBookMarks()

    Me.SaveDoc()
    Me.Close()
  End Sub

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()
  End Sub


#End Region

#Region "CustomProps"

  Public ReadOnly Property GetBookmarks() As W.Bookmarks
    Get
      If Not m_Doc Is Nothing Then
        Return m_Doc.Bookmarks()
      End If
    End Get
  End Property

  Public ReadOnly Property DocSaved() As Boolean
    Get
      Return m_DocSaved
    End Get
  End Property

  Public Property OpenPath() As String
    Get
      Return m_OpenPath
    End Get
    Set(ByVal Value As String)
      m_OpenPath = Value
    End Set
  End Property

  Public Property SavePath() As String
    Get
      Return (m_SavePath)
    End Get
    Set(ByVal Value As String)
      m_SavePath = Value
    End Set
  End Property

  Public Property FileName() As String
    Get
      Return m_FileName
    End Get
    Set(ByVal Value As String)
      m_FileName = Value
    End Set
  End Property

  Public ReadOnly Property SavedFileName() As String
    Get
      Return m_SavedPathAndFileName
    End Get
  End Property

  Public Property Client() As OL.ContactItem
    Get
      Client = m_Contact
    End Get
    Set(ByVal Value As OL.ContactItem)
      m_Contact = Value
    End Set
  End Property

#End Region


End Class
