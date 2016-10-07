Imports System.IO
Imports System.Windows.Forms

Public Class Documents
  Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call

  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents cbOK As System.Windows.Forms.Button
  Friend WithEvents cbCancel As System.Windows.Forms.Button
  Friend WithEvents lblSelectedDocs As System.Windows.Forms.Label
  Friend WithEvents lblClients As System.Windows.Forms.Label
  Friend WithEvents cboClients As System.Windows.Forms.ComboBox
  Friend WithEvents cbRemoveAll As System.Windows.Forms.Button
  Friend WithEvents cbSelectAll As System.Windows.Forms.Button
  Friend WithEvents cbRemoveOne As System.Windows.Forms.Button
  Friend WithEvents cbSelectOne As System.Windows.Forms.Button
  Friend WithEvents lstSelectedDocs As System.Windows.Forms.ListBox
  Friend WithEvents lstAvailableDocs As System.Windows.Forms.ListBox
  Friend WithEvents lblAvailableDocs As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Documents))
    Me.cbOK = New System.Windows.Forms.Button
    Me.cbCancel = New System.Windows.Forms.Button
    Me.lblSelectedDocs = New System.Windows.Forms.Label
    Me.lblAvailableDocs = New System.Windows.Forms.Label
    Me.lblClients = New System.Windows.Forms.Label
    Me.cboClients = New System.Windows.Forms.ComboBox
    Me.cbRemoveAll = New System.Windows.Forms.Button
    Me.cbSelectAll = New System.Windows.Forms.Button
    Me.cbRemoveOne = New System.Windows.Forms.Button
    Me.cbSelectOne = New System.Windows.Forms.Button
    Me.lstSelectedDocs = New System.Windows.Forms.ListBox
    Me.lstAvailableDocs = New System.Windows.Forms.ListBox
    Me.SuspendLayout()
    '
    'cbOK
    '
    Me.cbOK.Location = New System.Drawing.Point(312, 304)
    Me.cbOK.Name = "cbOK"
    Me.cbOK.TabIndex = 6
    Me.cbOK.Text = "&OK"
    '
    'cbCancel
    '
    Me.cbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cbCancel.Location = New System.Drawing.Point(392, 304)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.TabIndex = 7
    Me.cbCancel.Text = "C&ancel"
    '
    'lblSelectedDocs
    '
    Me.lblSelectedDocs.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblSelectedDocs.Location = New System.Drawing.Point(272, 87)
    Me.lblSelectedDocs.Name = "lblSelectedDocs"
    Me.lblSelectedDocs.Size = New System.Drawing.Size(168, 16)
    Me.lblSelectedDocs.TabIndex = 31
    Me.lblSelectedDocs.Text = "Selected Documents"
    '
    'lblAvailableDocs
    '
    Me.lblAvailableDocs.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblAvailableDocs.Location = New System.Drawing.Point(32, 87)
    Me.lblAvailableDocs.Name = "lblAvailableDocs"
    Me.lblAvailableDocs.Size = New System.Drawing.Size(168, 16)
    Me.lblAvailableDocs.TabIndex = 30
    Me.lblAvailableDocs.Text = "Available Documents"
    '
    'lblClients
    '
    Me.lblClients.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblClients.Location = New System.Drawing.Point(32, 55)
    Me.lblClients.Name = "lblClients"
    Me.lblClients.Size = New System.Drawing.Size(120, 23)
    Me.lblClients.TabIndex = 29
    Me.lblClients.Text = "Select a Client"
    '
    'cboClients
    '
    Me.cboClients.Location = New System.Drawing.Point(160, 55)
    Me.cboClients.Name = "cboClients"
    Me.cboClients.Size = New System.Drawing.Size(288, 21)
    Me.cboClients.TabIndex = 28
    '
    'cbRemoveAll
    '
    Me.cbRemoveAll.Location = New System.Drawing.Point(216, 247)
    Me.cbRemoveAll.Name = "cbRemoveAll"
    Me.cbRemoveAll.Size = New System.Drawing.Size(48, 24)
    Me.cbRemoveAll.TabIndex = 27
    Me.cbRemoveAll.Text = "<<"
    '
    'cbSelectAll
    '
    Me.cbSelectAll.Location = New System.Drawing.Point(216, 223)
    Me.cbSelectAll.Name = "cbSelectAll"
    Me.cbSelectAll.Size = New System.Drawing.Size(48, 24)
    Me.cbSelectAll.TabIndex = 26
    Me.cbSelectAll.Text = ">>"
    '
    'cbRemoveOne
    '
    Me.cbRemoveOne.Location = New System.Drawing.Point(216, 151)
    Me.cbRemoveOne.Name = "cbRemoveOne"
    Me.cbRemoveOne.Size = New System.Drawing.Size(48, 24)
    Me.cbRemoveOne.TabIndex = 25
    Me.cbRemoveOne.Text = "<"
    '
    'cbSelectOne
    '
    Me.cbSelectOne.Location = New System.Drawing.Point(216, 127)
    Me.cbSelectOne.Name = "cbSelectOne"
    Me.cbSelectOne.Size = New System.Drawing.Size(48, 24)
    Me.cbSelectOne.TabIndex = 24
    Me.cbSelectOne.Text = ">"
    '
    'lstSelectedDocs
    '
    Me.lstSelectedDocs.Location = New System.Drawing.Point(272, 119)
    Me.lstSelectedDocs.Name = "lstSelectedDocs"
    Me.lstSelectedDocs.Size = New System.Drawing.Size(176, 160)
    Me.lstSelectedDocs.TabIndex = 23
    '
    'lstAvailableDocs
    '
    Me.lstAvailableDocs.Location = New System.Drawing.Point(32, 119)
    Me.lstAvailableDocs.Name = "lstAvailableDocs"
    Me.lstAvailableDocs.Size = New System.Drawing.Size(176, 160)
    Me.lstAvailableDocs.TabIndex = 22
    '
    'Documents
    '
    Me.AcceptButton = Me.cbOK
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.CancelButton = Me.cbCancel
    Me.ClientSize = New System.Drawing.Size(480, 334)
    Me.Controls.Add(Me.lblSelectedDocs)
    Me.Controls.Add(Me.lblAvailableDocs)
    Me.Controls.Add(Me.lblClients)
    Me.Controls.Add(Me.cboClients)
    Me.Controls.Add(Me.cbRemoveAll)
    Me.Controls.Add(Me.cbSelectAll)
    Me.Controls.Add(Me.cbRemoveOne)
    Me.Controls.Add(Me.cbSelectOne)
    Me.Controls.Add(Me.lstSelectedDocs)
    Me.Controls.Add(Me.lstAvailableDocs)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.cbOK)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Location = New System.Drawing.Point(20, 50)
    Me.Name = "Documents"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Project Documents"
    Me.ResumeLayout(False)

  End Sub

#End Region

#Region "Methods"

  Private Function GetFiles(ByVal ctl As ListBox)
    Dim diFolder As New DirectoryInfo(UserSettings.TemplatesFolder)
    Dim fi As FileInfo

    For Each fi In diFolder.GetFiles("*.dot")
      ctl.Sorted = True
      ctl.Items.Add(fi.Name)
    Next
  End Function

  Private Function MoveSelected(ByVal SourceList As ListBox, _
    ByVal TargetList As ListBox) As Boolean

    Try
      TargetList.Items.Add(SourceList.SelectedItem)
      SourceList.Items.Remove(SourceList.SelectedItem)
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  Private Function MoveAll(ByVal SourceList As ListBox, _
    ByVal TargetList As ListBox) As Boolean

    Try
      Dim si


      For Each si In SourceList.Items
        TargetList.Items.Add(si)

      Next
      SourceList.Items.Clear()
      Return True

    Catch ex As Exception
      Return False
    End Try

  End Function

  Private Sub GenerateDocs()

    Dim i As Integer
    Dim strClient As String
    Dim strFile As String
    Dim strFiles(lstSelectedDocs.Items.Count) As String

    strClient = cboClients.SelectedItem
    Try
      For i = 0 To lstSelectedDocs.Items.Count - 1
        strFile = lstSelectedDocs.Items(i)
        Dim dp As New DocumentProcessor

        With dp
          .Client = appOutlook.GetContact(strClient)
          .FileName = strFile.ToString
          .OpenPath = UserSettings.TemplatesFolder
          .SavePath = UserSettings.SaveFolder
          .OpenDoc()
          .PopulateForm()
          .ShowDialog()
          'Populate String Array.  These will be inserted as
          'Attachements to an email.
          If .DocSaved Then
            strFiles(i) = .SavePath & "\" & .SavedFileName
          End If
          .Visible = False
          .Dispose()

        End With

      Next
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try

    appOutlook.CreateEmail(strFiles, strClient)
  End Sub

#End Region

#Region "Events"

  Private Sub Documents_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load

    GetFiles(lstAvailableDocs)
    appOutlook.ListContacts(cboClients)

  End Sub

  Private Sub cbSelectOne_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbSelectOne.Click

    MoveSelected(lstAvailableDocs, lstSelectedDocs)

  End Sub

  Private Sub cbRemoveOne_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbRemoveOne.Click

    MoveSelected(lstSelectedDocs, lstAvailableDocs)

  End Sub

  Private Sub cbSelectAll_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbSelectAll.Click

    MoveAll(lstAvailableDocs, lstSelectedDocs)

  End Sub

  Private Sub cbRemoveAll_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbRemoveAll.Click

    MoveAll(lstSelectedDocs, lstAvailableDocs)

  End Sub

  Private Sub cbOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbOK.Click

    GenerateDocs()
    Me.Close()

  End Sub

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()

  End Sub

  Private Sub cboClients_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cboClients.Click

    cboClients.DroppedDown = Not cboClients.DroppedDown

  End Sub

#End Region
End Class
