Public Class Settings
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
  Friend WithEvents txtTemplatesFolder As System.Windows.Forms.TextBox
  Friend WithEvents cbSelectFolder As System.Windows.Forms.Button
  Friend WithEvents cbOK As System.Windows.Forms.Button
  Friend WithEvents cbCancel As System.Windows.Forms.Button
  Friend WithEvents gbFolders As System.Windows.Forms.GroupBox
  Friend WithEvents gbUserInfo As System.Windows.Forms.GroupBox
  Friend WithEvents lblItem As System.Windows.Forms.Label
  Friend WithEvents txtContactItem As System.Windows.Forms.TextBox
  Friend WithEvents lblFolder As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Settings))
    Me.gbFolders = New System.Windows.Forms.GroupBox
    Me.cbSelectFolder = New System.Windows.Forms.Button
    Me.lblFolder = New System.Windows.Forms.Label
    Me.txtTemplatesFolder = New System.Windows.Forms.TextBox
    Me.cbOK = New System.Windows.Forms.Button
    Me.cbCancel = New System.Windows.Forms.Button
    Me.gbUserInfo = New System.Windows.Forms.GroupBox
    Me.lblItem = New System.Windows.Forms.Label
    Me.txtContactItem = New System.Windows.Forms.TextBox
    Me.gbFolders.SuspendLayout()
    Me.gbUserInfo.SuspendLayout()
    Me.SuspendLayout()
    '
    'gbFolders
    '
    Me.gbFolders.Controls.Add(Me.cbSelectFolder)
    Me.gbFolders.Controls.Add(Me.lblFolder)
    Me.gbFolders.Controls.Add(Me.txtTemplatesFolder)
    Me.gbFolders.Location = New System.Drawing.Point(8, 32)
    Me.gbFolders.Name = "gbFolders"
    Me.gbFolders.Size = New System.Drawing.Size(400, 80)
    Me.gbFolders.TabIndex = 0
    Me.gbFolders.TabStop = False
    Me.gbFolders.Text = "Templates Folder"
    '
    'cbSelectFolder
    '
    Me.cbSelectFolder.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cbSelectFolder.Location = New System.Drawing.Point(352, 32)
    Me.cbSelectFolder.Name = "cbSelectFolder"
    Me.cbSelectFolder.Size = New System.Drawing.Size(40, 23)
    Me.cbSelectFolder.TabIndex = 2
    Me.cbSelectFolder.Text = "&..."
    '
    'lblFolder
    '
    Me.lblFolder.Location = New System.Drawing.Point(8, 32)
    Me.lblFolder.Name = "lblFolder"
    Me.lblFolder.Size = New System.Drawing.Size(72, 24)
    Me.lblFolder.TabIndex = 1
    Me.lblFolder.Text = "Folder ID"
    '
    'txtTemplatesFolder
    '
    Me.txtTemplatesFolder.Location = New System.Drawing.Point(104, 32)
    Me.txtTemplatesFolder.Name = "txtTemplatesFolder"
    Me.txtTemplatesFolder.Size = New System.Drawing.Size(240, 20)
    Me.txtTemplatesFolder.TabIndex = 0
    Me.txtTemplatesFolder.Text = ""
    '
    'cbOK
    '
    Me.cbOK.Location = New System.Drawing.Point(256, 240)
    Me.cbOK.Name = "cbOK"
    Me.cbOK.TabIndex = 1
    Me.cbOK.Text = "O&K"
    '
    'cbCancel
    '
    Me.cbCancel.Location = New System.Drawing.Point(336, 240)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.TabIndex = 2
    Me.cbCancel.Text = "C&ancel"
    '
    'gbUserInfo
    '
    Me.gbUserInfo.Controls.Add(Me.lblItem)
    Me.gbUserInfo.Controls.Add(Me.txtContactItem)
    Me.gbUserInfo.Location = New System.Drawing.Point(8, 152)
    Me.gbUserInfo.Name = "gbUserInfo"
    Me.gbUserInfo.Size = New System.Drawing.Size(400, 48)
    Me.gbUserInfo.TabIndex = 3
    Me.gbUserInfo.TabStop = False
    Me.gbUserInfo.Text = "Contact Item Storing User Information"
    '
    'lblItem
    '
    Me.lblItem.Location = New System.Drawing.Point(8, 24)
    Me.lblItem.Name = "lblItem"
    Me.lblItem.Size = New System.Drawing.Size(80, 16)
    Me.lblItem.TabIndex = 0
    Me.lblItem.Text = "Name"
    '
    'txtContactItem
    '
    Me.txtContactItem.Location = New System.Drawing.Point(104, 18)
    Me.txtContactItem.Name = "txtContactItem"
    Me.txtContactItem.Size = New System.Drawing.Size(288, 20)
    Me.txtContactItem.TabIndex = 3
    Me.txtContactItem.Text = ""
    '
    'Settings
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(424, 278)
    Me.Controls.Add(Me.gbUserInfo)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.cbOK)
    Me.Controls.Add(Me.gbFolders)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "Settings"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Email Templates Settings"
    Me.gbFolders.ResumeLayout(False)
    Me.gbUserInfo.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region
  Private entryID As String
  Private storeID As String

  Private Sub cbSelectFolder_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbSelectFolder.Click

    Dim str() As String
    str = appOutlook.PickOutlookFolder()
    If str Is Nothing Then
      Return
    Else
      txtTemplatesFolder.Text = str(0)
      entryID = str(1)
      storeID = str(2)
    End If

  End Sub

  Private Sub cbOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbOK.Click
    Try
      SaveSettings()
    Catch ex As Exception
      MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Critical Error")
    Finally
      Me.Close()
    End Try
  End Sub

  Private Function SaveSettings()
    Try
      UserSettings.TemplatesFolderPath = _
        txtTemplatesFolder.Text
      UserSettings.EntryID = entryID
      UserSettings.StoreID = storeID
      UserSettings.UserName = txtContactItem.Text
      UserSettings.SaveSettings()

    Catch ex As Exception
      MsgBox(ex.Message)
    End Try


  End Function

  Private Function LoadSettings()
    Try
      txtTemplatesFolder.Text = UserSettings.TemplatesFolderPath
      entryID = UserSettings.EntryID
      storeID = UserSettings.StoreID
      txtContactItem.Text = UserSettings.UserName
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try
  End Function

  Private Sub Settings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    LoadSettings()
  End Sub

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()
  End Sub


End Class
