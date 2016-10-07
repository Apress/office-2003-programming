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
  Friend WithEvents cbCancel As System.Windows.Forms.Button
  Friend WithEvents cbOK As System.Windows.Forms.Button
  Friend WithEvents cbBrowseTemplates As System.Windows.Forms.Button
  Friend WithEvents gbFolders As System.Windows.Forms.GroupBox
  Friend WithEvents gbSQLSettings As System.Windows.Forms.GroupBox
  Friend WithEvents txtServer As System.Windows.Forms.TextBox
  Friend WithEvents txtUserID As System.Windows.Forms.TextBox
  Friend WithEvents txtPassword As System.Windows.Forms.TextBox
  Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
  Friend WithEvents txtSave As System.Windows.Forms.TextBox
  Friend WithEvents txtTemplates As System.Windows.Forms.TextBox
  Friend WithEvents cbBrowseSave As System.Windows.Forms.Button
  Friend WithEvents lblSaveFolder As System.Windows.Forms.Label
  Friend WithEvents lblTemplatesFolder As System.Windows.Forms.Label
  Friend WithEvents lblUserName As System.Windows.Forms.Label
  Friend WithEvents lblServerName As System.Windows.Forms.Label
  Friend WithEvents lblDatabase As System.Windows.Forms.Label
  Friend WithEvents lblPassword As System.Windows.Forms.Label
  Friend WithEvents gbOutlook As System.Windows.Forms.GroupBox
  Friend WithEvents cbPickFolder As System.Windows.Forms.Button
  Friend WithEvents txtClientsFolder As System.Windows.Forms.TextBox
  Friend WithEvents lblFolderID As System.Windows.Forms.Label
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.cbCancel = New System.Windows.Forms.Button
    Me.cbOK = New System.Windows.Forms.Button
    Me.gbFolders = New System.Windows.Forms.GroupBox
    Me.lblSaveFolder = New System.Windows.Forms.Label
    Me.lblTemplatesFolder = New System.Windows.Forms.Label
    Me.cbBrowseSave = New System.Windows.Forms.Button
    Me.txtSave = New System.Windows.Forms.TextBox
    Me.cbBrowseTemplates = New System.Windows.Forms.Button
    Me.txtTemplates = New System.Windows.Forms.TextBox
    Me.gbSQLSettings = New System.Windows.Forms.GroupBox
    Me.lblDatabase = New System.Windows.Forms.Label
    Me.lblPassword = New System.Windows.Forms.Label
    Me.lblUserName = New System.Windows.Forms.Label
    Me.txtDatabase = New System.Windows.Forms.TextBox
    Me.txtPassword = New System.Windows.Forms.TextBox
    Me.txtUserID = New System.Windows.Forms.TextBox
    Me.lblServerName = New System.Windows.Forms.Label
    Me.txtServer = New System.Windows.Forms.TextBox
    Me.gbOutlook = New System.Windows.Forms.GroupBox
    Me.cbPickFolder = New System.Windows.Forms.Button
    Me.txtClientsFolder = New System.Windows.Forms.TextBox
    Me.lblFolderID = New System.Windows.Forms.Label
    Me.gbFolders.SuspendLayout()
    Me.gbSQLSettings.SuspendLayout()
    Me.gbOutlook.SuspendLayout()
    Me.SuspendLayout()
    '
    'cbCancel
    '
    Me.cbCancel.Location = New System.Drawing.Point(356, 324)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.TabIndex = 3
    Me.cbCancel.Text = "C&ancel"
    '
    'cbOK
    '
    Me.cbOK.Location = New System.Drawing.Point(276, 324)
    Me.cbOK.Name = "cbOK"
    Me.cbOK.TabIndex = 2
    Me.cbOK.Text = "O&K"
    '
    'gbFolders
    '
    Me.gbFolders.Controls.Add(Me.lblSaveFolder)
    Me.gbFolders.Controls.Add(Me.lblTemplatesFolder)
    Me.gbFolders.Controls.Add(Me.cbBrowseSave)
    Me.gbFolders.Controls.Add(Me.txtSave)
    Me.gbFolders.Controls.Add(Me.cbBrowseTemplates)
    Me.gbFolders.Controls.Add(Me.txtTemplates)
    Me.gbFolders.Location = New System.Drawing.Point(8, 40)
    Me.gbFolders.Name = "gbFolders"
    Me.gbFolders.Size = New System.Drawing.Size(424, 88)
    Me.gbFolders.TabIndex = 0
    Me.gbFolders.TabStop = False
    Me.gbFolders.Text = "Default Folder Settings"
    '
    'lblSaveFolder
    '
    Me.lblSaveFolder.Location = New System.Drawing.Point(16, 55)
    Me.lblSaveFolder.Name = "lblSaveFolder"
    Me.lblSaveFolder.Size = New System.Drawing.Size(112, 23)
    Me.lblSaveFolder.TabIndex = 15
    Me.lblSaveFolder.Text = "Default Save Folder"
    '
    'lblTemplatesFolder
    '
    Me.lblTemplatesFolder.Location = New System.Drawing.Point(16, 26)
    Me.lblTemplatesFolder.Name = "lblTemplatesFolder"
    Me.lblTemplatesFolder.TabIndex = 14
    Me.lblTemplatesFolder.Text = "Templates Folder"
    '
    'cbBrowseSave
    '
    Me.cbBrowseSave.Location = New System.Drawing.Point(391, 50)
    Me.cbBrowseSave.Name = "cbBrowseSave"
    Me.cbBrowseSave.Size = New System.Drawing.Size(25, 23)
    Me.cbBrowseSave.TabIndex = 3
    Me.cbBrowseSave.Text = "..."
    '
    'txtSave
    '
    Me.txtSave.Location = New System.Drawing.Point(135, 50)
    Me.txtSave.Name = "txtSave"
    Me.txtSave.Size = New System.Drawing.Size(248, 20)
    Me.txtSave.TabIndex = 2
    Me.txtSave.Text = ""
    '
    'cbBrowseTemplates
    '
    Me.cbBrowseTemplates.Location = New System.Drawing.Point(391, 23)
    Me.cbBrowseTemplates.Name = "cbBrowseTemplates"
    Me.cbBrowseTemplates.Size = New System.Drawing.Size(25, 23)
    Me.cbBrowseTemplates.TabIndex = 1
    Me.cbBrowseTemplates.Text = "..."
    '
    'txtTemplates
    '
    Me.txtTemplates.Location = New System.Drawing.Point(135, 23)
    Me.txtTemplates.Name = "txtTemplates"
    Me.txtTemplates.Size = New System.Drawing.Size(248, 20)
    Me.txtTemplates.TabIndex = 0
    Me.txtTemplates.Text = ""
    '
    'gbSQLSettings
    '
    Me.gbSQLSettings.Controls.Add(Me.lblDatabase)
    Me.gbSQLSettings.Controls.Add(Me.lblPassword)
    Me.gbSQLSettings.Controls.Add(Me.lblUserName)
    Me.gbSQLSettings.Controls.Add(Me.txtDatabase)
    Me.gbSQLSettings.Controls.Add(Me.txtPassword)
    Me.gbSQLSettings.Controls.Add(Me.txtUserID)
    Me.gbSQLSettings.Controls.Add(Me.lblServerName)
    Me.gbSQLSettings.Controls.Add(Me.txtServer)
    Me.gbSQLSettings.Location = New System.Drawing.Point(8, 188)
    Me.gbSQLSettings.Name = "gbSQLSettings"
    Me.gbSQLSettings.Size = New System.Drawing.Size(424, 128)
    Me.gbSQLSettings.TabIndex = 1
    Me.gbSQLSettings.TabStop = False
    Me.gbSQLSettings.Text = "SQL Server Settings"
    '
    'lblDatabase
    '
    Me.lblDatabase.Location = New System.Drawing.Point(16, 96)
    Me.lblDatabase.Name = "lblDatabase"
    Me.lblDatabase.Size = New System.Drawing.Size(112, 23)
    Me.lblDatabase.TabIndex = 7
    Me.lblDatabase.Text = "Database"
    '
    'lblPassword
    '
    Me.lblPassword.Location = New System.Drawing.Point(16, 72)
    Me.lblPassword.Name = "lblPassword"
    Me.lblPassword.Size = New System.Drawing.Size(112, 23)
    Me.lblPassword.TabIndex = 6
    Me.lblPassword.Text = "Password"
    '
    'lblUserName
    '
    Me.lblUserName.Location = New System.Drawing.Point(16, 48)
    Me.lblUserName.Name = "lblUserName"
    Me.lblUserName.Size = New System.Drawing.Size(112, 23)
    Me.lblUserName.TabIndex = 5
    Me.lblUserName.Text = "User Name"
    '
    'txtDatabase
    '
    Me.txtDatabase.Location = New System.Drawing.Point(136, 96)
    Me.txtDatabase.Name = "txtDatabase"
    Me.txtDatabase.Size = New System.Drawing.Size(248, 20)
    Me.txtDatabase.TabIndex = 3
    Me.txtDatabase.Text = ""
    '
    'txtPassword
    '
    Me.txtPassword.Font = New System.Drawing.Font("Wingdings", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
    Me.txtPassword.Location = New System.Drawing.Point(136, 72)
    Me.txtPassword.Name = "txtPassword"
    Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(108)
    Me.txtPassword.Size = New System.Drawing.Size(248, 20)
    Me.txtPassword.TabIndex = 2
    Me.txtPassword.Text = ""
    '
    'txtUserID
    '
    Me.txtUserID.Location = New System.Drawing.Point(136, 48)
    Me.txtUserID.Name = "txtUserID"
    Me.txtUserID.Size = New System.Drawing.Size(248, 20)
    Me.txtUserID.TabIndex = 1
    Me.txtUserID.Text = ""
    '
    'lblServerName
    '
    Me.lblServerName.Location = New System.Drawing.Point(16, 24)
    Me.lblServerName.Name = "lblServerName"
    Me.lblServerName.Size = New System.Drawing.Size(112, 23)
    Me.lblServerName.TabIndex = 1
    Me.lblServerName.Text = "Server  Name/IP"
    '
    'txtServer
    '
    Me.txtServer.Location = New System.Drawing.Point(136, 24)
    Me.txtServer.Name = "txtServer"
    Me.txtServer.Size = New System.Drawing.Size(248, 20)
    Me.txtServer.TabIndex = 0
    Me.txtServer.Text = ""
    '
    'gbOutlook
    '
    Me.gbOutlook.Controls.Add(Me.cbPickFolder)
    Me.gbOutlook.Controls.Add(Me.txtClientsFolder)
    Me.gbOutlook.Controls.Add(Me.lblFolderID)
    Me.gbOutlook.Location = New System.Drawing.Point(8, 136)
    Me.gbOutlook.Name = "gbOutlook"
    Me.gbOutlook.Size = New System.Drawing.Size(424, 40)
    Me.gbOutlook.TabIndex = 4
    Me.gbOutlook.TabStop = False
    Me.gbOutlook.Text = "Bravo Client Folder"
    '
    'cbPickFolder
    '
    Me.cbPickFolder.Location = New System.Drawing.Point(391, 11)
    Me.cbPickFolder.Name = "cbPickFolder"
    Me.cbPickFolder.Size = New System.Drawing.Size(25, 23)
    Me.cbPickFolder.TabIndex = 16
    Me.cbPickFolder.Text = "..."
    '
    'txtClientsFolder
    '
    Me.txtClientsFolder.Location = New System.Drawing.Point(135, 13)
    Me.txtClientsFolder.Name = "txtClientsFolder"
    Me.txtClientsFolder.Size = New System.Drawing.Size(248, 20)
    Me.txtClientsFolder.TabIndex = 16
    Me.txtClientsFolder.Text = ""
    '
    'lblFolderID
    '
    Me.lblFolderID.Location = New System.Drawing.Point(16, 16)
    Me.lblFolderID.Name = "lblFolderID"
    Me.lblFolderID.Size = New System.Drawing.Size(112, 16)
    Me.lblFolderID.TabIndex = 16
    Me.lblFolderID.Text = "Folder ID"
    '
    'Settings
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(440, 366)
    Me.Controls.Add(Me.gbOutlook)
    Me.Controls.Add(Me.gbSQLSettings)
    Me.Controls.Add(Me.gbFolders)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.cbOK)
    Me.Name = "Settings"
    Me.Text = "Document Generator Settings"
    Me.gbFolders.ResumeLayout(False)
    Me.gbSQLSettings.ResumeLayout(False)
    Me.gbOutlook.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region
  Private entryID As String
  Private storeID As String

  Private Sub Settings_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load
    LoadSettings()
  End Sub

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()
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
      UserSettings.TemplatesFolder = txtTemplates.Text
      UserSettings.SaveFolder = txtSave.Text
      UserSettings.DatabaseName = txtDatabase.Text
      UserSettings.ServerName = txtServer.Text
      UserSettings.UserName = txtUserID.Text
      UserSettings.Password = txtPassword.Text
      UserSettings.ClientsFolderPath = txtClientsFolder.Text
      UserSettings.EntryID = EntryID
      UserSettings.StoreID = StoreID
      UserSettings.SaveSettings()
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try


  End Function

  Private Function LoadSettings()
    Try
      txtTemplates.Text = UserSettings.TemplatesFolder
      txtSave.Text = UserSettings.SaveFolder
      txtDatabase.Text = UserSettings.DatabaseName
      txtServer.Text = UserSettings.ServerName
      txtUserID.Text = UserSettings.UserName
      txtPassword.Text = UserSettings.Password
      txtClientsFolder.Text = UserSettings.ClientsFolderPath
      entryID = UserSettings.EntryID
      storeID = UserSettings.StoreID
    Catch ex As Exception
      MsgBox(ex.Message)
    End Try
  End Function

  Private Sub cbBrowseTemplates_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbBrowseTemplates.Click
    txtTemplates.Text = BrowseForFolder("Templates Folder")
  End Sub

  Private Sub cbBrowseSave_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbBrowseSave.Click
    txtSave.Text = BrowseForFolder("Save Folder")
  End Sub

  Private Function BrowseForFolder(ByVal FolderType As String) As String
    Dim fdPath As New Windows.Forms.FolderBrowserDialog
    Dim strTitle As String
    Dim strSelectedPath As String

    Try
      strTitle = "Select the "
      strTitle = strTitle.Concat(strTitle, FolderType.ToUpper)
      strTitle = strTitle.Concat(strTitle, " for your Bravo Project Documents")

      With fdPath
        .Description = strTitle
        .ShowNewFolderButton = True

        If .ShowDialog() = DialogResult.OK Then
          strSelectedPath = .SelectedPath
          strSelectedPath = strSelectedPath.Concat(strSelectedPath, "\")
          Return strSelectedPath.ToString
        Else
          Return ""
        End If

      End With


    Catch ex As Exception

    Finally
      fdPath = Nothing
      strSelectedPath = Nothing

    End Try

  End Function

  Private Sub cbPickFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPickFolder.Click
    Dim str() As String
    str = appOutlook.PickContactsFolder()
    txtClientsFolder.Text = str(0)
    EntryID = str(1)
    StoreID = str(2)
  End Sub


End Class
