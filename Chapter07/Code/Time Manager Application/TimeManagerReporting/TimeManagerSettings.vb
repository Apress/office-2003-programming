Public Class TimeManagerSettings
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
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents tabPageUsers As System.Windows.Forms.TabPage
    Friend WithEvents tabPageProjects As System.Windows.Forms.TabPage
    Friend WithEvents listProjects As System.Windows.Forms.ListBox
    Friend WithEvents btnAddProject As System.Windows.Forms.Button
    Friend WithEvents btnAddUser As System.Windows.Forms.Button
    Friend WithEvents listUsers As System.Windows.Forms.ListBox
    Friend WithEvents btnRenameProject As System.Windows.Forms.Button
    Friend WithEvents btnDeleteProject As System.Windows.Forms.Button
    Friend WithEvents btnDeleteUser As System.Windows.Forms.Button
    Friend WithEvents btnEditUser As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(TimeManagerSettings))
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tabPageProjects = New System.Windows.Forms.TabPage
        Me.btnDeleteProject = New System.Windows.Forms.Button
        Me.btnRenameProject = New System.Windows.Forms.Button
        Me.btnAddProject = New System.Windows.Forms.Button
        Me.listProjects = New System.Windows.Forms.ListBox
        Me.tabPageUsers = New System.Windows.Forms.TabPage
        Me.btnDeleteUser = New System.Windows.Forms.Button
        Me.btnEditUser = New System.Windows.Forms.Button
        Me.btnAddUser = New System.Windows.Forms.Button
        Me.listUsers = New System.Windows.Forms.ListBox
        Me.cmdClose = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.tabPageProjects.SuspendLayout()
        Me.tabPageUsers.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 56)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(80, 288)
        Me.PictureBox2.TabIndex = 9
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(402, 48)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabPageProjects)
        Me.TabControl1.Controls.Add(Me.tabPageUsers)
        Me.TabControl1.Location = New System.Drawing.Point(96, 56)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(296, 248)
        Me.TabControl1.TabIndex = 11
        '
        'tabPageProjects
        '
        Me.tabPageProjects.Controls.Add(Me.btnDeleteProject)
        Me.tabPageProjects.Controls.Add(Me.btnRenameProject)
        Me.tabPageProjects.Controls.Add(Me.btnAddProject)
        Me.tabPageProjects.Controls.Add(Me.listProjects)
        Me.tabPageProjects.Location = New System.Drawing.Point(4, 22)
        Me.tabPageProjects.Name = "tabPageProjects"
        Me.tabPageProjects.Size = New System.Drawing.Size(288, 222)
        Me.tabPageProjects.TabIndex = 1
        Me.tabPageProjects.Text = "Projects"
        '
        'btnDeleteProject
        '
        Me.btnDeleteProject.Location = New System.Drawing.Point(208, 192)
        Me.btnDeleteProject.Name = "btnDeleteProject"
        Me.btnDeleteProject.Size = New System.Drawing.Size(72, 23)
        Me.btnDeleteProject.TabIndex = 4
        Me.btnDeleteProject.Text = "Delete"
        '
        'btnRenameProject
        '
        Me.btnRenameProject.Location = New System.Drawing.Point(128, 192)
        Me.btnRenameProject.Name = "btnRenameProject"
        Me.btnRenameProject.Size = New System.Drawing.Size(72, 23)
        Me.btnRenameProject.TabIndex = 3
        Me.btnRenameProject.Text = "Rename"
        '
        'btnAddProject
        '
        Me.btnAddProject.Location = New System.Drawing.Point(8, 192)
        Me.btnAddProject.Name = "btnAddProject"
        Me.btnAddProject.Size = New System.Drawing.Size(112, 23)
        Me.btnAddProject.TabIndex = 2
        Me.btnAddProject.Text = "Add Project"
        '
        'listProjects
        '
        Me.listProjects.Location = New System.Drawing.Point(8, 8)
        Me.listProjects.Name = "listProjects"
        Me.listProjects.Size = New System.Drawing.Size(272, 173)
        Me.listProjects.TabIndex = 0
        '
        'tabPageUsers
        '
        Me.tabPageUsers.Controls.Add(Me.btnDeleteUser)
        Me.tabPageUsers.Controls.Add(Me.btnEditUser)
        Me.tabPageUsers.Controls.Add(Me.btnAddUser)
        Me.tabPageUsers.Controls.Add(Me.listUsers)
        Me.tabPageUsers.Location = New System.Drawing.Point(4, 22)
        Me.tabPageUsers.Name = "tabPageUsers"
        Me.tabPageUsers.Size = New System.Drawing.Size(288, 222)
        Me.tabPageUsers.TabIndex = 0
        Me.tabPageUsers.Text = "Users"
        '
        'btnDeleteUser
        '
        Me.btnDeleteUser.Location = New System.Drawing.Point(200, 192)
        Me.btnDeleteUser.Name = "btnDeleteUser"
        Me.btnDeleteUser.Size = New System.Drawing.Size(80, 23)
        Me.btnDeleteUser.TabIndex = 7
        Me.btnDeleteUser.Text = "Delete"
        '
        'btnEditUser
        '
        Me.btnEditUser.Location = New System.Drawing.Point(104, 192)
        Me.btnEditUser.Name = "btnEditUser"
        Me.btnEditUser.Size = New System.Drawing.Size(80, 23)
        Me.btnEditUser.TabIndex = 6
        Me.btnEditUser.Text = "Edit"
        '
        'btnAddUser
        '
        Me.btnAddUser.Location = New System.Drawing.Point(8, 192)
        Me.btnAddUser.Name = "btnAddUser"
        Me.btnAddUser.Size = New System.Drawing.Size(80, 23)
        Me.btnAddUser.TabIndex = 5
        Me.btnAddUser.Text = "Add User"
        '
        'listUsers
        '
        Me.listUsers.Location = New System.Drawing.Point(8, 8)
        Me.listUsers.Name = "listUsers"
        Me.listUsers.Size = New System.Drawing.Size(272, 173)
        Me.listUsers.TabIndex = 3
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(96, 312)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(296, 32)
        Me.cmdClose.TabIndex = 12
        Me.cmdClose.Text = "&Close Settings Window"
        '
        'TimeManagerSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(402, 352)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TimeManagerSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "TimeManagerSettings"
        Me.TabControl1.ResumeLayout(False)
        Me.tabPageProjects.ResumeLayout(False)
        Me.tabPageUsers.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim TimeManager As New TMW.TimeManagerWeb


  Private Sub PopulateProjectList()
    Dim projectArray As TMW.Project() = TimeManager.GetAllProjects()

    Me.listProjects.Items.Clear()
    For Each projectObj As TMW.Project In projectArray
      Me.listProjects.Items.Add(projectObj)
    Next

  End Sub
  Private Sub PopulateUserList()
    Dim userArray As TMW.tmUser() = TimeManager.GetAllUsers()

    Me.listUsers.Items.Clear()
    For Each userObj As TMW.tmUser In userArray
      Me.listUsers.Items.Add(New UserWrapper(userObj))
    Next

  End Sub

  Private Sub TimeManagerSettings_Load(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load

    PopulateProjectList()
    PopulateUserList()
  End Sub

  Private Sub btnAddProject_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnAddProject.Click

    Dim newProjectName As String = InputBox("Please enter the name of the " & _
      "new project", "New Project Title")

    If newProjectName = "" Then
      MsgBox("You must specify a project name before you can add it.", _
        MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Invalid Entry")
      Exit Sub
    End If

    Dim projectObj As New TMW.Project
    projectObj.ProjectName = newProjectName
    If TimeManager.AddProject(projectObj) Then
      PopulateProjectList()
    Else
      MsgBox("The project could not be added.  Please make sure the project " & _
        "name does not already exist.", MsgBoxStyle.Exclamation, "Error Adding Project")
    End If

  End Sub

  Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles cmdClose.Click

    Me.Close()
  End Sub

  Private Sub btnAddUser_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btnAddUser.Click

    Dim UserEntryForm As New frmUser
    Dim UserObj As New TMW.tmUser

    If UserEntryForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
      UserObj.userId = UserEntryForm.txtUsername.Text
      UserObj.nameLast = UserEntryForm.txtNameLast.Text
      UserObj.nameFirst = UserEntryForm.txtNameFirst.Text
      UserObj.password = UserEntryForm.txtPassword.Text
      UserObj.admin = UserEntryForm.chkAdmin.Checked

      If TimeManager.AddUser(UserObj) Then
        PopulateUserList()
      Else
        MsgBox("There was an error trying to add the user.  Make sure this users does " & _
          "not already exist.", MsgBoxStyle.Exclamation, "Error Adding User")
      End If

    End If

  End Sub

  Private Sub btnDeleteUser_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btnDeleteUser.Click

    Dim UserWrapperObj As UserWrapper
    Dim UserObj As TMW.tmUser

    If Me.listUsers.SelectedItem Is Nothing Then
      MsgBox("You must select a user to delete", MsgBoxStyle.Exclamation, "Invalid Selection")
      Exit Sub
    End If

    UserWrapperObj = Me.listUsers.SelectedItem
    UserObj = UserWrapperObj.UserObj

    If MsgBox("Are you sure you want to delete " & UserObj.nameFirst & " " & _
      UserObj.nameLast & "?  This action cannot be undone.", MsgBoxStyle.YesNo Or _
      MsgBoxStyle.Question, "Confirm Delete") = MsgBoxResult.Yes Then
      If TimeManager.DeleteUser(UserObj) Then
        PopulateUserList()
      Else
        MsgBox("Could not delete the user.", MsgBoxStyle.Exclamation, "Error")
      End If
    End If

  End Sub

  Private Sub btnEditUser_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btnEditUser.Click

    If Me.listUsers.SelectedItem Is Nothing Then
      MsgBox("You must select a user to edit", MsgBoxStyle.Exclamation, "Invalid Selection")
      Exit Sub
    End If

    Dim Obj As TMW.tmUser = DirectCast(Me.listUsers.SelectedItem, UserWrapper).UserObj
    Dim UserEntryForm As New frmUser
    Dim OriginalUserId As String

    UserEntryForm.txtUsername.Text = Obj.userId
    UserEntryForm.txtNameLast.Text = Obj.nameLast
    UserEntryForm.txtNameFirst.Text = Obj.nameFirst
    UserEntryForm.txtPassword.Text = Obj.password
    UserEntryForm.txtConfirm.Text = Obj.password
    UserEntryForm.chkAdmin.Checked = Obj.admin
    OriginalUserId = Obj.userId

    If UserEntryForm.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
      Obj.userId = UserEntryForm.txtUsername.Text
      Obj.nameLast = UserEntryForm.txtNameLast.Text
      Obj.nameFirst = UserEntryForm.txtNameFirst.Text
      Obj.password = UserEntryForm.txtPassword.Text
      Obj.admin = UserEntryForm.chkAdmin.Checked

      If TimeManager.UpdateUser(OriginalUserId, Obj) Then
        PopulateUserList()
      Else
        MsgBox("There was an error trying to add the user.  Make sure this users does " & _
          "not already exist.", MsgBoxStyle.Exclamation, "Error Adding User")
      End If
    End If

  End Sub

  Private Sub btnDeleteProject_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btnDeleteProject.Click
    If Me.listProjects.SelectedItem Is Nothing Then
      MsgBox("You must select a project to delete", MsgBoxStyle.Exclamation, _
        "Invalid Selection")
      Exit Sub
    End If

    Dim Obj As TMW.Project = DirectCast(Me.listProjects.SelectedItem, _
      ProjectWrapper).ProjectObj
    If MsgBox("Are you sure you want to delete the project named """ & _
      Obj.ProjectName & """?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, _
        "Confirm Delete") = MsgBoxResult.Yes Then
      If TimeManager.DeleteProject(Obj) Then
        PopulateProjectList()
      Else
        MsgBox("There was an error trying to delete the project", _
          MsgBoxStyle.Exclamation, "Error")
      End If
    End If
  End Sub

  Private Sub btnRenameProject_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btnRenameProject.Click
    If Me.listProjects.SelectedItem Is Nothing Then
      MsgBox("You must select a project to rename", MsgBoxStyle.Exclamation, _
        "Invalid Selection")
      Exit Sub
    End If

    Dim Obj As TMW.Project = DirectCast(Me.listProjects.SelectedItem, _
      ProjectWrapper).ProjectObj
    Dim OriginalProjectName As String = Obj.ProjectName
    Obj.ProjectName = InputBox("Please enter the new name of the project", _
      "Rename Project")

    If Obj.ProjectName = "" Then
      MsgBox("You must enter a name if you want to rename the project", _
        MsgBoxStyle.Exclamation, "Invalid Entry")
      Exit Sub
    End If

    If TimeManager.UpdateProject(OriginalProjectName, Obj) Then
      PopulateProjectList()
    Else
      MsgBox("There was an error trying to rename the project", _
        MsgBoxStyle.Exclamation, "Error")
    End If

  End Sub

End Class