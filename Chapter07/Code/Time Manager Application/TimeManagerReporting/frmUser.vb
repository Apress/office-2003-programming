Public Class frmUser
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
    Friend WithEvents lblUsername As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtUsername As System.Windows.Forms.TextBox
    Friend WithEvents txtNameLast As System.Windows.Forms.TextBox
    Friend WithEvents txtNameFirst As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtConfirm As System.Windows.Forms.TextBox
    Friend WithEvents chkAdmin As System.Windows.Forms.CheckBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnAddUpdate As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnAddUpdate = New System.Windows.Forms.Button
        Me.txtUsername = New System.Windows.Forms.TextBox
        Me.lblUsername = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtNameLast = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtNameFirst = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtConfirm = New System.Windows.Forms.TextBox
        Me.chkAdmin = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(160, 240)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 0
        Me.btnCancel.Text = "&Cancel"
        '
        'btnAddUpdate
        '
        Me.btnAddUpdate.Location = New System.Drawing.Point(8, 240)
        Me.btnAddUpdate.Name = "btnAddUpdate"
        Me.btnAddUpdate.Size = New System.Drawing.Size(144, 23)
        Me.btnAddUpdate.TabIndex = 1
        Me.btnAddUpdate.Text = "Add / Update User"
        '
        'txtUsername
        '
        Me.txtUsername.Location = New System.Drawing.Point(8, 24)
        Me.txtUsername.Name = "txtUsername"
        Me.txtUsername.Size = New System.Drawing.Size(240, 20)
        Me.txtUsername.TabIndex = 2
        Me.txtUsername.Text = ""
        '
        'lblUsername
        '
        Me.lblUsername.AutoSize = True
        Me.lblUsername.Location = New System.Drawing.Point(8, 8)
        Me.lblUsername.Name = "lblUsername"
        Me.lblUsername.Size = New System.Drawing.Size(56, 16)
        Me.lblUsername.TabIndex = 3
        Me.lblUsername.Text = "Username"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Last Name"
        '
        'txtNameLast
        '
        Me.txtNameLast.Location = New System.Drawing.Point(8, 64)
        Me.txtNameLast.Name = "txtNameLast"
        Me.txtNameLast.Size = New System.Drawing.Size(240, 20)
        Me.txtNameLast.TabIndex = 4
        Me.txtNameLast.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "First Name"
        '
        'txtNameFirst
        '
        Me.txtNameFirst.Location = New System.Drawing.Point(8, 104)
        Me.txtNameFirst.Name = "txtNameFirst"
        Me.txtNameFirst.Size = New System.Drawing.Size(240, 20)
        Me.txtNameFirst.TabIndex = 6
        Me.txtNameFirst.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 16)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Password"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(8, 144)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(240, 20)
        Me.txtPassword.TabIndex = 8
        Me.txtPassword.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 168)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Confirm"
        '
        'txtConfirm
        '
        Me.txtConfirm.Location = New System.Drawing.Point(8, 184)
        Me.txtConfirm.Name = "txtConfirm"
        Me.txtConfirm.Size = New System.Drawing.Size(240, 20)
        Me.txtConfirm.TabIndex = 10
        Me.txtConfirm.Text = ""
        '
        'chkAdmin
        '
        Me.chkAdmin.Location = New System.Drawing.Point(8, 208)
        Me.chkAdmin.Name = "chkAdmin"
        Me.chkAdmin.TabIndex = 12
        Me.chkAdmin.Text = "Administrator"
        '
        'frmUser
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(258, 272)
        Me.Controls.Add(Me.chkAdmin)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtConfirm)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtNameFirst)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNameLast)
        Me.Controls.Add(Me.lblUsername)
        Me.Controls.Add(Me.txtUsername)
        Me.Controls.Add(Me.btnAddUpdate)
        Me.Controls.Add(Me.btnCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUser"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Time Manager User Info"
        Me.ResumeLayout(False)

    End Sub

#End Region

  Private Function HasText(ByRef txtBox As Windows.Forms.TextBox, _
    ByVal ErrorMsg As String) As Boolean

    If txtBox.Text = "" Then
      MsgBox(ErrorMsg, MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation, _
        "Invalid Entry")
      txtBox.Focus()
      Return False
    Else
      Return True
    End If
  End Function

  Private Sub btnAddUpdate_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles btnAddUpdate.Click

    If HasText(Me.txtUsername, "You must specify a username") AndAlso _
       HasText(Me.txtNameLast, "You must specify a last name") AndAlso _
       HasText(Me.txtNameFirst, "You must specify a first name") AndAlso _
       HasText(Me.txtPassword, "You must specify a password") AndAlso _
       HasText(Me.txtConfirm, "You must confirm the password") Then
      If Not Me.txtPassword.Text = Me.txtConfirm.Text Then
        MsgBox("Your passwords do not match.", MsgBoxStyle.Exclamation, _
          "Invalid Passwords")
        Me.txtPassword.Focus()
        Return
      Else
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
      End If
    End If
  End Sub

 
End Class
