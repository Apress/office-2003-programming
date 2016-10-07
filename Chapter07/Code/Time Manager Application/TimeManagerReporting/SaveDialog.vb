Public Class SaveDialog
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents dpStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents gboxSaveInfo As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents chkUpload As System.Windows.Forms.CheckBox
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents chkClearContents As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SaveDialog))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblInfo = New System.Windows.Forms.Label
        Me.dpStartDate = New System.Windows.Forms.DateTimePicker
        Me.lblDate = New System.Windows.Forms.Label
        Me.gboxSaveInfo = New System.Windows.Forms.GroupBox
        Me.chkClearContents = New System.Windows.Forms.CheckBox
        Me.chkUpload = New System.Windows.Forms.CheckBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.gboxSaveInfo.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(496, 48)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'lblInfo
        '
        Me.lblInfo.BackColor = System.Drawing.Color.White
        Me.lblInfo.Location = New System.Drawing.Point(16, 8)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(456, 32)
        Me.lblInfo.TabIndex = 1
        Me.lblInfo.Text = "Please indicate below whether or not you would like to upload this information to" & _
        " the Time Management system.  You must specify a ""Monday"" for the week start."
        Me.lblInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dpStartDate
        '
        Me.dpStartDate.Location = New System.Drawing.Point(24, 40)
        Me.dpStartDate.Name = "dpStartDate"
        Me.dpStartDate.Size = New System.Drawing.Size(328, 20)
        Me.dpStartDate.TabIndex = 2
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Location = New System.Drawing.Point(24, 24)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(163, 16)
        Me.lblDate.TabIndex = 3
        Me.lblDate.Text = "Week Start (Must be a Monday)"
        '
        'gboxSaveInfo
        '
        Me.gboxSaveInfo.Controls.Add(Me.chkClearContents)
        Me.gboxSaveInfo.Controls.Add(Me.chkUpload)
        Me.gboxSaveInfo.Controls.Add(Me.dpStartDate)
        Me.gboxSaveInfo.Controls.Add(Me.lblDate)
        Me.gboxSaveInfo.Location = New System.Drawing.Point(96, 56)
        Me.gboxSaveInfo.Name = "gboxSaveInfo"
        Me.gboxSaveInfo.Size = New System.Drawing.Size(392, 96)
        Me.gboxSaveInfo.TabIndex = 4
        Me.gboxSaveInfo.TabStop = False
        '
        'chkClearContents
        '
        Me.chkClearContents.Location = New System.Drawing.Point(24, 64)
        Me.chkClearContents.Name = "chkClearContents"
        Me.chkClearContents.Size = New System.Drawing.Size(256, 24)
        Me.chkClearContents.TabIndex = 5
        Me.chkClearContents.Text = "Clear Project Contents After Saving"
        '
        'chkUpload
        '
        Me.chkUpload.Location = New System.Drawing.Point(8, -5)
        Me.chkUpload.Name = "chkUpload"
        Me.chkUpload.Size = New System.Drawing.Size(256, 24)
        Me.chkUpload.TabIndex = 4
        Me.chkUpload.Text = "Upload Hours to Time Management System"
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(408, 160)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "&Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOK.Location = New System.Drawing.Point(312, 160)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.TabIndex = 7
        Me.cmdOK.Text = "&OK"
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 56)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(80, 128)
        Me.PictureBox2.TabIndex = 8
        Me.PictureBox2.TabStop = False
        '
        'SaveDialog
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(496, 189)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.gboxSaveInfo)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "SaveDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Upload Hours to Time Management System?"
        Me.gboxSaveInfo.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Methods"
  Private Sub SaveDialog_Load(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles MyBase.Load

    chkUpload_CheckedChanged(Nothing, Nothing)
    Me.dpStartDate.Value = Now()
    While dpStartDate.Value.DayOfWeek <> DayOfWeek.Monday
      dpStartDate.Value = dpStartDate.Value.AddDays(-1)
    End While
  End Sub
  Private Sub dpStartDate_Validating(ByVal sender As Object, ByVal e As _
    System.ComponentModel.CancelEventArgs) Handles dpStartDate.Validating

    If Not Me.dpStartDate.Value.DayOfWeek = DayOfWeek.Monday Then
      MsgBox("You must specify a monday as the start of the week.", _
        MsgBoxStyle.Exclamation, "Wrong Day of Week")
      e.Cancel = True
    End If

  End Sub
  Private Sub chkUpload_CheckedChanged(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles chkUpload.CheckedChanged

    Me.dpStartDate.Enabled = chkUpload.Checked
    Me.chkClearContents.Enabled = chkUpload.Checked
  End Sub

#End Region

#Region "Properites"
  Public ReadOnly Property StartDay() As Date
    Get
      Return Me.dpStartDate.Value
    End Get
  End Property
  Public ReadOnly Property SaveToWeb() As Boolean
    Get
      Return Me.chkUpload.Checked
    End Get
  End Property
  Public ReadOnly Property ClearContents() As Boolean
    Get
      Return Me.chkClearContents.Checked
    End Get
  End Property
#End Region


End Class
