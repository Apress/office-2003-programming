Imports MSC = Microsoft.Office.Core
Imports System.IO

Public Class ListPresentations
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
  Friend WithEvents lstPPTs As System.Windows.Forms.ListBox
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.cbOK = New System.Windows.Forms.Button
    Me.cbCancel = New System.Windows.Forms.Button
    Me.lstPPTs = New System.Windows.Forms.ListBox
    Me.SuspendLayout()
    '
    'cbOK
    '
    Me.cbOK.Location = New System.Drawing.Point(120, 232)
    Me.cbOK.Name = "cbOK"
    Me.cbOK.TabIndex = 0
    Me.cbOK.Text = "&OK"
    '
    'cbCancel
    '
    Me.cbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cbCancel.Location = New System.Drawing.Point(200, 232)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.TabIndex = 1
    Me.cbCancel.Text = "&Cancel"
    '
    'lstPPTs
    '
    Me.lstPPTs.Location = New System.Drawing.Point(8, 24)
    Me.lstPPTs.Name = "lstPPTs"
    Me.lstPPTs.Size = New System.Drawing.Size(272, 199)
    Me.lstPPTs.TabIndex = 0
    '
    'ListPresentations
    '
    Me.AcceptButton = Me.cbOK
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.CancelButton = Me.cbCancel
    Me.ClientSize = New System.Drawing.Size(292, 266)
    Me.Controls.Add(Me.lstPPTs)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.cbOK)
    Me.Name = "ListPresentations"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Available PowerPoint Templates"
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub cbOK_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbOK.Click
    'Use the Templates folder specified in the user settings
    Dim strSelectedPath As String = UserSettings.TemplatesFolder

    Try
      strSelectedPath = strSelectedPath.Concat(strSelectedPath, lstPPTs.SelectedItem)
      'Open the file at specified path
      OpenPPTFile(strSelectedPath)
      Me.Close()
    Catch ex As Exception
      MsgBox(Err.GetException)
    Finally

    End Try
  End Sub

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()

  End Sub

  Private Sub listPresentations_Load(ByVal sender As System.Object, _
  ByVal e As System.EventArgs) Handles MyBase.Load
    'Need a DirectoryInfo object for accessing folder properties
    Dim diFolder As New DirectoryInfo(UserSettings.TemplatesFolder)
    Dim fiPPT As FileInfo
    Try
      'move through each template in the folder and 
      'load their names into the ListBox
      For Each fiPPT In diFolder.GetFiles("*.pot")
        lstPPTs.Items.Add(fiPPT.Name)
      Next

    Catch ex As Exception
      MsgBox(Err.GetException)
    End Try

  End Sub

  Private Function OpenPPTFile(ByVal strFilePath As String) As Boolean
    Try
      'open file using PowerPoint's Open method
      AppPPT.App.Presentations.Open(strFilePath.ToString, MSC.MsoTriState.msoFalse, _
        MSC.MsoTriState.msoTrue)

      Return True
    Catch ex As Exception
      Return False
    End Try

  End Function
End Class
