Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing


Public Class GetUpdates
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
  Friend WithEvents cbDownload As System.Windows.Forms.Button
  Friend WithEvents lblResult As System.Windows.Forms.Label
  Friend WithEvents pnlInstructions As System.Windows.Forms.Panel
  Friend WithEvents lblInstructions As System.Windows.Forms.Label
  Friend WithEvents cbCancel As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.pnlInstructions = New System.Windows.Forms.Panel
    Me.lblInstructions = New System.Windows.Forms.Label
    Me.cbDownload = New System.Windows.Forms.Button
    Me.lblResult = New System.Windows.Forms.Label
    Me.cbCancel = New System.Windows.Forms.Button
    Me.pnlInstructions.SuspendLayout()
    Me.SuspendLayout()
    '
    'pnlInstructions
    '
    Me.pnlInstructions.BackColor = System.Drawing.Color.White
    Me.pnlInstructions.Controls.Add(Me.lblInstructions)
    Me.pnlInstructions.Location = New System.Drawing.Point(0, 0)
    Me.pnlInstructions.Name = "pnlInstructions"
    Me.pnlInstructions.Size = New System.Drawing.Size(300, 80)
    Me.pnlInstructions.TabIndex = 0
    '
    'lblInstructions
    '
    Me.lblInstructions.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblInstructions.Location = New System.Drawing.Point(8, 20)
    Me.lblInstructions.Name = "lblInstructions"
    Me.lblInstructions.Size = New System.Drawing.Size(275, 40)
    Me.lblInstructions.TabIndex = 1
    Me.lblInstructions.Text = "Press the Download button to update the templates installed on your system."
    '
    'cbDownload
    '
    Me.cbDownload.Location = New System.Drawing.Point(80, 112)
    Me.cbDownload.Name = "cbDownload"
    Me.cbDownload.Size = New System.Drawing.Size(112, 23)
    Me.cbDownload.TabIndex = 0
    Me.cbDownload.Text = "&Download Updates"
    '
    'lblResult
    '
    Me.lblResult.Location = New System.Drawing.Point(8, 160)
    Me.lblResult.Name = "lblResult"
    Me.lblResult.Size = New System.Drawing.Size(275, 96)
    Me.lblResult.TabIndex = 2
    Me.lblResult.Visible = False
    '
    'cbCancel
    '
    Me.cbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cbCancel.Location = New System.Drawing.Point(208, 112)
    Me.cbCancel.Name = "cbCancel"
    Me.cbCancel.Size = New System.Drawing.Size(72, 23)
    Me.cbCancel.TabIndex = 2
    Me.cbCancel.Text = "&Cancel"
    '
    'GetUpdates
    '
    Me.AcceptButton = Me.cbDownload
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.CancelButton = Me.cbCancel
    Me.ClientSize = New System.Drawing.Size(292, 266)
    Me.Controls.Add(Me.cbCancel)
    Me.Controls.Add(Me.lblResult)
    Me.Controls.Add(Me.cbDownload)
    Me.Controls.Add(Me.pnlInstructions)
    Me.Name = "GetUpdates"
    Me.Text = "Download Template Updates"
    Me.pnlInstructions.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub cbDownload_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbDownload.Click

    If DownloadTemplates() Then
      lblResult.Text = "Templates downloaded to the folder path " & _
        UserSettings.TemplatesFolder & "."
      lblResult.Visible = True
    Else
      lblResult.Text = "An error occurred attempting to update your templates."
      lblResult.Visible = True
    End If

    Me.Size = New Size(300, 300)


  End Sub



  Private Function DownloadTemplates() As Boolean
    Dim cnnPPT As New SqlConnection
    Dim daPPT As New SqlDataAdapter("Select * From tblPresentations", cnnPPT)
    Dim dsPPTs As New DataSet
    Dim dtPPTs As DataTable
    Dim drPPTRecord As DataRow
    Dim btPPTBinary() As Byte
    Dim iPPTSize As Long
    Dim strCnn As String

    Try
      'Create connection string
      strCnn = "Server=" & UserSettings.ServerName
      strCnn = strCnn.Concat(strCnn, ";uid=" & UserSettings.UserName)
      strCnn = strCnn.Concat(strCnn, ";pwd=" & UserSettings.Password)
      strCnn = strCnn.Concat(strCnn, ";database=" & UserSettings.DatabaseName)

      cnnPPT.ConnectionString = strCnn.ToString
      cnnPPT.Open()
      'Fill the Dataset with Presentation templates
      daPPT.Fill(dsPPTs, "tblPresentations")


      'Loop through all records and save to the Default Save Location
      For Each drPPTRecord In dsPPTs.Tables("tblPresentations").Rows
        btPPTBinary = drPPTRecord("PresentationBinary")

        Dim strPath As String = UserSettings.TemplatesFolder
        strPath = strPath.Concat(strPath, drPPTRecord("PresentationName").ToString)

        Dim fsFile As New FileStream(strPath, FileMode.Create)
        iPPTSize = UBound(btPPTBinary)
        fsFile.Write(btPPTBinary, 0, iPPTSize)
        fsFile.Close()
        fsFile = Nothing
        strPath = Nothing
      Next drPPTRecord

      Return True

    Catch ex As Exception
      Return False
    Finally

      cnnPPT.Close()
      drPPTRecord = Nothing
      dsPPTs = Nothing
      dtPPTs = Nothing
      daPPT = Nothing
      cnnPPT = Nothing
    End Try
  End Function

  Private Sub cbCancel_Click(ByVal sender As System.Object, _
    ByVal e As System.EventArgs) Handles cbCancel.Click

    Me.Close()
  End Sub
  Private Sub GetUpdates_Load(ByVal sender As Object, _
    ByVal e As System.EventArgs) Handles MyBase.Load

    Me.Size = New Size(300, 175)
  End Sub

#Region "Not Covered in Book"
  Private Sub SaveFileToDB(ByVal FilePath As String)
    'TODO:  Review this code and make it yours

    Dim con As New SqlConnection("Server=CoreServ;uid=sa;pwd=cr3d3ra;database=BravoCorp")
    Dim da As New SqlDataAdapter("Select * From tblPresentations", con)
    Dim MyCB As SqlCommandBuilder = New SqlCommandBuilder(da)
    Dim ds As New DataSet
    Dim fs As New FileStream(FilePath.ToString, FileMode.OpenOrCreate, FileAccess.Read)
    Dim fi As New FileInfo(FilePath.ToString)

    Dim MyData(fs.Length) As Byte
    Dim myRow As DataRow

    da.MissingSchemaAction = MissingSchemaAction.AddWithKey

    fs.Read(MyData, 0, fs.Length)
    fs.Close()
    con.Open()
    da.Fill(ds, "tblPresentations")

    myRow = ds.Tables("tblPresentations").NewRow()
    myRow("PresentationDesc") = "This would be description text"
    myRow("PresentationID") = 200
    myRow("PresentationName") = fi.Name
    myRow("PresentationBinary") = MyData
    ds.Tables("tblPresentations").Rows.Add(myRow)
    da.Update(ds, "tblPresentations")

    fs = Nothing
    MyCB = Nothing
    ds = Nothing
    da = Nothing

    con.Close()
    con = Nothing
    MsgBox("Slide saved to database")
  End Sub

#End Region


End Class
