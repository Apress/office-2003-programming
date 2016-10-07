Public Class EventIssueTracking
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lblHeading As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_CustomerName_Label As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_CustomerEmail_Label As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_IssueDescription_Label As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_CustomerName_Value As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_CustomerEmail_Value As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_IssueDescription_Value As System.Web.UI.WebControls.Label
    Protected WithEvents lblTasks As System.Web.UI.WebControls.Label
    Protected WithEvents repTasks As System.Web.UI.WebControls.Repeater
    Protected WithEvents lbl_CurrentStatus_Label As System.Web.UI.WebControls.Label
    Protected WithEvents lbl_CurrentStatus_Value As System.Web.UI.WebControls.Label
    Protected WithEvents panel_RepHolder As System.Web.UI.WebControls.Panel
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    '**************************************************************************
    Private ReadOnly Property TrackingID() As Long
        'Strongly Type the TrackingID querystring variable
        Get
            If Request.QueryString("TrackingID") = Nothing Then Return 0
            If Not IsNumeric(Request.QueryString("TrackingID")) Then Return 0
            Return CInt(Request.QueryString("TrackingID"))
        End Get
    End Property

    '**************************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim EventIssueObj As EventIssue = EventIssue.Get(TrackingID)
        Dim Tasks As Collection = EventIssueTask.GetTasks(TrackingID)
        Dim CompletedTasks As Integer = 0

        'Ensure that we have something to show
        If EventIssueObj Is Nothing Then
            Page.RegisterStartupScript("error_script", "<SCRIPT language=javascript>alert('The Tracking ID provided could not be located.  Please contact a customer service representative for assistance.');</SCRIPT>")
            Me.lbl_CurrentStatus_Value.Text = "Invalid Tracking ID Number"
            Exit Sub
        End If

        'Setup customer information values
        Me.lbl_CustomerName_Value.Text = EventIssueObj.CustomerName
        Me.lbl_CustomerEmail_Value.Text = EventIssueObj.CustomerEmail
        Me.lbl_IssueDescription_Value.Text = EventIssueObj.IssueDescription

        'Determine completion status information
        For Each Task As EventIssueTask In Tasks
            If Task.Complete Then CompletedTasks += 1
        Next

        'Set status text
        If Tasks.Count > 0 Then
            If CompletedTasks = Tasks.Count Then
                Me.lbl_CurrentStatus_Value.Text = "Resolved (All tasks have been completed)"
            Else
                Me.lbl_CurrentStatus_Value.Text = "In Progress (" & CompletedTasks & " of " & Tasks.Count & " tasks complete)"
            End If
        Else
            Me.lbl_CurrentStatus_Value.Text = "Resolved (There were no tasks to complete)"
        End If

        'Databind the Tasks to the Repeater
        Me.repTasks.DataSource = Tasks
        Me.repTasks.DataBind()

    End Sub


    '**************************************************************************
    Private Sub repTasks_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles repTasks.ItemDataBound
        'Display task information in the repeater

        If e.Item.ItemType = UI.WebControls.ListItemType.Item Or e.Item.ItemType = UI.WebControls.ListItemType.AlternatingItem Then

            Dim lblTask As UI.WebControls.Label = e.Item.FindControl("lblTask")
            Dim lblTaskStatus As UI.WebControls.Label = e.Item.FindControl("lblTaskStatus")
            Dim Task As EventIssueTask = CType(e.Item.DataItem, EventIssueTask)

            lblTask.Text = Task.Task
            lblTaskStatus.Text = IIf(Task.Complete, "Complete", "Incomplete")

        End If

    End Sub

End Class
