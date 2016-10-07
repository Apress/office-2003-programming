Imports System.Web.Services
Imports System.Data
Imports System.Data.OleDb
Imports System.Xml
Imports Microsoft.VisualBasic.ControlChars


<System.Web.Services.WebService(Namespace:="http://localhost/EventIssues/")> _
Public Class EventIssues
    Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

  <WebMethod()> _
  Public Sub SubmitEventIssueForm(ByVal XML As String)
    'Allows for the submission of the Event Issue Resolution form from
    'Infopath.  There is no all-encompasing error handling because 
    'InfoPath will automatically recongnize errors and return the error
    'information to the user.

    Dim EventIssueObj As New EventIssue
    Dim TaskCollection As New Collection
    Dim EventIssueTaskObj As EventIssueTask
    Dim TaskOrder As Long = 1
    Dim xmlReader As XmlTextReader = New XmlTextReader(XML, _
      XmlNodeType.Element, Nothing)


    'Read through each XML node and process the node according to the 
    'node type and element name.

    Do While xmlReader.Read()

      'We are only interested in processing actual elements in this 
      'example.  You may need to use a 'Select Case' for the node 
      'type for other processing scenarios.

      If xmlReader.NodeType = XmlNodeType.Element Then

        'The first select case is used to capture Issue information
        'from the XML data.  There is only one Issue represented in
        'the XML data.

        Select Case xmlReader.Name

          Case "my:customer_name"
            xmlReader.Read()
            EventIssueObj.CustomerName = xmlReader.Value

          Case "my:email_address"
            xmlReader.Read()
            EventIssueObj.CustomerEmail = xmlReader.Value

          Case "my:issue_description"
            xmlReader.Read()
            EventIssueObj.IssueDescription = xmlReader.Value

          Case "my:Task"

            'There can be multiple tasks associated with any given
            'issue.  Whenever we run across the 'my:Task' element,
            'we need to instantiate a new EventIssueTask object
            'and populate it with the coorsponding data.  We do
            'not exit this part of the Case statement until we 
            'encounter the </my:Task> end element.

            EventIssueTaskObj = New EventIssueTask

            Do Until xmlReader.NodeType = XmlNodeType.EndElement And _
              xmlReader.Name = "my:Task"

              xmlReader.Read()

              If xmlReader.NodeType = XmlNodeType.Element Then

                Select Case xmlReader.Name

                  Case "my:employee_email"
                    xmlReader.Read()
                    EventIssueTaskObj.AssigneeEmail = xmlReader.Value

                  Case "my:task_description"
                    xmlReader.Read()
                    EventIssueTaskObj.Task = xmlReader.Value

                  Case "my:task_due_date"
                    xmlReader.Read()

                    If Not xmlReader.Value = String.Empty Then
                      EventIssueTaskObj.TaskDueDate = CDate(xmlReader.Value)
                    End If

                  Case "my:complete_before_next_task"
                    xmlReader.Read()
                    EventIssueTaskObj.TaskOrder = TaskOrder

                    If xmlReader.Value = "true" Then
                      TaskOrder += 1
                    End If

                End Select

              End If

            Loop

            TaskCollection.Add(EventIssueTaskObj)

        End Select
      End If
    Loop

    'Save the Issue and acquire the TrackingId given to the item from
    'the database.  Then set the TrackingId on all of the tasks before
    'saving them.

    If EventIssue.Add(EventIssueObj) Then
      For Each EventIssueTaskObj In TaskCollection
        EventIssueTaskObj.TrackingId = EventIssueObj.TrackingId
        If Not EventIssueTask.Add(EventIssueTaskObj) Then
          Throw New Exception("A task failed to save but did not " & _
            "explicitly throw an exception")
        End If
      Next

      Dim EmployeeFullName As String = ""
      Dim EmployeeOccupation As String = ""

      SendNotifications(EventIssueObj.TrackingId, 1, _
        EventIssueObj.CustomerName, EventIssueObj.CustomerEmail, _
        EventIssueObj.IssueDescription)

      'Send a message to the customer informing them that they can
      'go to the website to view the status of their issue.
      Mail.SmtpMail.SmtpServer = Config.SmtpServer
      Mail.SmtpMail.Send("EventIssues@localhost.com", _
        EventIssueObj.CustomerEmail, "Event Issue Ticket Opened", _
        "You have recently informed us that you are having the" & _
        "following issue: " & CrLf & CrLf & EventIssueObj.IssueDescription & _
        CrLf & CrLf & CrLf & "We are committed to resolving this issue" & _
        "and have outlined a series of tasks that need to be completed" & _
        "in order to resolve your issue.  These will be completed as" & _
        "as possible.  You may review the status of this issue by " & _
        "navigating to " & Config.IssueTrackingPageURL & "?TrackingID=" & _
        EventIssueObj.TrackingId & CrLf & CrLf & "Thank You.")

    Else
      Throw New Exception("The event issue failed to save but did not" & _
        "explicitly throw an exception")
    End If

  End Sub

  <WebMethod()> _
  Public Sub SubmitTaskForm(ByVal XML As String)
    'Allows for the submission of the Event Issue Resolution form from
    'Infopath.  There is no all-encompasing error handling because 
    'InfoPath will automatically recongnize errors and return the error
    'information to the user.

    Dim EventIssueTaskObj As New EventIssueTask
    Dim xmlReader As XmlTextReader = New XmlTextReader(XML, _
      XmlNodeType.Element, Nothing)
    Dim CustomerName As String = ""
    Dim CustomerEmail As String = ""
    Dim IssueDescription As String = ""
    Dim EmployeeFullName As String = ""
    Dim EmployeeOccupation As String = ""

    'Read through each XML node and process the node according to the 
    'node type and element name.

    Do While xmlReader.Read()

      'We are only interested in processing actual elements in this 
      'example.  You may need to use a 'Select Case' for the node 
      'type for other processing scenarios.

      If xmlReader.NodeType = XmlNodeType.Element Then

        'The first select case is used to capture Issue information
        'from the XML data.  There is only one Issue represented in
        'the XML data.

        Select Case xmlReader.Name

          Case "my:task_id"
            xmlReader.Read()
            EventIssueTaskObj.TaskId = xmlReader.Value()

          Case "my:task_order"
            xmlReader.Read()
            EventIssueTaskObj.TaskOrder = xmlReader.Value()

          Case "my:tracking_id"
            xmlReader.Read()
            EventIssueTaskObj.TrackingId = xmlReader.Value

          Case "my:task"
            xmlReader.Read()
            EventIssueTaskObj.Task = xmlReader.Value

          Case "my:assignee_email"
            xmlReader.Read()
            EventIssueTaskObj.AssigneeEmail = xmlReader.Value

          Case "my:comments"
            xmlReader.Read()
            EventIssueTaskObj.Comments = xmlReader.Value

          Case "my:complete"
            xmlReader.Read()
            EventIssueTaskObj.Complete = CBool(xmlReader.Value)

          Case "my:task_due_date"
            xmlReader.Read()
            If Not xmlReader.Value = String.Empty Then
              EventIssueTaskObj.TaskDueDate = CDate(xmlReader.Value)
            End If

          Case "my:customer_name"
            xmlReader.Read()
            CustomerName = xmlReader.Value

          Case "my:customer_email"
            xmlReader.Read()
            CustomerEmail = xmlReader.Value

          Case "my:issue_description"
            xmlReader.Read()
            IssueDescription = xmlReader.Value

          Case "my:employee_fullname"
            xmlReader.Read()
            EmployeeFullName = xmlReader.Value

          Case "my:occupation"
            xmlReader.Read()
            EmployeeOccupation = xmlReader.Value

        End Select
      End If
    Loop

    'Save the Issue and acquire the TrackingId given to the item from
    'the database.  Then set the TrackingId on all of the tasks before
    'saving them.

    If EventIssueTask.Update(EventIssueTaskObj) Then
      SendNotifications(EventIssueTaskObj.TrackingId, _
        EventIssueTaskObj.TaskOrder, CustomerName, CustomerEmail, _
        IssueDescription)
    Else
      Throw New Exception("The event issue failed to save but did not " & _
        "explicitly throw an exception")
    End If

  End Sub

  <WebMethod()> _
  Public Function GetUsers() As System.Xml.XmlDocument

    'This will return an XML documenting containing user data for
    'use in a drop down list in InfoPath.

    Try
      Dim dbConn As OleDbConnection = GetConnection()
      Dim dbCmd As New OleDbCommand
      Dim dbAdapter As New OleDbDataAdapter
      Dim dbUserInfo As New DataSet

      dbCmd.CommandText = "SELECT *, [last_name] + ', ' + [first_name] + " & _
        "' (' + [occupation] + ')' as [name_and_occupation] FROM [Users] " & _
        "ORDER BY [last_name], [first_name]"
      dbCmd.Connection = dbConn
      dbAdapter.SelectCommand = dbCmd

      dbUserInfo.DataSetName = "UserInfo"
      'TODO: You can remove this line if it doesn't jack everything up: dbUserInfo.Namespace = "http://localhost/EventIssues/UserInfo"
      dbAdapter.Fill(dbUserInfo, "[Users]")

      dbConn.Close()

      'This will create an XML Document to return to InfoPath

      Dim xmlDoc As New System.Xml.XmlDocument
      xmlDoc.LoadXml(dbUserInfo.GetXml())
      Return xmlDoc

    Catch ex As Exception
      Return Nothing
    End Try

  End Function

  Public Sub SendNotifications(ByVal TrackingId As Long, _
    ByVal CurrentTaskOrderId As Long, ByVal CustomerName _
    As String, ByVal CustomerEmail As String, _
    ByVal IssueDescription As String)

    Try

      Dim SQL As String = "SELECT Task.*, Users.last_name + ', ' + " & _
        "Users.first_name as [employee_fullname], Users.occupation FROM " & _
        "Task INNER JOIN Users ON Task.assignee_email = Users.email " & _
        "WHERE [tracking_id]=" & TrackingId & " and [complete]=false " & _
        "order by [Task_Order]"
      Dim dbConn As OleDbConnection = GetConnection()
      Dim dbCmd As New OleDbCommand(SQL, dbConn)
      Dim dbReader As OleDbDataReader = dbCmd.ExecuteReader()
      Dim InitialTaskOrderId As Long = 0
      Dim HasTasks As Boolean = False
      Dim Done As Boolean = False
      Dim ns As String = "http://www.w3.org/2001/XMLSchema-instance"

      Dim XML As System.Text.StringBuilder
      Dim fileName As String
      Dim SW As System.IO.StreamWriter

      Dim Message As Mail.MailMessage
      Dim FileAttachment As Mail.MailAttachment

      If dbReader.Read Then
        HasTasks = True
        InitialTaskOrderId = dbReader.Item("task_order")

        While Not Done AndAlso dbReader.Item("task_order") = InitialTaskOrderId

          If dbReader.Item("sent") = False Then

            'Create the XML to attach
            'Copied the schema from Infopath
            XML = New System.Text.StringBuilder(500)
            XML.Append("<?xml version=""1.0""?><?mso-infoPathSolution" & _
              "productVersion=""11.0.6250"" PIVersion=""1.0.0.0"" href=""")
            XML.Append(Config.TaskFormHREF)
            XML.Append(""" name=""urn:schemas-microsoft-com:office:" & _
              "infopath:TaskForm:-myXSD-2004-07-12T03-02-23"" " & _
              "solutionVersion=""1.0.0.11"" ?><?mso-application " & _
              "progid=""InfoPath.Document""?><my:TaskData xmlns:" & _
              "my=""http://schemas.microsoft.com/office/infopath/2003/" & _
              "myXSD/2004-07-12T03:02:23"" xml:lang=""en-us"">")
            XML.Append(CrLf)
            XML.Append("    <my:task_id xmlns:xsi=" & ns & ">")
            XML.Append(dbReader.Item("task_id"))
            XML.Append("</my:task_id>")
            XML.Append(CrLf)
            XML.Append("    <my:task_order xmlns:xsi=" & ns & ">")
            XML.Append(dbReader.Item("task_order"))
            XML.Append("</my:task_order>")
            XML.Append(CrLf)
            XML.Append("    <my:tracking_id xmlns:xsi=" & ns & ">")
            XML.Append(dbReader.Item("tracking_id"))
            XML.Append("</my:tracking_id>")
            XML.Append(CrLf)
            XML.Append("    <my:task>")
            XML.Append(dbReader.Item("task"))
            XML.Append("</my:task>")
            XML.Append(CrLf)
            XML.Append("    <my:assignee_email>")
            XML.Append(dbReader.Item("assignee_email"))
            XML.Append("</my:assignee_email>")
            XML.Append(CrLf)
            XML.Append("    <my:comments>")
            XML.Append(dbReader.Item("comments"))
            XML.Append("</my:comments>")
            XML.Append(CrLf)
            XML.Append("    <my:complete>")
            XML.Append(CBool(dbReader.Item("complete")).ToString.ToLower)
            XML.Append("</my:complete>")
            XML.Append(CrLf)
            XML.Append("    <my:task_due_date xmlns:xsi=" & ns & ">")
            XML.Append(GetDateString(dbReader.Item("task_due_date")))
            XML.Append("</my:task_due_date>")
            XML.Append(CrLf)
            XML.Append("    <my:CustomerData>")
            XML.Append(CrLf)
            XML.Append("        <my:customer_name>")
            XML.Append(CustomerName)
            XML.Append("</my:customer_name>")
            XML.Append(CrLf)
            XML.Append("        <my:customer_email>")
            XML.Append(CustomerEmail)
            XML.Append("</my:customer_email>")
            XML.Append(CrLf)
            XML.Append("        <my:issue_description>")
            XML.Append(IssueDescription)
            XML.Append("</my:issue_description>")
            XML.Append(CrLf)
            XML.Append("    </my:CustomerData>")
            XML.Append(CrLf)
            XML.Append("    <my:EmployeeData>")
            XML.Append(CrLf)
            XML.Append("        <my:employee_fullname>")
            XML.Append(dbReader.Item("employee_fullname"))
            XML.Append("</my:employee_fullname>")
            XML.Append(CrLf)
            XML.Append("        <my:occupation>")
            XML.Append(dbReader.Item("occupation"))
            XML.Append("</my:occupation>")
            XML.Append(CrLf)
            XML.Append("    </my:EmployeeData>")
            XML.Append(CrLf)
            XML.Append("</my:TaskData>")

            'Write the XML to a temporary file
            fileName = Config.TempMailDirectory & "\Task" & _
              dbReader.Item("task_id") & ".xml"

            SW = New System.IO.StreamWriter(fileName, False)
            SW.Write(XML.ToString)
            SW.Close()

            'Create an Email to Send the XML File
            Message = New Mail.MailMessage
            Message.Subject = "Task to be Completed"
            Message.To = dbReader.Item("assignee_email")
            Message.Body = "Attached is a Task Completion Form.  Please " & _
              "complete the task as soon as possible and mark it as " & _
              "completed on the form.  You may also add any comments you" & _
              "want in the comments section.  Thank You."
            Message.From = "EventIssues@localhost.com"

            'Attach XML File
            FileAttachment = New Mail.MailAttachment(fileName)
            Message.Attachments.Add(FileAttachment)
            Mail.SmtpMail.SmtpServer = Config.SmtpServer
            Mail.SmtpMail.Send(Message)

            'Mark the Item as having been sent
            EventIssueTask.MarkAsSent(dbReader.Item("task_id"))
          End If

          'Move on to the next record
          Done = Not dbReader.Read()

        End While

      Else

        'There are no more tasks, so notify the customer
        Message = New Mail.MailMessage
        Message.Subject = "Notification of Issue Resolution"
        Message.To = CustomerEmail
        Message.Body = "This email was sent to notify you that your " & _
          "issue (Tracking ID: " & TrackingId & ") has been resolved.  " & _
          "If you are not satisfied with the resolution, please contact " & _
          "your customer service representative for more information."
        Message.From = "EventIssues@localhost.com"
        Mail.SmtpMail.SmtpServer = Config.SmtpServer
        Mail.SmtpMail.Send(Message)

      End If


    Catch ex As Exception
      Throw New Exception("Error sending notifications. Form data was " & _
        "successfully submitted.", ex)
    End Try

  End Sub

End Class
