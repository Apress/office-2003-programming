Imports System.Data.OleDb

Public Class EventIssueTask

    '**************************************************************************
    Public TaskId As Long = 0
    Public TaskOrder As Long = 0
    Public TrackingId As Long = 0
    Public Task As String = ""
    Public AssigneeEmail As String = ""
    Public Comments As String = ""
    Public Complete As Boolean = False
    Public TaskDueDate As Date = Nothing
    Public Sent As Boolean = False

    '***************************************************************************
    Public Shared Function Add(ByRef EventIssueTaskObj As EventIssueTask) As Boolean
        'Adds a task to the [Task] table in the database.  Also, returns the 
        'TaskId of the inserted record.

        Try
            Dim SQL As String = "INSERT INTO [Task](task_due_date, task_order, tracking_id, task, assignee_email, comments, complete, sent) VALUES (" & quoteDateForSQL(EventIssueTaskObj.TaskDueDate) & "," & quoteForSQL(EventIssueTaskObj.TaskOrder) & "," & quoteForSQL(EventIssueTaskObj.TrackingId) & "," & quoteForSQL(EventIssueTaskObj.Task) & "," & quoteForSQL(EventIssueTaskObj.AssigneeEmail) & "," & quoteForSQL(EventIssueTaskObj.Comments) & "," & quoteBooleanForSQL(EventIssueTaskObj.Complete) & ", false);"
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)

            If dbCmd.ExecuteNonQuery() > 0 Then
                dbCmd.CommandText = "SELECT @@IDENTITY;"
                EventIssueTaskObj.TaskId = dbCmd.ExecuteScalar()
                Add = True
            Else
                Add = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    Public Shared Function Update(ByRef EventIssueTaskObj As EventIssueTask) As Boolean
        'Updates the necessary fields in the [task] table.  There are only two
        'fields [comments] and [complete] that get changed during a Task Form
        'submission.  All other fields remain the same.

        Try
            Dim SQL As String = "UPDATE [Task] SET [comments]=" & quoteForSQL(EventIssueTaskObj.Comments) & ", [complete]=" & quoteBooleanForSQL(EventIssueTaskObj.Complete) & " WHERE [task_id]=" & EventIssueTaskObj.TaskId
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)

            If dbCmd.ExecuteNonQuery() > 0 Then
                Update = True
            Else
                Update = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    Public Shared Function MarkAsSent(ByVal TaskId As Long) As Boolean
        'Marks a task as having had a notification sent to the assignee.

        Try
            Dim SQL As String = "UPDATE [Task] SET [sent]=true WHERE [task_id]=" & TaskId
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)

            If dbCmd.ExecuteNonQuery() > 0 Then
                MarkAsSent = True
            Else
                MarkAsSent = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try


    End Function

    '***************************************************************************
    Public Shared Function GetTasks(ByVal TrackingId As Long) As Collection
        'Returns a collection of tasks for a specific Tracking ID

        Dim TaskCollection As New Collection

        Try

            Dim SQL As String = "SELECT * FROM [task] WHERE [tracking_id]=" & TrackingId
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)
            Dim dbReader As OleDbDataReader = dbCmd.ExecuteReader

            While dbReader.Read
                TaskCollection.Add(GetDataFromReader(dbReader))
            End While

            dbConn.Close()

            Return TaskCollection
        Catch ex As Exception
            Return TaskCollection
        End Try

    End Function

    '***************************************************************************
    Private Shared Function GetDataFromReader(ByRef dbReader As OleDbDataReader) As EventIssueTask
        'Places DataReader data into a Task Object

        Dim obj As New EventIssueTask
        obj.TaskId = dbReader("task_id")
        obj.TaskOrder = dbReader("task_order")
        obj.TrackingId = dbReader("tracking_id")
        obj.Task = dbReader("task")
        obj.AssigneeEmail = dbReader("assignee_email")
        obj.Comments = dbReader("comments")
        obj.Complete = dbReader("complete")
        obj.Sent = dbReader("sent")

        If Not IsDBNull(dbReader("task_due_date")) Then _
            obj.TaskDueDate = dbReader("task_due_date")
        Return obj

    End Function

End Class
