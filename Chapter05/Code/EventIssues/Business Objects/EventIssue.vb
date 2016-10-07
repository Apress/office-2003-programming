Imports System.Data.OleDb

Public Class EventIssue

    '**************************************************************************
    Public TrackingId As Long = 0
    Public CustomerName As String = ""
    Public CustomerEmail As String = ""
    Public IssueDescription As String = ""

    '**************************************************************************
    Public Shared Function Add(ByRef EventIssueObj As EventIssue) As Boolean

        'Adds an Issue to the [Issue] Table

        Try
            Dim SQL As String = "INSERT INTO [Issue](customer_name, customer_email, problem) VALUES (" & quoteForSQL(EventIssueObj.CustomerName) & "," & quoteForSQL(EventIssueObj.CustomerEmail) & "," & quoteForSQL(EventIssueObj.IssueDescription) & ");"
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)

            If dbCmd.ExecuteNonQuery() > 0 Then
                dbCmd.CommandText = "SELECT @@IDENTITY;"
                EventIssueObj.TrackingId = dbCmd.ExecuteScalar()
                Add = True
            Else
                Add = False                
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '**************************************************************************
    Public Shared Function [Get](ByVal TrackingID As Long) As EventIssue

        'Gets an Issue using the Tracking ID number

        Try

            Dim SQL As String = "SELECT * FROM [Issue] WHERE [tracking_id]=" & TrackingID
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand(SQL, dbConn)
            Dim dbReader As OleDbDataReader = dbCmd.ExecuteReader

            If dbReader.Read Then
                [Get] = GetDataFromReader(dbReader)
            Else
                [Get] = Nothing
            End If

            dbConn.Close()

        Catch ex As Exception
            Return Nothing
        End Try

    End Function


    '**************************************************************************
    Private Shared Function GetDataFromReader(ByRef dbReader As OleDbDataReader) As EventIssue

        'Pulls data from a Data Reader and Places it into an EventIssue object

        Dim obj As New EventIssue
        obj.TrackingId = dbReader("tracking_id")
        obj.CustomerName = dbReader("customer_name")
        obj.CustomerEmail = dbReader("customer_email")
        obj.IssueDescription = dbReader("problem")
        Return obj

    End Function


End Class
