

Public Class tmHours

    '***************************************************************************
    Public ProjectName As String = ""
    Public userId As String = ""
    Public hours As Single = 0
    Public startDate As DateTime = Nothing
    Public endDate As DateTime = Nothing
    Public description As String = ""

    '***************************************************************************
    Public Shared Function GetFromDR(ByRef DR As IDataReader) As tmHours
        Try
            Dim obj As New tmHours
            obj.ProjectName = DR("projectName")
            obj.userId = DR("userId")
            obj.startDate = DR("startDate")
            obj.endDate = DR("endDate")
            obj.hours = DR("hours")
            obj.description = DR("description")
            Return obj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '***************************************************************************
    Public Shared Function DeleteForGivenWeek(ByVal userId As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByRef dbConn As IDbConnection = Nothing, Optional ByRef dbTran As IDbTransaction = Nothing) As Boolean
        '-----------------------------------------------------------------------
        '   Information
        '-----------------------------------------------------------------------

        Try

            Dim dbConnOwner As Boolean = dbConn Is Nothing
            Dim dbCmd As IDbCommand

            If dbConn Is Nothing Then dbConn = Globals.GetConnection()

            dbCmd = dbConn.CreateCommand()
            dbCmd.Transaction = dbTran

            'Insert the new items
            dbCmd.CommandText = "DELETE FROM [tblHours] WHERE [userId]='" & sqlString(userId) & "' AND [startDate]=#" & Format(startDate, "MM/dd/yyyy") & "# and [endDate]=#" & Format(endDate, "MM/dd/yyyy") & "#;"
            dbCmd.ExecuteNonQuery()
            If dbConnOwner Then dbConn.Close()
            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    Public Function Save(Optional ByRef dbConn As IDbConnection = Nothing, Optional ByRef dbTran As IDbTransaction = Nothing) As Boolean
        '-----------------------------------------------------------------------
        '   Information
        '-----------------------------------------------------------------------

        Try

            Dim dbConnOwner As Boolean = dbConn Is Nothing
            Dim dbCmd As IDbCommand

            If dbConn Is Nothing Then dbConn = Globals.GetConnection()

            dbCmd = dbConn.CreateCommand()
            dbCmd.Transaction = dbTran

            'Insert the new items
            dbCmd.CommandText = "INSERT INTO [tblHours](projectName, userId, startDate, endDate, hours, description) VALUES ('" & sqlString(Me.ProjectName) & "','" & sqlString(Me.userId) & "',#" & Format(startDate, "MM/dd/yyyy") & "#,#" & Format(endDate, "MM/dd/yyyy") & "#,'" & Me.hours & "','" & sqlString(Me.description) & "');"
            If dbCmd.ExecuteNonQuery() = 1 Then
                Save = True
                If dbConnOwner Then dbConn.Close()
            Else
                Save = False
                If dbConnOwner Then dbConn.Close()
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function


    '***************************************************************************


End Class
