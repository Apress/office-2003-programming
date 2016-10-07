Imports System.Data.OleDb

Public Class tmUser

    '***************************************************************************
    Public userId As String = ""
    Public password As String = ""
    Public nameLast As String = ""
    Public nameFirst As String = ""
    Public admin As Boolean = False


    '***************************************************************************
    Public Shared Function GetFromDR(ByRef DR As IDataReader) As tmUser
        Try
            Dim obj As New tmUser
            obj.userId = DR("userId")
            obj.password = DR("password")
            obj.nameLast = DR("nameLast")
            obj.nameFirst = DR("nameFirst")
            obj.admin = DR("admin")
            Return obj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '***************************************************************************
    Public Shared Function AddProject(ByVal ProjectName As String, Optional ByVal dbConn As IDbConnection = Nothing, Optional ByVal dbTran As IDbTransaction = Nothing) As Boolean
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
            dbCmd.CommandText = "INSERT INTO [tblProject](projectName) VALUES ('" & sqlString(ProjectName) & "');"
            If dbCmd.ExecuteNonQuery = 1 Then
                AddProject = True
            Else
                AddProject = False
            End If

            If dbConnOwner Then dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

End Class