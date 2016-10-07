Imports System.Data.OleDb

Public Module Globals

    '***************************************************************************
    Public Function CreateStringArray(ByVal ParamArray Values() As String) As String()
        Return Values
    End Function

    '***************************************************************************
    Public Function GetConnection(Optional ByVal OpenConnection As Boolean = True) As OleDbConnection
        Dim dbConn As New OleDbConnection(Config.connectionString)
        If OpenConnection Then dbConn.Open()
        Return dbConn
    End Function


    '***************************************************************************
    Public Sub EnsureConnection(ByRef dbConn As OleDbConnection, Optional ByRef dbOwnerFlag As Boolean = False, Optional ByVal OpenConnection As Boolean = True)
        If dbConn Is Nothing Then
            dbOwnerFlag = True
            dbConn = GetConnection(OpenConnection)
        Else
            dbOwnerFlag = False
        End If
    End Sub

    '***************************************************************************
    Public Function quoteForSQL(ByVal rawString As String) As String
        Return "'" & sqlString(rawString) & "'"
    End Function

    '***************************************************************************
    Public Function sqlString(ByVal rawString As String) As String
        If rawString = String.Empty Then Return ""
        Return rawString.Replace("'", "''")
    End Function


End Module
