Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Module Globals

    '***************************************************************************
    'Contains a number of useful functions.


  Public Function GetConnection(Optional ByVal OpenConnection As Boolean = True) As OleDbConnection
    'Acquires a connection to the database

    Dim dbConn As New OleDbConnection(Config.connectionString)
    If OpenConnection Then dbConn.Open()
    Return dbConn

  End Function

  Public Function quoteBooleanForSQL(ByVal rawBoolean As Boolean)
    'Creates a string for use in a SQL statement

    Return rawBoolean.ToString

  End Function

  Public Function quoteDateForSQL(ByVal rawDate As Date, Optional ByVal DateFormatString As String = "MM/dd/yyyy") As String
    'Creates a string for use in a SQL statement

    If rawDate = Nothing Then
      Return "null"
    Else
      Return "'" & Format(rawDate, DateFormatString) & "'"
    End If

  End Function

  Public Function quoteForSQL(ByVal rawString As String) As String
    'Creates a quoted string for use in a SQL statement

    Return "'" & sqlString(rawString) & "'"

  End Function

  Public Function sqlString(ByVal rawString As String) As String
    'Creates a sql-compliant string for use in a SQL statement

    If rawString = String.Empty Then Return ""
    Return rawString.Replace("'", "''")

  End Function

  Public Function GetDateString(ByVal obj As Object, Optional ByVal DateFormat As String = "yyyy-MM-dd") As String
    'Creates a string for use in a SQL statement

    If obj Is Nothing OrElse IsDBNull(obj) Then
      Return ""
    Else
      Return Format(obj, DateFormat)
    End If

  End Function

End Module
