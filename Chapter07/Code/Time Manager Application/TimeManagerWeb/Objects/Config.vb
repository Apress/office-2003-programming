Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings

Public Class Config

    '***************************************************************************
    Public Shared ReadOnly Property connectionString()
        Get
            If dbPassword = String.Empty Then
                Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileLocation & ";User Id=admin;Password=;"
            Else
                Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileLocation & ";Jet OLEDB:Database Password=" & dbPassword & ";"
            End If
        End Get
    End Property

    '***************************************************************************
    Private Shared ReadOnly Property dbPassword() As String
        Get
            Return AppSettings("dbPassword")
        End Get
    End Property

    '***************************************************************************
    Private Shared ReadOnly Property dbFileLocation() As String
        Get
            Return AppSettings("dbFileLocation")
        End Get
    End Property

End Class
