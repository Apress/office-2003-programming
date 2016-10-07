Imports System.Data.OleDb
Imports System.Configuration.ConfigurationSettings

Public Class Config

    '***************************************************************************
    'Contains configuration information.  These settings are defined in the
    'web.config file in the <AppSettings> section.


  Public Shared ReadOnly Property connectionString()
    'Creates a connection string to the database.  The appropriate
    'connection string is used if there is a database password.
    Get
      If dbPassword = String.Empty Then
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileLocation & _
          ";User Id=admin;Password=;"
      Else
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbFileLocation & _
          ";Jet OLEDB:Database Password=" & dbPassword & ";"
      End If
    End Get
  End Property

  Private Shared ReadOnly Property dbPassword() As String
    'Password used to access database file (database password)
    Get
      Return AppSettings("dbPassword")
    End Get
  End Property

  Private Shared ReadOnly Property dbFileLocation() As String
    'Location of the access database file
    Get
      Return AppSettings("dbFileLocation")
    End Get
  End Property

  Public Shared ReadOnly Property TaskFormHREF() As String
    'URL of the Task Form Infopath Template file (.XSN file)
    Get
      Return AppSettings("taskFormHREF")
    End Get
  End Property

  Public Shared ReadOnly Property TempMailDirectory() As String
    'Directory in which to create attachments for email messages
    Get
      Return AppSettings("tempMailDirectory")
    End Get
  End Property

  Public Shared ReadOnly Property SmtpServer() As String
    'SMTP server used to send email
    Get
      Return AppSettings("SmtpServer")
    End Get
  End Property

  Public Shared ReadOnly Property IssueTrackingPageURL() As String
    'Fully qualified URL of the IssueTrackingPage
    Get
      Return AppSettings("IssueTrackingPageURL")
    End Get
  End Property

End Class
