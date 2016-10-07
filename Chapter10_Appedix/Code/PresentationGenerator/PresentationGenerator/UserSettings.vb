Imports System.Xml
Imports System.Environment 'TODO: update in book
'TODO:  Update Chapter 2 to reflect:
'        
'         1. SaveSettings
'         2. LoadSettings
'         3.  SETTINGS_XML_FILE_NAME
Public Class UserSettings
  Private Shared m_strDatabaseName As String
  Private Shared m_strUserName As String
  Private Shared m_strPassword As String
  Private Shared m_strServerName As String
  Private Shared m_strTemplatesFolder As String
  Private Shared m_strSaveFolder As String
  Private Shared m_strSettingsPath As String
  Private Const SETTINGS_XML_FILE_NAME As String = "\pptgen.xml"

#Region "Public Shared Functions"

  Public Shared Function SaveSettings() As Boolean
    Try
      Dim strPath As String
      strPath = strPath.Concat(UserSettings.SettingsPath, SETTINGS_XML_FILE_NAME)
      Dim xtwSettings As New XmlTextWriter(strPath.ToString, _
        System.Text.Encoding.UTF8)

      With xtwSettings
        .Formatting = Formatting.Indented
        .Indentation = 2
        .QuoteChar = """"c

        .WriteStartDocument(True) 'This creates the ?XML line and identifies the doc as an XML document
        .WriteComment("User Settings from the Bravo Powerpoint Add-In")
        .WriteStartElement("BravoPowerPointUserSettings")
        .WriteAttributeString("UserName", UserSettings.UserName)
        .WriteAttributeString("Password", UserSettings.Password)
        .WriteAttributeString("SaveFolder", UserSettings.SaveFolder)
        .WriteAttributeString("TemplatesFolder", UserSettings.TemplatesFolder)
        .WriteAttributeString("ServerName", UserSettings.ServerName)
        .WriteAttributeString("DatabaseName", UserSettings.DatabaseName)
        .WriteEndElement()

        .WriteEndDocument()
        .Close()
      End With

      Return True
    Catch ex As Exception
      MsgBox(ex.GetBaseException)
      Return False
    End Try
  End Function



  Public Shared Function LoadSettings(ByVal FilePath As String) As Boolean
    Try
      Dim strPath As String
      'TODO: Update the next line in bookin Book it was move from just above the .Close Statement
      m_strSettingsPath = FilePath
      strPath = strPath.Concat(FilePath, SETTINGS_XML_FILE_NAME)
      Dim xtrSettings As New XmlTextReader(strPath.ToString)

      With xtrSettings
        .MoveToContent()
        m_strTemplatesFolder = .GetAttribute("TemplatesFolder")
        m_strSaveFolder = .GetAttribute("SaveFolder")
        m_strDatabaseName = .GetAttribute("DatabaseName")
        m_strServerName = .GetAttribute("ServerName")
        m_strUserName = .GetAttribute("UserName")
        m_strPassword = .GetAttribute("Password")

        .Close()

        Return True
      End With
    Catch ex As Exception
      Return False
    End Try
  End Function
#End Region

#Region "Public Shared Properties"

  Public Shared ReadOnly Property SettingsPath() As String
    Get
      Return m_strSettingsPath
    End Get
    'Set(ByVal strValue As String)
    '  m_strSettingsPath = strValue
    'End Set
  End Property
  Public Shared Property DatabaseName() As String
    Get
      Return m_strDatabaseName
    End Get
    Set(ByVal Value As String)
      m_strDatabaseName = Value
    End Set
  End Property

  Public Shared Property UserName() As String
    Get
      Return m_strUserName

    End Get
    Set(ByVal Value As String)
      m_strUserName = Value
    End Set
  End Property

  Public Shared Property Password() As String
    Get
      Return m_strPassword
    End Get
    Set(ByVal Value As String)
      m_strPassword = Value
    End Set
  End Property

  Public Shared Property ServerName() As String
    Get
      Return m_strServerName
    End Get
    Set(ByVal Value As String)
      m_strServerName = Value
    End Set
  End Property

  Public Shared Property TemplatesFolder() As String
    Get
      Return m_strTemplatesFolder
    End Get
    Set(ByVal Value As String)
      m_strTemplatesFolder = Value
    End Set
  End Property

  Public Shared Property SaveFolder() As String
    Get
      Return m_strSaveFolder
    End Get
    Set(ByVal Value As String)
      m_strSaveFolder = Value
    End Set
  End Property
#End Region


End Class





