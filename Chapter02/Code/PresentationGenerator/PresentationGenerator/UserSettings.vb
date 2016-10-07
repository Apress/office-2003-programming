Imports System.Xml
Imports System.Environment
Imports System.IO

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
      strPath = strPath.Concat(m_strSettingsPath, SETTINGS_XML_FILE_NAME)
      'set the settings path variable just in case
      m_strSettingsPath = strPath
      Dim xtwSettings As New XmlTextWriter(strPath.ToString, _
        System.Text.Encoding.UTF8)

      'Set XML Writer Settings
      With xtwSettings
        .Formatting = Formatting.Indented
        .Indentation = 2
        .QuoteChar = """"c

        'This creates the ?XML line and identifies 
        'the doc as an XML document
        .WriteStartDocument(True)
        .WriteComment("User Settings from the Bravo Powerpoint Add-In")

        'Begin writing each setting to the XML file as 
        'Attributes of a single XML Element
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
      'if settings file exist then we know we are okay
      m_strSettingsPath = FilePath
      'create full path to settings file
      Dim strPath As String
      strPath = strPath.Concat(FilePath, SETTINGS_XML_FILE_NAME)
      'Test for the existence of the settings file
      Dim fi As New FileInfo(strPath)
      If fi.Exists Then
        'open the settings file for reading into memory
        Dim xtrSettings As New XmlTextReader(strPath.ToString)
        'move through the XML content and fill variables
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
      Else
        'Settings do not exist so alert the user...
        MsgBox("The Presentation Generator settings file does not exist." & _
         vbCrLf & vbCrLf & _
         "Please set your settings now.", MsgBoxStyle.Information, _
         "Settings Not Found")
        m_strSettingsPath = FilePath
        ''Show the settings form 
        ''Dim frmSettings As New Settings
        ''frmSettings.Show()

      End If

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





