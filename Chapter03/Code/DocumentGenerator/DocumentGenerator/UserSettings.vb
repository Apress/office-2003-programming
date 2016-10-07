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
  Private Shared m_strEntryID As String
  Private Shared m_strStoreID As String
  Private Shared m_strClientsFolderPath As String
  Private Const SETTINGS_XML_FILE_NAME As String = "\docgen.xml"

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

        .WriteStartDocument(True)
        .WriteComment("UUser Settings Document Generator Add-In")
        .WriteStartElement("BravoDocGenUserSettings") '
        .WriteAttributeString("UserName", UserSettings.UserName)
        .WriteAttributeString("Password", UserSettings.Password)
        .WriteAttributeString("SaveFolder", UserSettings.SaveFolder)
        .WriteAttributeString("TemplatesFolder", UserSettings.TemplatesFolder)
        .WriteAttributeString("ServerName", UserSettings.ServerName)
        .WriteAttributeString("DatabaseName", UserSettings.DatabaseName)
        .WriteAttributeString("ClientsFolderPath", UserSettings.ClientsFolderPath)
        .WriteAttributeString("EntryID", UserSettings.EntryID)
        .WriteAttributeString("StoreID", UserSettings.StoreID)
        .WriteEndElement()

        .WriteEndDocument()
        .Close()
      End With

      Return True
    Catch ex As Exception
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
          m_strClientsFolderPath = .GetAttribute("ClientsFolderPath")
          m_strEntryID = .GetAttribute("EntryID")
          m_strStoreID = .GetAttribute("StoreID")
          .Close()

          Return True
        End With
      Else
        'Settings do not exist so alert the user...
        MsgBox("The Document Generator settings file does not exist." & _
         vbCrLf & vbCrLf & _
         "Please set your settings now.", MsgBoxStyle.Information, _
         "Settings Not Found")
        m_strSettingsPath = FilePath
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

  Public Shared Property ClientsFolderPath() As String
    Get
      Return m_strClientsFolderPath
    End Get
    Set(ByVal Value As String)
      m_strClientsFolderPath = Value
    End Set
  End Property

  Public Shared Property EntryID() As String
    Get
      Return m_strEntryID
    End Get
    Set(ByVal Value As String)
      m_strEntryID = Value
    End Set
  End Property

  Public Shared Property StoreID() As String
    Get
      Return m_strStoreID
    End Get
    Set(ByVal Value As String)
      m_strStoreID = Value
    End Set
  End Property


#End Region


End Class





