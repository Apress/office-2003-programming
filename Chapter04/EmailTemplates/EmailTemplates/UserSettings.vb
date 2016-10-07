Imports System.Xml
Imports System.Environment
Imports System.IO

Public Class UserSettings
  Private Shared m_strTemplatesFolder As String
  Private Shared m_strSettingsPath As String
  Private Shared m_strEntryID As String
  Private Shared m_strStoreID As String
  Private Shared m_strUserName As String

  Private Const SETTINGS_XML_FILE_NAME As String = "\emailtemps.xml"

  Public Shared Event AfterSettingsChange(ByVal EntryID As String, ByVal StoreID As String)

  'This is here only to declare and expose one instance event.  
  'This will allow other classes to respond to the AfterSettingsChange event.   
  Public Event NadaMucho()


#Region "Public Shared Functions"

  Public Shared Function SaveSettings() As Boolean
    Try


      Dim strPath As String
      strPath = strPath.Concat(UserSettings.SettingsPath, _
        SETTINGS_XML_FILE_NAME)
      Dim xtwSettings As New XmlTextWriter(strPath.ToString, _
        System.Text.Encoding.UTF8)

      With xtwSettings
        .Formatting = Formatting.Indented
        .Indentation = 2
        .QuoteChar = """"c

        .WriteStartDocument(True)
        .WriteComment("User Settings from the Bravo Email Templates Add-In")
        .WriteStartElement("BravoEmailTemplatesUserSettings")
        .WriteAttributeString("TemplatesFolderPath", _
          UserSettings.TemplatesFolderPath)
        .WriteAttributeString("EntryID", UserSettings.EntryID)
        .WriteAttributeString("StoreID", UserSettings.StoreID)
        .WriteAttributeString("UserName", UserSettings.UserName)
        .WriteEndElement()
        .WriteEndDocument()
        .Close()
      End With

      RaiseEvent AfterSettingsChange(UserSettings.EntryID, _
        UserSettings.StoreID)

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
          m_strTemplatesFolder = .GetAttribute("TemplatesFolderPath")
          m_strEntryID = .GetAttribute("EntryID")
          m_strStoreID = .GetAttribute("StoreID")
          m_strUserName = .GetAttribute("UserName")
          .Close()

          Return True
        End With
      Else
        'Settings do not exist so alert the user...
        MsgBox("The Email Templates Engine's settings file does not exist." & _
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

  Public Shared Property UserName() As String
    Get
      Return m_strUserName
    End Get
    Set(ByVal Value As String)
      m_strUserName = Value
    End Set
  End Property

  Public Shared ReadOnly Property SettingsPath() As String
    Get
      Return m_strSettingsPath
    End Get
    'Set(ByVal strValue As String)
    '  m_strSettingsPath = strValue
    'End Set
  End Property

  Public Shared Property TemplatesFolderPath() As String
    Get
      Return m_strTemplatesFolder
    End Get
    Set(ByVal Value As String)
      m_strTemplatesFolder = Value
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





