Imports System
Imports System.IO
Imports System.Xml
Imports System.Resources
Imports System.Threading
Imports System.Reflection
Imports System.Collections
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.SmartTag



Public Class Recognizer_backup
  'Inherits VBST.SmartTagBase
  Implements ISmartTagRecognizer
  Implements ISmartTagRecognizer2

  Friend _rm As New ResourceManager("VBST.SmartTag", _
     [Assembly].GetAssembly(GetType(SmartTagBase)))
  Protected Const tagNameExternal As String = _
    "http://schemas.microsoft.com/InformationBridge/2004#reference"

  Dim _xmlDocTerms As XmlDocument

#Region "ISmartTagRecognizer Members"
  Public Shadows ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Desc
    Get
      'Return "The Chapter 8 IBF Smart Tag"
      Return _rm.GetString("SmartTagDescription", _
        New CultureInfo(LocaleID))
    End Get
  End Property

  Public Shadows ReadOnly Property Name(ByVal LocaleID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Name
    Get
      'Return "Chp 08 IBF Smart Tag"
      Dim a As String = _rm.GetString("SmartTagName", New CultureInfo(LocaleID))
      Return _rm.GetString("SmartTagName", New CultureInfo(LocaleID))
    End Get
  End Property

  Public ReadOnly Property ProgId1() As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.ProgId
    Get
      Return "VBST.Recognizer"
      'Return _rm.GetString("SmartTagProgID", Thread.CurrentThread.CurrentCulture)
    End Get
  End Property

  Public Sub Recognize(ByVal Text As String, ByVal DataType As Microsoft.Office.Interop.SmartTag.IF_TYPE, ByVal LocaleID As Integer, ByVal RecognizerSite As Microsoft.Office.Interop.SmartTag.ISmartTagRecognizerSite) Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Recognize

  End Sub

  Public ReadOnly Property SmartTagCount1() As Integer Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagCount
    Get
      Return 2
    End Get
  End Property

  Public ReadOnly Property SmartTagDownloadURL(ByVal SmartTagID As _
    Integer) As String Implements _
    Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagDownloadURL
    Get
      Return _rm.GetString("SmartTagURL", _
        Thread.CurrentThread.CurrentCulture)
    End Get
  End Property

  Public Shadows ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagName
    Get
      If (SmartTagID = 1) Then
        SmartTagName = "http://schemas.microsoft.com/InformationBridge/2004#reference"
      End If
    End Get
  End Property
#End Region

#Region "ISmartTagRecognizer2 Members"
  Public Sub DisplayPropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.DisplayPropertyPage

  End Sub

  Public ReadOnly Property PropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) As Boolean Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.PropertyPage
    Get

    End Get
  End Property

  Public Sub Recognize2(ByVal Text As String, ByVal DataType As _
    Microsoft.Office.Interop.SmartTag.IF_TYPE, ByVal LocaleID As _
    Integer, ByVal RecognizerSite2 As _
    Microsoft.Office.Interop.SmartTag.ISmartTagRecognizerSite2, _
    ByVal ApplicationName As String, ByVal TokenList As _
    Microsoft.Office.Interop.SmartTag.ISmartTagTokenList) _
    Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.Recognize2

    'Set the culture info for string comparisions
    Thread.CurrentThread.CurrentCulture = New CultureInfo(LocaleID)
    Dim termItem As String = ""
    Dim nd As XmlNode
    For Each nd In _xmlDocTerms.SelectNodes("//Account")

      termItem = nd.InnerText
      'Implement a regular expression to determine if a term match occurred
      Dim r As Regex = New Regex(termItem, RegexOptions.IgnoreCase)
      Dim m As Match = r.Match(Text)

      'If a match is found, build a context command
      If m.Success Then
        Dim formatString As String = "<?xml version=\""1.0\""?>" & _
          "<ContextInformation xmlns:xsd=\""http://www.w3.org/2001" & _
            "/XMLSchema\"" xmlns:xsi=\""http://www.w3.org/2001/XMLSchema" & _
            "-instance\"" MetadataScopeName=\""http://WebServices" & _
            "/CRM\"" EntityName=\""Account\"" ViewName=\""Account" & _
            "Default\"" xmlns=\""http://schemas.microsoft.com/" & _
            "InformationBridge/2004/ContextInformation\"">" & _
         "<Reference>" & _
         "<AccountID xmlns='urn-IBFHOL-CRM' ID='{0}' " & _
          "iwb:MetadataScopeName='http://WebServices/CRM' xmlns:iwb=" & _
          "'http://schemas.microsoft.com/InformationBridge/2004\' " & _
          "iwb:EntityName='Account' iwb:ViewName='AccountDefault' />" & _
         "</Reference>" & _
         "</ContextInformation>"

        'Insert the term into the context string
        Dim context As String = String.Format(formatString, m.Value)

        'Add the context to a property bag 
        Dim propBag As ISmartTagProperties = _
          RecognizerSite2.GetNewPropertyBag()
        propBag.Write("data", context)

        'add the smart tag
        RecognizerSite2.CommitSmartTag(tagNameExternal, m.Index + 1, _
         m.Length, propBag)



      End If
    Next
  End Sub

  Public Sub SmartTagInitialize(ByVal ApplicationName As String) _
    Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.SmartTagInitialize

    Dim sFolder As String = _
      Path.GetDirectoryName([Assembly].GetExecutingAssembly().CodeBase).Replace("file:\", "")
    Dim sFile As String = sFolder & "\IBFHOLSmartTagTerms.xml"
    _xmlDocTerms = New XmlDocument

    Try
      _xmlDocTerms.Load(sFile)
    Catch ex As Exception
      System.Windows.Forms.MessageBox.Show(ex.Message, _
        "Information Bridge Framework Smart Tag")
    End Try

  End Sub

#End Region

End Class

