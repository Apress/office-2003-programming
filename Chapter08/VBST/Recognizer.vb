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

Public Class Recognizer
  Implements ISmartTagRecognizer
  Implements ISmartTagRecognizer2

  Protected Const tagNameExternal As String = _
  "http://schemas.microsoft.com/InformationBridge/2004#reference"

  Dim _xmlDocTerms As XmlDocument
#Region "Properties"

  Public ReadOnly Property Name(ByVal LocaleID As Integer) As String _
    Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Name
    Get
      Return "Bravo Corp Account Details Smart Tag"
    End Get
  End Property

  Public ReadOnly Property ProgId() As String _
    Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.ProgId
    Get
      Return "VBST.Recognizer"
    End Get
  End Property

  Public ReadOnly Property Desc(ByVal LocaleID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Desc
    Get
      Return "Bravo Corp Account Details Smart Tag"
    End Get
  End Property



  Public ReadOnly Property SmartTagCount() As Integer Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagCount
    Get
      Return 1
    End Get
  End Property

  Public ReadOnly Property SmartTagName(ByVal SmartTagID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagName
    Get
      If (SmartTagID = 1) Then
        SmartTagName = tagNameExternal
      End If
    End Get
  End Property


  Public ReadOnly Property SmartTagDownloadURL(ByVal SmartTagID As Integer) As String Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.SmartTagDownloadURL
    Get
      Return "http://msdn.microsoft.com/ibframework"

    End Get
  End Property

  Public ReadOnly Property PropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) As Boolean Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.PropertyPage
    Get

    End Get
  End Property
#End Region

  Public Sub Recognize(ByVal Text As String, ByVal DataType As Microsoft.Office.Interop.SmartTag.IF_TYPE, ByVal LocaleID As Integer, ByVal RecognizerSite As Microsoft.Office.Interop.SmartTag.ISmartTagRecognizerSite) Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer.Recognize

  End Sub

  Public Sub DisplayPropertyPage(ByVal SmartTagID As Integer, ByVal LocaleID As Integer) Implements Microsoft.Office.Interop.SmartTag.ISmartTagRecognizer2.DisplayPropertyPage

  End Sub

  Public Sub Recognize2(ByVal Text As String, _
    ByVal DataType As Microsoft.Office.Interop.SmartTag.IF_TYPE, _
    ByVal LocaleID As Integer, ByVal RecognizerSite2 As _
    Microsoft.Office.Interop.SmartTag.ISmartTagRecognizerSite2, _
    ByVal ApplicationName As _
    String, ByVal TokenList As _
    Microsoft.Office.Interop.SmartTag.ISmartTagTokenList) _
    Implements Microsoft.Office.Interop.SmartTag. _
    ISmartTagRecognizer2.Recognize2


    'Set the culture info for string comparisions
    'Thread.CurrentThread.CurrentCulture = New CultureInfo(LocaleID)
    Dim termItem As String = ""
    Dim nd As XmlNode
    For Each nd In _xmlDocTerms.SelectNodes("//Account")

      termItem = nd.InnerText
      'Implement a regular expression to determine if a term match occurred
      Dim r As Regex = New Regex(termItem, RegexOptions.IgnoreCase)
      Dim m As Match = r.Match(Text)

      'If a match is found, build a context command
      If m.Success Then

        'IBF HOL 4,5,6
        '==================
        ''Dim formatString As String = "<ContextInformation xmlns=""http://schemas.microsoft.com/InformationBridge/2004/ContextInformation"" " & _
        ''  "MetadataScopeName=""http://WebServices/CRM"" " & _
        ''"EntityName=""Account"" ViewName=""AccountDefault"">" & _
        ''"<Reference> " & _
        ''"<AccountID xmlns='urn-IBFHOL-CRM' ID='Contoso'" & _
        ''" iwb:MetadataScopeName='http://WebServices/CRM' " & _
        ''" xmlns:iwb='http://schemas.microsoft.com/InformationBridge/2004'" & _
        ''" iwb:EntityName='Account' iwb:ViewName='AccountDefault'/>" & _
        ''"</Reference>" & _
        ''"</ContextInformation>"

        ''IBF Sample Solution
        '==================
        Dim formatString As String = "<ContextInformation  " & _
          "xmlns=""http://schemas.microsoft.com/InformationBridge" & _
          "/2004/ContextInformation"" " & _
          " MetadataScopeName=""http://InformationBridge/Sample""" & _
          " EntityName=""Account"" ViewName=""AccountDefault""> " & _
          " <Reference> " & _
          " <AccountName xmlns=""urn-SampleSolution-Data"" " & _
          " ID=""{0}"" iwb:MetadataScopeName=" & _
          """http://InformationBridge/Sample"" " & _
          " xmlns:iwb=""http://schemas.microsoft.com/InformationBridge/2004"" " & _
          " iwb:EntityName=""Account"" iwb:ViewName=""AccountDefault"" />" & _
          "</Reference></ContextInformation>"



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
      Path.GetDirectoryName([Assembly].GetExecutingAssembly().CodeBase). _
      Replace("file:\", "")
    Dim sFile As String = sFolder & "\CustomerNames.xml"
    _xmlDocTerms = New XmlDocument

    Try
      _xmlDocTerms.Load(sFile)
    Catch ex As Exception
      System.Windows.Forms.MessageBox.Show(ex.Message, _
        "Information Bridge Framework Smart Tag")
    End Try
  End Sub
End Class
