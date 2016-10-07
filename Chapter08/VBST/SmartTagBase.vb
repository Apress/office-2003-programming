Imports System
Imports System.Threading
Imports System.Resources
Imports System.Reflection
Imports System.Globalization


''Implements the base functionality shared between ISmartTagRecognizer, ISmartTagRecognizer2, ISmartTagAction, ISmartTagAction2

Public Class SmartTagBase

  Friend _rm As New ResourceManager("VBST.SmartTag", _
    [Assembly].GetAssembly(GetType(SmartTagBase)))
  Protected Const tagNameExternal As String = _
    "http://schemas.microsoft.com/InformationBridge/2004#reference"

  Public Sub New()
    'blank
  End Sub

  Public Function get_SmartTagName(ByVal SmartTagID As Integer) As String
    Select Case SmartTagID
      Case 1
        Return tagNameExternal
      Case Else
        Return Nothing

    End Select
  End Function


  Public Function get_Desc(ByVal LocaleID As Integer) As String
    Try
      Return _rm.GetString("SmartTagDescription", _
        New CultureInfo(LocaleID))

    Catch ex As Exception
      Return Nothing
    End Try
  End Function

  Public ReadOnly Property SmartTagCount() As Integer
    Get
      Return 2
    End Get
  End Property

  Public ReadOnly Property ProgId()
    Get
      Return _rm.GetString("SmartTagProgID", _
        Thread.CurrentThread.CurrentCulture)

    End Get
  End Property

  Public Function get_Name(ByVal LocaleId As Integer) As String
    Try
      Dim a As String = _rm.GetString("SmartTagName", _
        New CultureInfo(LocaleId))
      Return _rm.GetString("SmartTagName", _
        New CultureInfo(LocaleId))

    Catch ex As Exception
      Return Nothing
    End Try
  End Function

End Class





