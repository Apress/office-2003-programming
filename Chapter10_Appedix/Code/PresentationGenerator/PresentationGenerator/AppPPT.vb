Imports Microsoft.Office.Interop.PowerPoint
'TODO:  Update Chapter 2 to reflect:
'         1. the passed SettingsPath variable
Public Class AppPPT
  Private Shared WithEvents m_appPPT As Application

  Public Shared Function Setup(ByVal oApp As Application) As Boolean
    m_appPPT = oApp
  End Function

  Public Shared Function ShutDown()
    m_appPPT.Quit()
    m_appPPT = Nothing
  End Function


  Public Shared Property App() As Application
    Get
      App = m_appPPT
    End Get
    Set(ByVal Value As Application)
      m_appPPT = Value
    End Set
  End Property


  Private Shared Sub m_appPPT_AfterPresentationOpen(ByVal Pres As Presentation) _
    Handles m_appPPT.AfterPresentationOpen

    Dim strMacroName As New String("!StartPPTWizard")
    Dim aryParams() = {UserSettings.SettingsPath}

    strMacroName = strMacroName.Concat(m_appPPT.ActivePresentation.Name, _
      strMacroName.ToString)
    m_appPPT.Run(strMacroName.ToString, aryParams)

  End Sub
End Class
