Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class AppPPT
  Private Shared WithEvents m_appPPT As PowerPoint.Application

  Public Shared Function Setup(ByVal oApp As PowerPoint.Application) As Boolean
    'store reference to PowerPoint in class variable
    m_appPPT = oApp
  End Function

  Public Shared Function ShutDown()
    'Close Powerpoint and clear it's reference
    m_appPPT.Quit()
    m_appPPT = Nothing
  End Function


  Public Shared Property App() As PowerPoint.Application
    Get
      App = m_appPPT
    End Get
    Set(ByVal Value As PowerPoint.Application)
      m_appPPT = Value
    End Set
  End Property


  Private Shared Sub m_appPPT_AfterPresentationOpen(ByVal Pres _
    As PowerPoint.Presentation) Handles m_appPPT.AfterPresentationOpen

    'Store the name of the VBA macro
    'contained in the Template
    Dim strMacroName As New String("!StartPPTWizard")
    'Create variable containing the location of the 
    'settings file
    Dim aryParams() = {UserSettings.SettingsPath}

    'combine the VBA macro name with the 
    '.PPT file name in order to call the macro
    strMacroName = strMacroName.Concat(m_appPPT.ActivePresentation.Name, _
      strMacroName.ToString)
    'Call the macro
    m_appPPT.Run(strMacroName.ToString, aryParams)

  End Sub
End Class
