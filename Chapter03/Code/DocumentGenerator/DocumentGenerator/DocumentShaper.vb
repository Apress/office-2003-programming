Imports Microsoft.Office.Interop.Word
Imports OL = Microsoft.Office.Interop.Outlook

Public Class appWord
  Private WithEvents m_appWord As Application
  Private m_Doc As Document
  Private m_Contact As OL.ContactItem

  Public Function Setup(ByVal oApp As Application) As Boolean
    m_appWord = oApp
  End Function

  Public Function ShutDown()
    m_appWord.Quit()
    m_appWord = Nothing
  End Function

 




End Class


