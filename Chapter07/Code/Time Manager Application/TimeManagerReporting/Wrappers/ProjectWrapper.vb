Public Class ProjectWrapper

  Public ProjectObj As TMW.Project

  Sub New(ByVal ProjectObj As TMW.Project)
    Me.ProjectObj = ProjectObj
  End Sub

  Public Overrides Function toString() As String
    Return ProjectObj.ProjectName
  End Function


End Class
