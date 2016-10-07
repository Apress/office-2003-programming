Public Class UserWrapper

  Public UserObj As TMW.tmUser

  Sub New(ByVal UserObj As TMW.tmUser)
    Me.UserObj = UserObj
  End Sub


  Public Overrides Function toString() As String
    Return UserObj.nameLast & ", " & UserObj.nameFirst
  End Function


End Class
