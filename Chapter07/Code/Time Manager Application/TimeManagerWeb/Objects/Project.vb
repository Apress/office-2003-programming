Public Class Project

    '***************************************************************************
    Public ProjectName As String

    '***************************************************************************
    Public Shared Function GetFromDR(ByRef dbDr As IDataReader) As Project

        Dim projectObj As New Project

        Try
            projectObj.ProjectName = dbDr("projectName")
        Catch ex As Exception
            projectObj = Nothing
        End Try

        Return projectObj

    End Function


End Class
