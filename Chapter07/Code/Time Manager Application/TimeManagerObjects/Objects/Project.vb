Public Class Project

    '***************************************************************************
    Public ProjectName As String

    '***************************************************************************
    Public Function Add() As Boolean
        Try
            Dim TimeManager As New TMW.TimeManagerWeb
            Return TimeManager.TestConnection()
            Return TimeManager.AddProject(ConvertTo)
        Catch ex As Exception
            Return False
        End Try
    End Function

    '***************************************************************************
    Public Shared Function ConvertFrom(ByRef obj As TMW.Project) As Project
        Dim projectObj As New Project
        projectObj.ProjectName = obj.ProjectName
    End Function

    '***************************************************************************
    Public Function ConvertTo() As TMW.Project
        Dim obj As New TMW.Project
        obj.ProjectName = Me.ProjectName
        Return obj
    End Function

End Class
