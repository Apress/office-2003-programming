Imports System.Web.Services
Imports System.Data.OleDb

<System.Web.Services.WebService(Namespace:="http://tempuri.org/TimeManagerWeb/TimeManagerWeb")> _
Public Class TimeManagerWeb
    Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    '***************************************************************************
    <WebMethod(Description:="Save an individual hour record")> _
    Public Function SaveHourObj(ByVal obj As tmHours) As Boolean
        Return obj.Save()
    End Function

    '***************************************************************************
    <WebMethod(Description:="Saves all the Hour objects in an array list.")> _
    Public Function SaveHourArrayList(ByVal userId As String, ByVal startDate As Date, ByVal endDate As Date, ByVal objCol() As Object) As Boolean

        Try
            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbTran As IDbTransaction = dbConn.BeginTransaction()
            Dim HasError As Boolean = False

            If Not tmHours.DeleteForGivenWeek(userId, startDate, endDate, dbConn, dbTran) Then
                HasError = True
            End If

            For Each obj As tmHours In objCol
                If Not HasError Then
                    If Not obj.Save(dbConn, dbTran) Then
                        HasError = True
                    End If
                End If
            Next

            If Not HasError Then
                dbTran.Commit()
                dbConn.Close()
                Return True
            Else
                dbTran.Rollback()
                dbConn.Close()
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Tests the database connection settings.  True indicates that a successful connection was made.  False indicates there was an error opening the connection.")> _
    Public Function TestConnection() As String
        Try
            Dim dbConn As OleDbConnection = GetConnection()
            dbConn.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    '***************************************************************************
    <WebMethod(Description:="Login will return a user object given a userId and a password.  If the userId and password do not match a user in the database, nothing will be returned.")> _
    Public Function Login(ByVal userId As String, ByVal password As String) As tmUser
        Try
            Dim tempObj As New tmUser
            Dim dbConn As OleDbConnection = GetConnection()
            Dim dbCmd As New OleDbCommand("SELECT * FROM [tblUser] WHERE [userId]='" & sqlString(userId) & "' AND [password]='" & sqlString(password) & "'", dbConn)
            Dim dbDr As OleDbDataReader = dbCmd.ExecuteReader(CommandBehavior.SingleRow)
            If dbDr.Read Then
                tempObj = tmUser.GetFromDR(dbDr)
            Else
                tempObj = Nothing
            End If
            dbDr.Close()
            dbConn.Close()
            Return tempObj
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '***************************************************************************
    <WebMethod(Description:="Returns a string array containing a list of the projects.")> _
    Public Function GetAllProjects() As Project()

        Dim returnArray As Project() = Nothing
        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand
            Dim dbDr As IDataReader
            Dim projectObj As Project

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "SELECT * FROM [tblProject] ORDER BY [projectName];"
            dbDr = dbCmd.ExecuteReader(CommandBehavior.SequentialAccess)

            While dbDr.Read
                projectObj = Project.GetFromDR(dbDr)
                If Not projectObj Is Nothing Then
                    If returnArray Is Nothing Then
                        ReDim returnArray(0)
                    Else
                        ReDim Preserve returnArray(returnArray.Length)
                    End If
                    returnArray(returnArray.Length - 1) = projectObj
                End If
            End While

            dbDr.Close()
            dbConn.Close()
            Return returnArray

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Adds a new project to the project table")> _
    Public Function AddProject(ByRef ProjectObj As Project) As Boolean

        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "INSERT INTO [tblProject](projectName) VALUES ('" & sqlString(ProjectObj.ProjectName) & "');"
            If dbCmd.ExecuteNonQuery = 1 Then
                AddProject = True
            Else
                AddProject = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try


    End Function

    '***************************************************************************
    <WebMethod(Description:="Returns an array containing all of the users")> _
    Public Function GetAllUsers() As tmUser()

        Dim returnArray As tmUser()
        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand
            Dim dbDr As IDataReader
            Dim userObj As tmUser

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "SELECT * FROM [tblUser] ORDER BY [nameLast], [nameFirst] ASC;"
            dbDr = dbCmd.ExecuteReader(CommandBehavior.SequentialAccess)

            While dbDr.Read
                userObj = tmUser.GetFromDR(dbDr)
                If Not userObj Is Nothing Then
                    If returnArray Is Nothing Then
                        ReDim returnArray(0)
                    Else
                        ReDim Preserve returnArray(returnArray.Length)
                    End If
                    returnArray(returnArray.Length - 1) = userObj
                End If
            End While

            dbDr.Close()
            dbConn.Close()
            Return returnArray

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Adds the specified user")> _
    Public Function AddUser(ByVal userObj As tmUser) As Boolean

        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "INSERT INTO [tblUser]([userId],[password],[nameLast],[nameFirst],[admin]) VALUES ('" & sqlString(userObj.userId) & "','" & sqlString(userObj.password) & "','" & sqlString(userObj.nameLast) & "','" & sqlString(userObj.nameFirst) & "'," & IIf(userObj.admin, "true", "false") & ");"
            If dbCmd.ExecuteNonQuery = 1 Then
                AddUser = True
            Else
                AddUser = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Adds the specified user")> _
    Public Function UpdateUser(ByVal originalUserId As String, ByVal userObj As tmUser) As Boolean

        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "UPDATE [tblUser] SET [userId]='" & sqlString(userObj.userId) & "', [password]='" & sqlString(userObj.password) & "', [nameLast]='" & sqlString(userObj.nameLast) & "', nameFirst='" & sqlString(userObj.nameFirst) & "', admin=" & IIf(userObj.admin, "true", "false") & " WHERE [userId]='" & sqlString(originalUserId) & "';"
            If dbCmd.ExecuteNonQuery = 1 Then
                UpdateUser = True
            Else
                UpdateUser = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Deletes the specified user")> _
    Public Function DeleteUser(ByVal userObj As tmUser) As Boolean

        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "DELETE FROM [tblUser] WHERE [userId]='" & sqlString(userObj.userId) & "';"
            If dbCmd.ExecuteNonQuery = 1 Then
                DeleteUser = True
            Else
                DeleteUser = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try

    End Function

    '***************************************************************************
    <WebMethod(Description:="Deletes the specified project")> _
    Public Function DeleteProject(ByVal ProjectObj As Project) As Boolean
        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "DELETE FROM [tblProject] WHERE [projectName]='" & sqlString(ProjectObj.ProjectName) & "';"
            If dbCmd.ExecuteNonQuery = 1 Then
                DeleteProject = True
            Else
                DeleteProject = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try
    End Function

    '***************************************************************************
    <WebMethod(Description:="Updates the specified project.")> _
    Public Function UpdateProject(ByVal OriginalProjectName As String, ByVal ProjectObj As Project) As Boolean
        Try

            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand

            dbCmd = dbConn.CreateCommand()

            'Insert the new items
            dbCmd.CommandText = "UPDATE [tblProject] SET [projectName]='" & sqlString(ProjectObj.ProjectName) & "' WHERE [projectName]='" & sqlString(OriginalProjectName) & "';"
            If dbCmd.ExecuteNonQuery = 1 Then
                UpdateProject = True
            Else
                UpdateProject = False
            End If

            dbConn.Close()

        Catch ex As Exception
            Return False
        End Try
    End Function

    '***************************************************************************
    <WebMethod(Description:="Returns project names and hour information.")> _
    Public Function GetProjectReportInfoByEmployee(ByVal userId As String, ByVal StartDate As Date, ByVal EndDate As Date) As tmHours()

        Try
            Dim returnArray As tmHours()
            Dim obj As tmHours
            Dim dbConn As IDbConnection = Globals.GetConnection()
            Dim dbCmd As IDbCommand = dbConn.CreateCommand
            Dim dbDr As IDataReader

            Dim SQL As String = "SELECT * FROM [tblHours] WHERE [userId]='" & sqlString(userId) & "'"

            If Not StartDate = Nothing And Not EndDate = Nothing Then
                'Use both start date and enddate
                SQL &= " AND [startDate]>=#" & Format(StartDate, "MM/dd/yyyy") & "#  AND [startDate]<=#" & Format(EndDate, "MM/dd/yyyy") & "#"
            ElseIf StartDate = Nothing And Not EndDate = Nothing Then
                'Only use startdate
                SQL &= " AND [startDate]<=#" & Format(EndDate, "MM/dd/yyyy") & "#"
            ElseIf Not StartDate = Nothing And EndDate = Nothing Then
                'Only use enddate                
                SQL &= " AND [startDate]>=#" & Format(StartDate, "MM/dd/yyyy") & "#"
            End If

            SQL &= " ORDER BY [startDate], [projectName];"

            dbCmd.CommandText = SQL
            dbDr = dbCmd.ExecuteReader(CommandBehavior.SequentialAccess)
            While dbDr.Read()
                obj = tmHours.GetFromDR(dbDr)
                If Not obj Is Nothing Then
                    If returnArray Is Nothing Then
                        ReDim returnArray(0)
                    Else
                        ReDim Preserve returnArray(returnArray.Length)
                    End If
                    returnArray(returnArray.Length - 1) = obj
                End If
            End While
            dbDr.Close()
            dbConn.Close()
            Return returnArray

        Catch ex As Exception
            Return Nothing
        End Try






    End Function


End Class
