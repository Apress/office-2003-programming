Imports Microsoft.Office.Core
imports Extensibility
imports System.Runtime.InteropServices


#Region " Read me for Add-in installation and setup information. "
' When run, the Add-in wizard prepared the registry for the Add-in.
' At a later time, if the Add-in becomes unavailable for reasons such as:
'   1) You moved this project to a computer other than which is was originally created on.
'   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
'   3) Registry corruption.
' you will need to re-register the Add-in by building the EmailTemplatesSetup project 
' by right clicking the project in the Solution Explorer, then choosing install.
#End Region

<GuidAttribute("ACD414A0-43D3-4963-AB9E-6669D087A056"), ProgIdAttribute("EmailTemplates.Connect")> _
Public Class Connect
	
	Implements Extensibility.IDTExtensibility2
	
  'Dim applicationObject As Object
  'Dim addInInstance As Object
	
	
	Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnBeginShutdown
	End Sub
	
	Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnAddInsUpdate
	End Sub
	
	Public Sub OnStartupComplete(ByRef custom As System.Array) Implements Extensibility.IDTExtensibility2.OnStartupComplete

  End Sub

  Public Sub OnDisconnection(ByVal RemoveMode As _
    Extensibility.ext_DisconnectMode, ByRef custom As System.Array) _
    Implements Extensibility.IDTExtensibility2.OnDisconnection

    appOutlook.ShutDown()
    UserSettings.SaveSettings()
  End Sub

  Public Sub OnConnection(ByVal application As Object _
    , ByVal connectMode As Extensibility.ext_ConnectMode, _
    ByVal addInInst As Object, ByRef custom As System.Array) _
    Implements Extensibility.IDTExtensibility2.OnConnection

    'applicationObject = application
    'addInInstance = addInInst

    UserSettings.LoadSettings(System.Windows.Forms.Application.StartupPath)
    appOutlook.Setup(application)



  End Sub
End Class
