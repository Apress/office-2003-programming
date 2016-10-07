VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9105
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10755
   _ExtentX        =   18971
   _ExtentY        =   16060
   _Version        =   393216
   Description     =   "Shim for the Bravo Corp Presentation Generator"
   DisplayName     =   "PresentationGeneratorShim"
   AppName         =   "Microsoft PowerPoint"
   AppVer          =   "Microsoft PowerPoint 11.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim managedAddIn As IDTExtensibility2
Private Sub AddinInstance_Initialize()

    Set managedAddIn = CreateObject("PresentationGenerator.Connect")

End Sub

Private Sub AddinInstance_OnAddInsUpdate(custom() As Variant)
    
    managedAddIn.OnAddInsUpdate custom

End Sub

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

    managedAddIn.OnBeginShutdown custom

End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    managedAddIn.OnConnection Application, ConnectMode, AddInInst, custom

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    managedAddIn.OnDisconnection RemoveMode, custom
    
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)

    managedAddIn.OnStartupComplete custom

End Sub
