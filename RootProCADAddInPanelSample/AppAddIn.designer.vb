﻿'------------------------------------------------------------------------------
' <auto-generated>
'     このコードはツールによって生成されました。
'     ランタイム バージョン:2.0.50727.5477
'
'     このファイルへの変更は、以下の状況下で不正な動作の原因になったり、
'     コードが再生成されるときに損失したりします。
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On



'''
<Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Applications.ProgrammingModel.dll", "9.0.0.0"),  _
 Global.System.CLSCompliantAttribute(false)>  _
Partial Public NotInheritable Class AppAddIn
    Inherits RootPro.RootProCAD.ApplicationEntryPoint
    Implements Microsoft.VisualStudio.Tools.Applications.Runtime.IEntryPoint
    
    Public Event Startup As System.EventHandler
    
    Public Event Shutdown As System.EventHandler
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub OnStartup()
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub FinishInitialization()
        MyBase.FinishInitialization
        Me.OnStartup
        If (Not (Me.StartupEvent) Is Nothing) Then
            RaiseEvent Startup(Me, System.EventArgs.Empty)
        End If
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub OnShutdown()
        If (Not (Me.ShutdownEvent) Is Nothing) Then
            RaiseEvent Shutdown(Me, System.EventArgs.Empty)
        End If
        MyBase.OnShutdown
    End Sub
    
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides ReadOnly Property PrimaryCookie() As String
        Get
            Return "AppAddIn"
        End Get
    End Property
End Class
