Option Explicit

Public Sub CloseSAPWindows()
    Dim Application As Variant
    Dim SapGuiAuto As Variant
    Dim Connection As Variant
    Dim Session As Variant
    Dim Wscript As Variant
    Dim i As Integer
    Dim PauseTime, Start As Double
    
On Error GoTo Err_Handler
    If Not IsObject(Application) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Application.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If
    If IsObject(Wscript) Then
       Wscript.ConnectObject Session, "on"
       Wscript.ConnectObject Application, "on"
    End If
    
    'Stop transactions
    Session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    'Close the window
    Session.findById("wnd[0]").Close
    
    

    Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    For i = 0 To 6
        Set Session = Connection.Children(0)
        Session.findById("wnd[0]").Close
        Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    Next

    
Err_Exit:
    Set SapGuiAuto = Nothing
    Set Application = Nothing
    Set Connection = Nothing
    Set Session = Nothing
    Exit Sub
    
Err_Handler:
    If Err.Number = 614 Then 'error number for "the enumerator of the collection cannot find element"
                             'occurs when sap is not open
        Resume Err_Exit
        
    ElseIf Err.Number = 65532 Then 'error is the application server is busy
        PauseTime = 15 'seconds
        Start = Timer
        Do While Timer < Start + PauseTime
            DoEvents 'yield to other processes
        Loop
        Resume Next
    ElseIf Err.Number = -2147417848 Then 'error occurs if the object has been disconnected from clients
        Resume Next
    Else
        MsgBox Err.Description & " " & Err.Number
        Resume Next
    End If
End Sub

Public Sub SAPClickP01()
    Dim x As Variant
    x = Shell("C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe " & Chr(34) & "T:\_Shared_Workspace\Joe_Dyke\NCM and Aged WIP\QN Database\Live Data Scripts\SAPstart and setup.ahk" & Chr(34), vbNormalFocus)
End Sub

Public Sub SAPLogon()
    Dim Application As Variant
    Dim SapGuiAuto As Variant
    Dim Connection As Variant
    Dim Session As Variant
    Dim Wscript As Variant
    Dim i As Integer
    
    If Not IsObject(Application) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set Application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Application.Children(0)
    End If
    If Not IsObject(Session) Then
       Set Session = Connection.Children(0)
    End If
    If IsObject(Wscript) Then
       Wscript.ConnectObject Session, "on"
       Wscript.ConnectObject Application, "on"
    End If
    
    'Enter login info
    Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "8058649"
    Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Mju78ik,"
    Session.findById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
    Session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
    Session.findById("wnd[0]").sendVKey 0
    
    'Try error handling to deal with month end pop up box
    'and if any other instances of SAP are open
    On Error Resume Next
    Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
    Session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    Session.findById("wnd[1]/tbar[0]/btn[12]").press
    Session.findById("wnd[1]/usr/btnBUTTON_1").press
    
    If Err.Number <> 0 Then
        Session.findById("wnd[1]/usr/btnBUTTON_1").press
        Err.Clear
    End If
    
    On Error GoTo 0
    
    Session.findById("wnd[0]").resizeWorkingPane 82, 38, False
    
    Session.createSession
    On Error Resume Next
    Set Session = Connection.Children(0)
    Do While Err.Number <> 0
        Wscript.Sleep 500
        Set Session = Connection.Children(0)
        Err.Clear
    Loop
    On Error GoTo 0
    
    Session.findById("wnd[0]").resizeWorkingPane 82, 38, False
    Session.createSession

End Sub

Public Sub KillcmdWindows()
On Error GoTo Err_Handler
    Shell "taskkill /IM cmd.exe /FI ""WINDOWTITLE eq RunSAPReportsHUMS.bat - Shortcut""", 0
    Shell "taskkill /IM cmd.exe /FI ""WINDOWTITLE eq RunSAPReportsCOMP.bat - Shortcut""", 0
    Shell "taskkill /IM cmd.exe /FI ""WINDOWTITLE eq c:\Windows\system32\cmd.exe""", 0

Err_Handler:
    Resume Next
End Sub

Public Sub KillScriptErrorWindows()
    Shell "taskkill /IM Wscript.exe /FI ""WINDOWTITLE eq Windows Script Host""", 0

End Sub

Public Sub KillSAPErrorWindow()
    Shell "taskkill /IM saplogon.exe /FI ""WINDOWTITLE eq SAP GUI for Windows 740""", 0

End Sub

Public Sub StartSAPBatchScripts()
On Error GoTo Err_Handler
    Dim x As Variant
    Dim y As Variant

    x = Shell(Chr(34) & "T:\_Shared_Workspace\CI\Events\SMART Factory\HUMS\Scripts\RunSAPReports.bat" & Chr(34), vbNormalFocus)
    y = Shell(Chr(34) & "T:\_Shared_Workspace\CI\Events\SMART Factory\Computers\Scripts\RunSAPReports.bat" & Chr(34), vbNormalFocus)
    
Err_Handler:
    Resume Next
End Sub
