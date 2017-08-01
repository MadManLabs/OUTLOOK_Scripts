Option Explicit

Public Sub HelloWorld(item As Outlook.MailItem)
    MsgBox "Hello World"
End Sub

'This rule runs scripts that exit out of SAP
'and restart running the scripts.
'Allows me to remotely restart the scripts.
Public Sub RestartSAPScripts(item As Outlook.MailItem)
    Dim PauseTime, Start As Double
    'Close all command windows
    Call KillcmdWindows
    PauseTime = 1 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Kill any SAP error windows
    Call KillSAPErrorWindow
    
    'Wait for current scripts to complete
    PauseTime = 10 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Close all SAP session windows
    Call CloseSAPWindows

    PauseTime = 10 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Close all error windows if there are any
    Call KillScriptErrorWindows
    
    'Wait for current scripts to complete
    PauseTime = 2 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Open SAP P01
    Call SAPClickP01
    
    'Wait for SAP window to open
    PauseTime = 15 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Login on my SAP account
    Call SAPLogon
    
    'Wait for SAP window to open
    PauseTime = 5 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Restart SAP scripts
    Call StartSAPBatchScripts
        
    'Do some error checking to verify scripts are running successfully
    'Email myself to notify that the script was successfull or not
    
End Sub

Public Sub EndScriptsLogout(item As Outlook.MailItem)
    Dim PauseTime, Start As Double

    'Close all command windows
    Call KillcmdWindows
    
    'Wait for current scripts to complete
    PauseTime = 1 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Kill any SAP error windows
    Call KillSAPErrorWindow
    
    'Wait for current scripts to complete
    PauseTime = 10 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Close all SAP session windows
    Call CloseSAPWindows

    PauseTime = 10 'seconds
    Start = Timer
    Do While Timer < Start + PauseTime
        DoEvents 'yield to other processes
    Loop
    
    'Close all error windows if there are any
    Call KillScriptErrorWindows
    
    
End Sub
