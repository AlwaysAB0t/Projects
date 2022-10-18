Dim objShell
Dim result
Dim warning
warning = MsgBox("Beware you might have to do this script 2 times, for some reason sometimes it won't close", 48 , "WARNING")
result = msgbox("Select Yes to deactivate the monitoring system. Or select No to enable the monitoring system", 4 , "Toggle the monitoring system made by noCOM")
If result=6 then

KillProc "WinNetRef.exe"
Sub KillProc( myProcess )    
Dim blnRunning, colProcesses, objProcess    
blnRunning = False    
Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process", , 48 )    
For Each objProcess in colProcesses        
If LCase( myProcess ) = LCase( objProcess.Name ) Then             ' Confirm that the process was actually running             blnRunning = True             ' Get exact case for the actual process name            
myProcess  = objProcess.Name             ' Kill all instances of the process            
objProcess.Terminate()        
End If    
Next    
If blnRunning Then        
Do Until Not blnRunning            
Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process Where Name = '"& myProcess & "'" )            
WScript.Sleep 100 'Wait for 100 MilliSeconds            
If colProcesses.Count = 0 Then 'If no more processes are running, exit loop                
blnRunning = False            
End If        
Loop      
End If
End Sub

KillProc "AWindowsProcess.exe"
Sub KillProc( myProcess )    
Dim blnRunning, colProcesses, objProcess    
blnRunning = False    
Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process", , 48 )    
For Each objProcess in colProcesses        
If LCase( myProcess ) = LCase( objProcess.Name ) Then             ' Confirm that the process was actually running             blnRunning = True             ' Get exact case for the actual process name            
myProcess  = objProcess.Name             ' Kill all instances of the process            
objProcess.Terminate()        
End If    
Next    
If blnRunning Then        
Do Until Not blnRunning            
Set colProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery( "Select * From Win32_Process Where Name = '"& myProcess & "'" )            
WScript.Sleep 100 'Wait for 100 MilliSeconds            
If colProcesses.Count = 0 Then 'If no more processes are running, exit loop                
blnRunning = False            
End If        
Loop      
End If
End Sub

msgbox "You have successfully disabled the monitoring system."
else


Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("""C:\Program Files (x86)\NetRef\NetRef Student\AWindowsProcess.exe""")
Set objShell = Nothing

Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("""C:\Program Files (x86)\NetRef\NetRef Student\WinNetRef.exe""")
Set objShell = Nothing

msgbox "You have successfully enabled the monitoring system."
end if
