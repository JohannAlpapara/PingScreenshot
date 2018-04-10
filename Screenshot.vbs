Option Explicit  
 ' ------- Declare the variables  -----------------  
 Dim WshShell, counter, addresses, ips, x, item  
    
	
	
 ' ------- Only edit this part ------------- 
 ' ------- Supply IP addresses in the array  -----------------  
 ips = Array ("127.0.0.1","192.168.0.1","etc")
 ' ------- Only edit this part ------------- 
 
 
 
 For Each x in ips
			counter = counter + 1    
 Next	
 
 ' ------- Blocks of code for the test steps  -------------  
 Sub OpenWord  
                 Set WshShell = WScript.CreateObject("WScript.Shell")  
                 WshShell.Run "winword"  
                 WScript.Sleep 5000  
 End Sub  
 
 Sub OpenCMD  

				For each item in ips
 
				 addresses=addresses&item

				 Set WshShell = WScript.CreateObject("WScript.Shell")  
                 WshShell.Run "cmd"  
				 WScript.Sleep 1000  
				 WshShell.SendKeys "ping" + " " + (addresses)
				 WScript.Sleep 1000  
				 WshShell.SendKeys "{ENTER}"
                 WScript.Sleep 5000  
				 
				 addresses = ""
				 
				 Call ActivateCMD 
				 Call TakeScreenShot  
				 Call ActivateWordAndSaveTheImage  
				 Call CloseCMD
				Next
				
				msgbox "Finished"
				 
 End Sub  
    
 Sub ActivateCMD  
                 WshShell.AppActivate "Command Prompt"
                 WScript.Sleep 1000  
 End Sub  
    
 Sub TakeScreenShot  
                 Set Wshshell = CreateObject("Word.Basic")  
                 WshShell.SendKeys "(%{1068})" 'Screenshots the currently active window, not the whole screen  
                 WScript.Sleep 1000  
 End Sub  
    
 Sub ActivateWordAndSaveTheImage  
                 WshShell.AppActivate "Document1 - Microsoft Word"  
                 WScript.Sleep 1000  
    
                 WshShell.sendkeys "^(v)"  
                 WScript.Sleep 1000  
    
                 WshShell.sendkeys "{ENTER}"  
                 WScript.Sleep 1000  
 End Sub  
    
 Sub CloseCMD  
                 WScript.Sleep 1000  
                 WshShell.AppClose "C:\Windows\System32\cmd.exe"  
                 WScript.Sleep 1000  
 End Sub  
    
 ' ------- Call the Blocks of code  ----------------  
 Call OpenWord  
 Call OpenCMD  