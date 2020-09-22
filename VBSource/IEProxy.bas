Attribute VB_Name = "IEProxy"

'############################################################################
'Proxy Checker
'
'Dear friends!
'This is a free Proxy Checker . If you like my program and want me to
'continue to improve it and it's capabilities,your donations would be welcome.
'For more information please contact me !
'
'reza kargar
'
'web:    www.ragrak.com
'
'email:  kargar.reza@ gmail.com
'
'phone : +98-9122767401
'#############################################################################



Public Function SaveProxySettings(IPStr As String, PortStr As String, Enable As Integer)
 
 On Error GoTo errors
 
 Dim Create
 Dim Key
 Dim Address As String
 
 Address = IPStr & ":" & PortStr
 
 Const ProxyServer = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
 Const ProxyEnable = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
 
 Set Create = CreateObject("wscript.shell")

   Key = ProxyEnable
   Create.RegWrite Key, Enable, "REG_DWORD"

   Key = ProxyServer
   Create.RegWrite Key, Address, "REG_SZ"
   Exit Function
   
errors:
   MsgBox "Error!" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error Found"
   
End Function


