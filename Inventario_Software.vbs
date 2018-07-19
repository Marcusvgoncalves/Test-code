On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Shell = WScript.CreateObject("WScript.Shell")
Set ShellApplication = WScript.CreateObject("Shell.Application")
Set objNetwork = CreateObject("WScript.Network")

Dim strCaminho 
Dim STRcomputer
Dim STRTipoServer

'Pega nome da maquina que esta sendo executado o script
STRcomputer = objNetwork.ComputerName

'Coloque aqui a maquina que será verificada
'STRcomputer =Inputbox("Informe o nome maquina","aperte ok!")

WScript.Echo "Nome da maquina " & STRcomputer


'Nome do arquivo com a lista de software instalados na maquina "Inventario_de_software.txt"
strLogFile = "Inventario_de_software.txt"


Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)
objLogFile.WriteLine "*****************************************************************"
objLogFile.WriteLine "Lista de software da maquina: " &  STRcomputer & "Data  " & now
objLogFile.WriteLine "*****************************************************************"
  

  'Coleta Software
  Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
  strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
  strEntry1a = "DisplayName"
  strEntry1b = "QuietDisplayName"
  strEntry3 = "VersionMajor"
  strEntry4 = "VersionMinor"


  
  Set objReg = GetObject("winmgmts://" & STRcomputer & _
  "/root/default:StdRegProv")
  objReg.EnumKey HKLM, strKey, arrSubkeys
  For Each strSubkey In arrSubkeys
   wlinha=""
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
   objReg.GetStringValue HKLM, strKey & strSubkey, _
   strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
   wlinha = wlinha + strValue1 
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
  strEntry3, intValue3
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry4, intValue4
  If intValue3 <> "" Then
   wlinha = wlinha + intValue3 + "." + intValue4 
  End If
  If Len(wlinha) > 8 Then
   objLogFile.WriteLine wlinha 
  End If
  Next
  
WScript.Echo "Script executado com sucesso!"
wscript.quit