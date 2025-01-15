Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = Wscript.CreateObject("Wscript.Network")
On Error Resume Next

'Get the IP address of the current machine
strcomputer="."
Set objWMIService = GetObject("winmgmts:\\" & strcomputer & "\root\CIMV2")
Set IPItems = objWMIService.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each IPConfig In IPItems
	If Not IsNull(IPConfig.IPAddress) Then
		For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
			If varIP="" Then
				varIP=IPConfig.IPAddress(0)
			End If
		Next
	End If
Next

'Disable windows 10 default printer behaviour
objShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows\LegacyDefaultPrinterMode","00000001", "REG_DWORD"

'Split the IP address up into 4 separate parts and put it into an array
ArrayIP=Split(varIP,".")

'Create a variable containing the 3rd octet of the IP address
varThirdOctet=ArrayIP(2)

'Check value of varThirdOctet and run appropriate code
Select Case True
	Case varThirdOctet="1"
		'Christchurch Branch Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.1.249 -h 192.168.1.249 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Christchurch Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.1.249", 0, true
		objNetwork.SetDefaultPrinter "Christchurch Branch Copier"
	Case varThirdOctet="2"
		'Timaru Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.2.249 -h 192.168.2.249 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Timaru Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.2.249", 0, true
		'Timaru Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.2.250 -h 192.168.2.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Timaru Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.2.250", 0, true
		objNetwork.SetDefaultPrinter "Timaru Branch Copier"
	Case varThirdOctet="3"
		'Oamaru Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.3.250 -h 192.168.3.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Oamaru Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.3.250", 0, true
		'Oamaru Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.3.253 -h 192.168.3.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Oamaru Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.3.253", 0, true
		objNetwork.SetDefaultPrinter "Oamaru Branch Copier"
	Case varThirdOctet="4"
		'Dunedin Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.4.250 -h 192.168.4.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Dunedin Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.4.250", 0, true
		'Dunedin Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.4.253 -h 192.168.4.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Dunedin Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.4.253", 0, true
		objNetwork.SetDefaultPrinter "Dunedin Branch Copier"
	Case varThirdOctet="5"
		'Alexandra Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.5.250 -h 192.168.5.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Alexandra Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.5.250", 0, true
		'Alexandra Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.5.253 -h 192.168.5.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Alexandra Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.5.253", 0, true
		objNetwork.SetDefaultPrinter "Alexandra Branch Copier"
	Case varThirdOctet="6"
		'Queenstown Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.6.250 -h 192.168.6.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Queenstown Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.6.250", 0, true
		'Queenstown Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.6.252 -h 192.168.6.252 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Queenstown Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.6.252", 0, true
		objNetwork.SetDefaultPrinter "Queenstown Branch Copier"
	Case varThirdOctet="7"
		'Invercargill Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.7.253 -h 192.168.7.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Invercargill Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.7.253", 0, true
		objNetwork.SetDefaultPrinter "Invercargill Branch Copier"
	Case varThirdOctet="8"
		'Sockburn Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.8.250 -h 192.168.8.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Sockburn Branch Printer"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.8.250", 0, true
		'Sockburn Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.8.253 -h 192.168.8.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Sockburn Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.8.253", 0, true
		objNetwork.SetDefaultPrinter "Sockburn Branch Copier"
	Case varThirdOctet="9"
		'Cromwell Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.9.250 -h 192.168.9.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Cromwell Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.9.250", 0, true
		objNetwork.SetDefaultPrinter "Cromwell Branch Copier"
	Case varThirdOctet="10"
		'Rangiora Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.10.250 -h 192.168.10.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Rangiora Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.10.250", 0, true
		'Rangiora Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.10.253 -h 192.168.10.253 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Rangiora Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.10.253", 0, true
		objNetwork.SetDefaultPrinter "Rangiora Branch Copier"
	Case varThirdOctet="11"
		'Nelson Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.11.250 -h 192.168.11.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Nelson Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.11.250", 0, true
		objNetwork.SetDefaultPrinter "Nelson Branch Copier"
	Case varThirdOctet="12"
		'Downstairs Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.12.245 -h 192.168.12.245 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Head Office Downstairs Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.12.245", 0, true
		'Head Office Boardroom
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.12.218 -h 192.168.12.218 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Head Office Boardroom Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.12.218", 0, true
		objNetwork.SetDefaultPrinter "Head Office Downstairs Copier"
	Case varThirdOctet="13"
		'Rolleston Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.13.250 -h 192.168.13.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Rolleston Branch Copier"" -m ""Brother MFC-L9570CDW series"" -r ""IP_192.168.13.250", 0, true
		objNetwork.SetDefaultPrinter "Rolleston Branch Copier"
	Case varThirdOctet="14"
		'Ashburton Copier
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.14.250 -h 192.168.14.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Ashburton Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.14.250", 0, true
		'Ashburton Counter Printer
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.14.252 -h 192.168.14.252 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Ashburton Counter Printer"" -m ""Brother HL-L6400DW series"" -r ""IP_192.168.14.252", 0, true
		objNetwork.SetDefaultPrinter "Ashburton Branch Copier"
	Case varThirdOctet="15"
		'Wanaka Counter
		objShell.Run "cscript %WINDIR%\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r IP_192.168.15.250 -h 192.168.15.250 -me -y public -i 1 -o raw -n 9100", 0, true
		objShell.Run "cscript %WINDIR%\system32\printing_admin_scripts\en-us\prnmngr.vbs -a -p ""Wanaka Branch Copier"" -m ""SHARP BP-50C26 PCL6"" -r ""IP_192.168.15.250", 0, true
		objNetwork.SetDefaultPrinter "Wanaka Branch Copier"
End Select