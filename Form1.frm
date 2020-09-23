VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hades iNet Connection Booster"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Welcome"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   $"Form1.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Windows 98 SE"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Windows 2000"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Windows NT"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Windows 98"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Windows 95"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Please click on your OS:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All this information was gathered off of various websites and
'text files. I would rather if you have an idea or info send it
'to me instead of making your own release. Contact: TripleXXX@bigfoot.com
'if you need anything of have more information for me.
' -Hades
'
'
'Also all the comments from now on are not written by me, they are
'taken out of the file I read the information on. This is because
'I am not good with these terms and could not explain them very
'well


Private Sub Label3_Click()
'Windows 95 section
'I didnt write this comments below, they are taken directly from the TXT i read about optimizing the sppeeds
'Optimizing the DefaultRcvWindow & DefaultTTL Settings (Windows 9x) Cool!
'The optimization of RcvWindow and DefaultTTL along with other registry settings such as MaxMTU and MaxMSS can speed up TCP/IP modem networking connections (eg. Internet connections).
'RWIN (Receive WINdow) is the buffer your machine waits to fill with data before attending to whatever other TCP transactions are occurring on the other threads and sockets WinSock has open while a connection is in progress.
'The value of TTL (Time To Live) defines how long a packet can stay active before being discarded. The default value is '32'.
If MsgBox("Hitting 'Yes' will make your internet settings optimized for Windows 95 systems. If you are not running Windows 95 or do not wish to update your settings hit 'No'", vbYesNo) = vbYes Then
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "64240", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
End If
MsgBox "Upon rebooting you *should* notice improvement!"
End Sub

Private Sub Label4_Click()
'Windows 98 section
If MsgBox("Hitting 'Yes' will make your internet settings optimized for Windows 95 systems. If you are not running Windows 98 or do not wish to update your settings hit 'No'", vbYesNo) = vbYes Then
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "372300", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    '(string var, recommended setting is 3. The possible settings are 0 - No Windowscaling and Timestamp Options, 1 - Window scaling but no Timestamp options, 3 - Window scaling and Time stamp options.)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "Tcp1323Opts", "3", REG_SZ
    '(string var, recommended setting is 1. Possible settings are 0 - No Sack options or 1 - Sack Option enabled)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "SackOpts", "1", REG_SZ
    '(DWORD decimal var, taking integer values from 2 to N. Recommended setting is 3)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "MaxDupAcks", "3", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
End If
MsgBox "Upon rebooting you *should* notice improvement!"
End Sub

Private Sub Label5_Click()
'Windows NT section
If MsgBox("Hitting 'Yes' will make your internet settings optimized for Windows 95 systems. If you are not running Windows NT 4.0 or do not wish to update your settings hit 'No'", vbYesNo) = vbYes Then
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "TcpWindowSize", "64240", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "DefaultTTL", "128", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "EnablePMTUDiscovery", "0", REG_SZ
    CreateNewKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", HKEY_CURRENT_USER
    SetKeyValue "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters", "EnablePMTUBHDetect", "0", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
End If
MsgBox "Upon rebooting you *should* notice improvement!"
End Sub

Private Sub Label6_Click()
'Windows 2000 section
MsgBox "Sorry I have not learned of any updates yet for Windows 2000. Please email them to: TripleXXX@bigfoot.com", vbOKOnly, "Help me update..."
End Sub

Private Sub Label7_Click()
'Windows 98 second edition section
If MsgBox("Hitting 'Yes' will make your internet settings optimized for Windows 95 systems. If you are not running Windows 98 or do not wish to update your settings hit 'No'", vbYesNo) = vbYes Then
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultRcvWindow", "372300", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "DefaultTTL", "128", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUBlackHoleDetect", "0", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\VxD\MSTCP", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VxD\MSTCP", "PMTUDiscovery", "0", REG_SZ
    '(string var, recommended setting is 3. The possible settings are 0 - No Windowscaling and Timestamp Options, 1 - Window scaling but no Timestamp options, 3 - Window scaling and Time stamp options.)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "Tcp1323Opts", "3", REG_SZ
    '(string var, recommended setting is 1. Possible settings are 0 - No Sack options or 1 - Sack Option enabled)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "SackOpts", "1", REG_SZ
    '(DWORD decimal var, taking integer values from 2 to N. Recommended setting is 3)
    CreateNewKey "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\VXD\MSTCP\Parameters", "MaxDupAcks", "3", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPer1_0Server", "dword:00000010", REG_SZ
    CreateNewKey "Software\Microsoft\Windows\CurrentVersion\Internet Settings", HKEY_CURRENT_USER
    SetKeyValue "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "MaxConnectionsPerServer", "dword:00000008", REG_SZ
    CreateNewKey "System\CurrentControlSet\Services\ICSharing\Settings\General", HKEY_LOCAL_MACHINE
    SetKeyValue "System\CurrentControlSet\Services\ICSharing\Settings\General", "internetMTU", "1500", REG_SZ
End If
MsgBox "Upon rebooting you *should* notice improvement!"
End Sub
