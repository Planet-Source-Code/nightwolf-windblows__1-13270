VERSION 5.00
Object = "{B6C1EA38-375B-11D4-93AB-E7C32384627A}#3.0#0"; "FREELIB.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "WindBlows"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "R&efresh"
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reboot"
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Shutdown"
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400000&
      Caption         =   "Start Menu"
      ForeColor       =   &H00FF8080&
      Height          =   3150
      Left            =   0
      TabIndex        =   8
      Top             =   3885
      Width           =   5895
      Begin VB.Frame Frame4 
         BackColor       =   &H00400000&
         ForeColor       =   &H00FF8080&
         Height          =   2735
         Left            =   2225
         TabIndex        =   11
         Top             =   240
         Width           =   60
      End
      Begin VB.FileListBox File2 
         Height          =   2820
         Left            =   2320
         TabIndex        =   10
         Top             =   240
         Width           =   3495
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Caption         =   "QuickDesk"
      ForeColor       =   &H00FF8080&
      Height          =   3555
      Left            =   5930
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
      Begin VB.ListBox List1 
         Height          =   3180
         ItemData        =   "Form1.frx":08CA
         Left            =   120
         List            =   "Form1.frx":0901
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Desktop"
      ForeColor       =   &H00FF8080&
      Height          =   3375
      Left            =   5930
      TabIndex        =   3
      Top             =   0
      Width           =   2535
      Begin VB.FileListBox File1 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.OLE OLE1 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin FreeLibSrc.FreeLib FreeLib1 
      Height          =   480
      Left            =   6000
      TabIndex        =   2
      Top             =   1800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      IniFile         =   ""
      IniSection      =   ""
      IniSize         =   ""
      IniDefault      =   ""
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00400000&
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   5895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FreeLib1.Shutdown = True
MsgBox "Shutdown canceled."
End Sub

Private Sub Command2_Click()
FreeLib1.RebootSystem = True
MsgBox "Reboot canceled."
End Sub

Private Sub Command3_Click()
On Error Resume Next
File1.Refresh
File2.Refresh
Dir1.Refresh
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Dir1.Refresh
File2.Path = Dir1.Path
If Not Left$(Dir1.Path, Len(StartProgram)) = StartProgram Then Dir1.Path = StartProgram: File2.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
On Error Resume Next
RunFile File1, File1.Path + "\", 5
FreeLib1.ConnectTo Left$(File1, Len(File1) - 4)
End Sub

Private Sub File2_DblClick()
On Error Resume Next
RunFile File2, File2.Path + "\", 5
End Sub

Private Sub Form_Activate()
On Error Resume Next
File1.Refresh
Dir1.Refresh
File2.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
StartProgram = FreeLib1.RegistryRead(USERS, ".Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs")
FreeLib1.TaskBarVisible = True
File1.Path = FreeLib1.RegistryRead(USERS, ".Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
Dir1.Path = StartProgram
File2.Path = StartProgram
'MsgBox FreeLib1.RegistryRead(USERS, ".Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", Desktop)
File1.Refresh
LineFeed = CStr(Chr$(13) + Chr$(10))
WindBlows_Version = "v0.1a"
Copyright_String = "(C) NightWolf Creations 1997-2000"
TopString = "WindBlows " & WindBlows_Version & LineFeed & Copyright_String & LineFeed & LineFeed & LineFeed
'-----------------------------------------
'|WindBlows v0.1a (C) NightWolf Creations|
'|All Rights Reserved 1997-2000 © N  w  C|
'|Source given away please do not modify.|
'-----------------------------------------

Text1.Text = Text1.Text + "WindBlows " & WindBlows_Version & LineFeed & Copyright_String & LineFeed & LineFeed
Text2.SetFocus
End Sub

Private Sub List1_DblClick()
On Error Resume Next
If List1 = "NotePad" Then Shell FreeLib1.PathWindows + "\" + "notepad.exe", vbNormalFocus
If List1 = "WordPad" Then Shell FreeLib1.PathWindows + "\" + "write.exe", vbNormalFocus
If List1 = "Paint Brush" Then Shell FreeLib1.PathWindows + "\" + "pbrush.exe", vbNormalFocus
If List1 = "Registry Editor" Then Shell FreeLib1.PathWindows + "\" + "regedit.exe", vbNormalFocus
If List1 = "Calculator" Then Shell FreeLib1.PathWindows + "\" + "calc.exe", vbNormalFocus
If List1 = "CD Player" Then Shell FreeLib1.PathWindows + "\" + "cdplayer.exe", vbNormalFocus
If List1 = "Character Map" Then Shell FreeLib1.PathWindows + "\" + "charmap.exe", vbNormalFocus
If List1 = "Defrag" Then Shell FreeLib1.PathWindows + "\" + "defrag.exe", vbNormalFocus
If List1 = "Explorer" Then Shell FreeLib1.PathWindows + "\" + "explorer.exe", vbNormalFocus
If List1 = "Media Player" Then Shell FreeLib1.PathWindows + "\" + "mplayer.exe", vbNormalFocus
If List1 = "Scandisk" Then Shell FreeLib1.PathWindows + "\" + "scandskw.exe", vbNormalFocus
If List1 = "Sound Recorder" Then Shell FreeLib1.PathWindows + "\" + "sndrec32.exe", vbNormalFocus
If List1 = "Sound Control" Then Shell FreeLib1.PathWindows + "\" + "sndvol32.exe", vbNormalFocus
If List1 = "System Monitor" Then Shell FreeLib1.PathWindows + "\" + "sysmon.exe", vbNormalFocus
If List1 = "Telnet" Then Shell FreeLib1.PathWindows + "\" + "telnet.exe", vbNormalFocus
If List1 = "Windows File Manager" Then Shell FreeLib1.PathWindows + "\" + "winfile", vbNormalFocus
If List1 = "Windows IP Configuration" Then Shell FreeLib1.PathWindows + "\" + "winipcfg.exe", vbNormalFocus
End Sub

Private Sub Text1_Change()
On Error Resume Next
Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 1
Text1.Refresh
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
 If UCase$(Text2.Text) = "EXIT" Then End
 If UCase$(Text2.Text) = "VER" Then Text1.Text = TopString + "WindBlows [" + WindBlows_Version + "]"
 If UCase$(Text2.Text) = "ABOUT" Then Text1.Text = TopString + "WindBlows was created by NightWolf, member and founder of NightWolf Creations." + LineFeed + LineFeed + "WindBlows is (C) to NightWolf Creations 1997-2000" + LineFeed + LineFeed + "It's source is given away, but if you wan't to modify it or use it in part for your software, you must include credits to NightWolf Creations" + LineFeed + LineFeed + "Big thanks to FreeLib's creator."
 If UCase$(Left$(Text2.Text, 4)) = "EXEC" Then Shell Right$(Text2.Text, Len(Text2.Text) - 5), vbNormalFocus: Text1.Text = TopString + "Executing " + Right$(Text2.Text, Len(Text2.Text) - 5) + "..."
 If UCase$(Text2.Text) = "INSTALL" Then Text1.Text = TopString + "To install WindBlows, you must follows these steps:" + LineFeed + LineFeed + "1) Shut down your computer to MS-DOS mode" + LineFeed + "2) edit system.ini" + LineFeed + "3) add 'shell=windblows.exe'" + LineFeed + "4) type 'win' to get back to Windows"
 If UCase$(Text2.Text) = "HELP" Then Text1.Text = TopString + "Here are the following WindBlows commands:" + LineFeed + LineFeed + "exec   - executes an app." + LineFeed + "         example: exec c:\autoexec.bat" + LineFeed + "about  - Displays 'about' message" + LineFeed + "exit   - quits WindBlows" + LineFeed + "help   - this screen" + LineFeed + "install- instructions on installation" + LineFeed + "ver    - displays software version"
  Text2.Text = ""
End If
End Sub
