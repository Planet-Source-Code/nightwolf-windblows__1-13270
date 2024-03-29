Attribute VB_Name = "Module1"
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public TopString
Public StartProgram
Public TempString
Public LineFeed
Public WindBlows_Version
Public Copyright_String

'Attribute VB_Name = "Shell_Execute"
Sub RunFile(ByVal File$, FilePath$, RunStyle)
    ' SW_HIDE = 0
    ' SW_SHOWNORMAL = 1
    ' SW_NORMAL = 1
    ' SW_SHOWMINIMIZED = 2
    ' SW_SHOWMAXIMIZED = 3
    ' SW_MAXIMIZE = 3
    ' SW_SHOWNOACTIVATE = 4
    ' SW_SHOW = 5
    ' SW_MINIMIZE = 6
    ' SW_SHOWMINNOACTIVE = 7
    ' SW_SHOWNA = 8
    ' SW_RESTORE = 9
    Const MB_ICONSTOP = 16
    Dim temp, Msg As String
    Dim X
    temp = GetActiveWindow()
    X = ShellExecute(temp, "Open", File$, "", FilePath$, RunStyle)


    If X < 32 Then


        Select Case X
            Case 0
            Msg = "The file could Not be run due To insufficient system memory or a corrupt program file"
            Case 2
            Msg = "File Not Found"
            Case 3
            Msg = "Invalid Path"
            Case 5
            Msg = "Sharing or protection error"
            Case 6
            Msg = "Separate data segments are required For Each task "
            Case 8
            Msg = "Insufficient memory To run the program"
            Case 10
            Msg = "Incorrect Windows version"
            Case 11
            Msg = "Invalid Program File"
            Case 12
            Msg = "Program file requires a different operating System "
            Case 13
            Msg = "Program requires MS-DOS 4.0"
            Case 14
            Msg = "Unknown program file type"
            Case 15
            Msg = "Windows prgram does Not support protected memory mode"
            Case 16
            Msg = "Invalid use of data segments when loading a second instance of a program"
            Case 19
            Msg = "Attempt To run a compressed program file"
            Case 20
            Msg = "Invalid dynamic link library"
            Case 21
            Msg = "Program requires Windows 32-bit extensions"
            Case 31
            Msg = "No application found For this file"
        End Select
End If
End Sub
