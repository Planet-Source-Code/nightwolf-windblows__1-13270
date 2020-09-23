VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Height          =   3615
      Left            =   85
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   3000
         Top             =   3120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A fast alternative to the common explorer.exe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   3120
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   3285
         Left            =   120
         Picture         =   "Splash.frx":0000
         Top             =   240
         Width           =   7080
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Splash
Form1.Show
End Sub
