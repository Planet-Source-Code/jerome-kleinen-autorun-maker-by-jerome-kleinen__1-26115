VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "autorun maker"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog savedialog 
      Left            =   3120
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".INF"
      DialogTitle     =   "save an INF-file"
      FileName        =   "Autorun.INF"
      Filter          =   "INF-files|*.INF"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      Caption         =   "icon"
      Height          =   2295
      Left            =   3840
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
      Begin VB.TextBox Text2 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "for example : menu.ico"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Caption         =   "program"
      Height          =   2295
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
      Begin VB.TextBox Text1 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "for example : menu.exe"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "how to use"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10695
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   $"autorun.frx":0000
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10335
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000080FF&
      Height          =   2295
      Left            =   7320
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Menu file 
      Caption         =   "file"
      Begin VB.Menu new 
         Caption         =   "new"
      End
      Begin VB.Menu save 
         Caption         =   "save"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "about"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim program As String
Dim icoon As String
Dim result As Integer

Private Sub about_Click()
result = MsgBox("autorun maker by Jerome Kleinen, produced in 30 mins", vbOKOnly, "about")
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub new_Click()
Text1.Text = ""
Text2.Text = ""
Label4.Caption = ""
Text1.SetFocus
End Sub

Private Sub save_Click()
If Text1.Text = "" Or Text2.Text = "" Then
result = MsgBox("fill out icon and/or program", vbOKOnly, "Error")
Else
program = Text1.Text
icoon = Text2.Text
Label4.Caption = "[autorun]" & vbNewLine & "open=" & program & vbNewLine & "icon=" & icoon
savedialog.ShowSave
Open savedialog.filename For Output As #1
Print #1, Label4.Caption
Close #1
End If
End Sub
