VERSION 5.00
Begin VB.Form loginfrm 
   BackColor       =   &H80000007&
   Caption         =   "Login"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4005
      Left            =   3945
      TabIndex        =   0
      Top             =   4080
      Width           =   12360
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         IMEMode         =   3  'DISABLE
         Left            =   7680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1920
         Width           =   4000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8280
         TabIndex        =   3
         Top             =   2955
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6240
         TabIndex        =   2
         Top             =   2955
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   7680
         TabIndex        =   1
         Top             =   480
         Width           =   4000
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   360
         Picture         =   "Login.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   5
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   0
      Picture         =   "Login.frx":91C0
      ScaleHeight     =   10875
      ScaleWidth      =   20670
      TabIndex        =   7
      Top             =   0
      Width           =   20730
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SUNSHINE NURSERY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   3578
         TabIndex        =   9
         Top             =   1320
         Width           =   13095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME TO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   6338
         TabIndex        =   8
         Top             =   600
         Width           =   7575
      End
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim unm As String
    Dim pwd As String
    unm = "admin"
    pwd = 12345
    If unm = Text1.Text And pwd = Text2.Text Then
        Text1.Text = ""
        Text2.Text = ""
        mainfrm.Show
        loginfrm.Hide
        
    Else
        MsgBox " Username or Password is incorrect ", vbExclamation, "Warning!"
    End If
End Sub
Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub
