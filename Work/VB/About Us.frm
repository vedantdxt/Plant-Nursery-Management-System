VERSION 5.00
Begin VB.Form aboutfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "About Us"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "About Us.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   360
      TabIndex        =   0
      Top             =   900
      Width           =   15735
      Begin VB.Label Label7 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   600
         TabIndex        =   7
         Top             =   6480
         Width           =   10005
      End
      Begin VB.Label Label6 
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   6
         Top             =   3480
         Width           =   1845
      End
      Begin VB.Label Label5 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   600
         TabIndex        =   5
         Top             =   4160
         Width           =   10005
      End
      Begin VB.Label Label4 
         Caption         =   "Contact Us :"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   4
         Top             =   5800
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   600
         TabIndex        =   3
         Top             =   2080
         Width           =   10005
      End
      Begin VB.Label Label2 
         Caption         =   "Founders :"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   1400
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "SUNSHINE NURSERY"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   4965
      End
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label3.Caption = "Vedant Dixit" & vbNewLine & "Shivam Gaikwad"
Label5.Caption = "Plot No. 7, Sunshine Nursery," & vbNewLine & "Sunny Roadways, Destiny Lane," & vbNewLine & "Nashik - 422024"
Label7.Caption = "Landline - (0253)-426426" & vbNewLine & "Mobile No. - (+91)-9773693690" & vbNewLine & "Email - nursery@sunshine.com" & vbNewLine & "Website - www.sunshinenursery.com"
End Sub
