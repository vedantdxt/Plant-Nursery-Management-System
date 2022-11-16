VERSION 5.00
Begin VB.Form custfrm 
   BackColor       =   &H80000017&
   Caption         =   "Customer"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   DrawMode        =   1  'Blackness
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   2138
      TabIndex        =   0
      Top             =   3540
      Width           =   15975
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   11880
         TabIndex        =   12
         Top             =   480
         Width           =   3000
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   10920
         TabIndex        =   11
         Top             =   2928
         Width           =   4000
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   3840
         TabIndex        =   10
         Top             =   2928
         Width           =   4000
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   3840
         TabIndex        =   9
         Top             =   2112
         Width           =   11000
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   3840
         TabIndex        =   8
         Top             =   1296
         Width           =   8000
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label lbldateformat 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Date Format : (dd-mm-yyyy)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11760
         TabIndex        =   13
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Registration Date :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   8760
         TabIndex        =   6
         Top             =   480
         Width           =   3000
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "E-mail ID :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   8760
         TabIndex        =   5
         Top             =   2928
         Width           =   3000
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Contact No. :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   2928
         Width           =   3000
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   2112
         Width           =   3000
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   1296
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Customer ID :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   3000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   5318
      TabIndex        =   14
      Top             =   7200
      Width           =   9615
      Begin VB.CommandButton cmddel 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8400
         TabIndex        =   22
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   7320
         TabIndex        =   21
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6240
         TabIndex        =   20
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   5160
         TabIndex        =   19
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next ->"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "<- Prev"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2400
         TabIndex        =   17
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdlast 
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   0
      Picture         =   "Customer.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   20670
      TabIndex        =   23
      Top             =   0
      Width           =   20730
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER DETAILS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   4238
         TabIndex        =   24
         Top             =   1080
         Width           =   11775
      End
   End
End
Attribute VB_Name = "custfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As Connection
Dim rs1 As Recordset

Private Sub cmdadd_Click()
If cmdadd.Caption = "Add" Then
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = Day(Now) & "-" & "0" & Month(Now) & "-" & Year(Now)
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
cmdadd.Caption = "Save"
Else
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
If (Text1.Text <> "") And (Text2.Text <> "") And (Text3.Text <> "") And (Text4.Text <> "") And (Text5.Text <> "") And (Text6.Text <> "") Then
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
n = Text1.Text
s = "Select * from cust where cid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
Y = MsgBox("Are you sure to Add this record?", vbQuestion + vbYesNo, "ADD")
If (Y = vbNo) Then
Else
    rs1.AddNew
    rs1!cid = Text1.Text
    rs1!cname = Text2.Text
    rs1!cadd = Text3.Text
    rs1!ccont = Text4.Text
    rs1!cmail = Text5.Text
    rs1!regdate = Text6.Text
    rs1.Update
    MsgBox ("New Record Added!")
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Set con1 = New Connection
    Set rs1 = New Recordset
    con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
    rs1.Open "select * from cust", con1, adOpenStatic, adLockOptimistic
End If
Else
MsgBox ("Record Already Exits!")
End If
Else
MsgBox "Fill all Fields"
End If
cmdedit.Caption = "Add"
End If
End Sub

Private Sub cmddel_Click()
Dim n As Integer
Dim s As String
Dim rsDel As Recordset
Set rsDel = New Recordset
n = Val(Text1.Text)
s = "delete from cust where cid=" & n
Y = MsgBox("Are you sure to Delete this record?", vbQuestion + vbYesNo, "DELETE")
If (Y = vbNo) Then
Else
rsDel.Open s, con1, adOpenStatic, adLockOptimistic
MsgBox ("Record Deleted")
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from cust", con1, adOpenStatic, adLockOptimistic
End If
End Sub

Private Sub cmdedit_Click()
If cmdedit.Caption = "Edit" Then
cb = TBenableTrue()
cmdedit.Caption = "Save"
Else
c = TBenableFalse()
cmdedit.Caption = "Edit"
rs1!cadd = Text3.Text
rs1!ccont = Text4.Text
rs1!cmail = Text5.Text
rs1.Update
MsgBox ("Changes Saved!")
End If
End Sub

Private Sub cmdfirst_Click()
rs1.MoveFirst
Text1.Text = rs1!cid
Text2.Text = rs1!cname
Text3.Text = rs1!cadd
Text4.Text = rs1!ccont
Text5.Text = rs1!cmail
Text6.Text = rs1!regdate
End Sub


Private Sub cmdlast_Click()
rs1.MoveLast
Text1.Text = rs1!cid
Text2.Text = rs1!cname
Text3.Text = rs1!cadd
Text4.Text = rs1!ccont
Text5.Text = rs1!cmail
Text6.Text = rs1!regdate
End Sub

Private Sub cmdnext_Click()
rs1.MoveNext
If (rs1.EOF = True) Then
    rs1.MovePrevious
End If
Text1.Text = rs1!cid
Text2.Text = rs1!cname
Text3.Text = rs1!cadd
Text4.Text = rs1!ccont
Text5.Text = rs1!cmail
Text6.Text = rs1!regdate
End Sub

Private Sub cmdprev_Click()
rs1.MovePrevious
If (rs1.BOF = True) Then
 rs1.MoveNext
End If
Text1.Text = rs1!cid
Text2.Text = rs1!cname
Text3.Text = rs1!cadd
Text4.Text = rs1!ccont
Text5.Text = rs1!cmail
Text6.Text = rs1!regdate
End Sub

Private Sub cmdsearch_Click()
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
n = Val(InputBox("Enter Customer ID"))
s = "Select * from cust where cid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
MsgBox "Record Not Found."
End If
While Not rsSearch.EOF
Text1.Text = rsSearch!cid
Text2.Text = rsSearch!cname
Text3.Text = rsSearch!cadd
Text4.Text = rsSearch!ccont
Text5.Text = rsSearch!cmail
Text6.Text = rsSearch!regdate
'and soon
rsSearch.MoveNext
Wend
End Sub
Private Sub Form_Load()
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from cust", con1, adOpenStatic, adLockOptimistic
rs1.MoveFirst
Text1.Text = rs1!cid
Text2.Text = rs1!cname
Text3.Text = rs1!cadd
Text4.Text = rs1!ccont
Text5.Text = rs1!cmail
Text6.Text = rs1!regdate
End Sub
Function TBenableTrue()
Text1.Enabled = False
Text2.Enabled = False
Text6.Enabled = False
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Function
Function TBenableFalse()
Text1.Enabled = True
Text2.Enabled = True
Text6.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Function
