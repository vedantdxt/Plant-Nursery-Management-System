VERSION 5.00
Begin VB.Form empfrm 
   BackColor       =   &H80000007&
   Caption         =   "Employee"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   2198
      TabIndex        =   0
      Top             =   2700
      Width           =   15855
      Begin VB.TextBox Text9 
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
         Left            =   11520
         TabIndex        =   18
         Top             =   360
         Width           =   3000
      End
      Begin VB.TextBox Text8 
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
         Left            =   11520
         TabIndex        =   17
         Top             =   4680
         Width           =   3000
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   16
         Top             =   4680
         Width           =   2000
      End
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
         Left            =   3840
         TabIndex        =   15
         Top             =   3840
         Width           =   10700
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
         Left            =   11520
         TabIndex        =   14
         Top             =   2952
         Width           =   3000
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
         TabIndex        =   13
         Top             =   2952
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
         TabIndex        =   12
         Top             =   2088
         Width           =   8000
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
         TabIndex        =   11
         Top             =   1224
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
         TabIndex        =   10
         Top             =   480
         Width           =   3000
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000007&
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
         Left            =   11640
         TabIndex        =   31
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label11 
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
         Left            =   11520
         TabIndex        =   30
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
         Caption         =   "Date of Joining :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8760
         TabIndex        =   9
         Top             =   360
         Width           =   2505
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
         Caption         =   "Blood Group :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   4680
         Width           =   2505
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Gender :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         Top             =   4680
         Width           =   2500
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Date of Birth :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   8760
         TabIndex        =   6
         Top             =   2955
         Width           =   2505
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   3816
         Width           =   2500
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Contact No. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   500
         Left            =   1080
         TabIndex        =   4
         Top             =   2952
         Width           =   2500
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Post :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   2088
         Width           =   2500
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   1224
         Width           =   2500
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Employee ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Frame2"
      Height          =   1215
      Left            =   5258
      TabIndex        =   19
      Top             =   8040
      Width           =   9735
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
         Left            =   5280
         TabIndex        =   27
         Top             =   480
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
         Left            =   7440
         TabIndex        =   26
         Top             =   480
         Width           =   1000
      End
      Begin VB.CommandButton cmddelete 
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
         Left            =   8520
         TabIndex        =   25
         Top             =   480
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
         Left            =   6360
         TabIndex        =   24
         Top             =   480
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
         TabIndex        =   23
         Top             =   480
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
         TabIndex        =   22
         Top             =   480
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
         TabIndex        =   21
         Top             =   480
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
         TabIndex        =   20
         Top             =   480
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10935
      Left            =   0
      Picture         =   "Employee.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   20670
      TabIndex        =   28
      Top             =   0
      Width           =   20730
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE DETAILS"
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
         Height          =   1455
         Left            =   3998
         TabIndex        =   29
         Top             =   720
         Width           =   12255
      End
   End
End
Attribute VB_Name = "empfrm"
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
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = Day(Now) & "-" & "0" & Month(Now) & "-" & Year(Now)
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
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
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
n = Text1.Text
s = "Select * from emp where eid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
Y = MsgBox("Are you sure to Add this record?", vbQuestion + vbYesNo, "ADD")
If (Y = vbNo) Then
Else
    rs1.AddNew
    rs1!eid = Text1.Text
    rs1!enm = Text2.Text
    rs1!epost = Text3.Text
    rs1!econt = Text4.Text
    rs1!edob = Text5.Text
    rs1!eadd = Text6.Text
    rs1!egen = Text7.Text
    rs1!ebgrp = Text8.Text
    rs1!edoj = Text9.Text
    rs1.Update
    MsgBox ("New Record Added!")
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Set con1 = New Connection
    Set rs1 = New Recordset
    con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
    rs1.Open "select * from emp", con1, adOpenStatic, adLockOptimistic
End If
Else
MsgBox ("Record Already Exits!")
End If
Else
MsgBox "Fill all Fields"
End If
cmdadd.Caption = "Add"
End If
End Sub

Private Sub cmddelete_Click()
Dim n As Integer
Dim s As String
Dim rsDel As Recordset
Set rsDel = New Recordset
n = Val(Text1.Text)
s = "delete from emp where eid=" & n
Y = MsgBox("Are you sure to Delete this record?", vbQuestion + vbYesNo, "DELETE")
If (Y = vbNo) Then
Else
rsDel.Open s, con1, adOpenStatic, adLockOptimistic
MsgBox ("Record Deleted")
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text5.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from emp", con1, adOpenStatic, adLockOptimistic
End If
End Sub
Private Sub cmdedit_Click()
If cmdedit.Caption = "Edit" Then
cb = TBenableTrue()
cmdedit.Caption = "Save"
Else
c = TBenableFalse()
cmdedit.Caption = "Edit"
rs1!econt = Text4.Text
rs1!eadd = Text6.Text
rs1.Update
MsgBox ("Changes Saved!")
End If
End Sub
Private Sub cmdlast_Click()
rs1.MoveLast
Text1.Text = rs1!eid
Text2.Text = rs1!enm
Text3.Text = rs1!epost
Text4.Text = rs1!econt
Text6.Text = rs1!eadd
Text5.Text = rs1!edob
Text7.Text = rs1!egen
Text8.Text = rs1!ebgrp
Text9.Text = rs1!edoj
End Sub

Private Sub cmdnext_Click()
rs1.MoveNext
If (rs1.EOF = True) Then
    rs1.MovePrevious
End If
Text1.Text = rs1!eid
Text2.Text = rs1!enm
Text3.Text = rs1!epost
Text4.Text = rs1!econt
Text6.Text = rs1!eadd
Text5.Text = rs1!edob
Text7.Text = rs1!egen
Text8.Text = rs1!ebgrp
Text9.Text = rs1!edoj
End Sub

Private Sub cmdprev_Click()
rs1.MovePrevious
If (rs1.BOF = True) Then
 rs1.MoveNext
End If
Text1.Text = rs1!eid
Text2.Text = rs1!enm
Text3.Text = rs1!epost
Text4.Text = rs1!econt
Text6.Text = rs1!eadd
Text5.Text = rs1!edob
Text7.Text = rs1!egen
Text8.Text = rs1!ebgrp
Text9.Text = rs1!edoj
End Sub

Private Sub cmdsearch_Click()
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
n = Val(InputBox("Enter Employee ID"))
s = "Select * from emp where eid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
MsgBox "Record Not Found."
End If
While Not rsSearch.EOF
Text1.Text = rsSearch!eid
Text2.Text = rsSearch!enm
Text3.Text = rsSearch!epost
Text4.Text = rsSearch!econt
Text6.Text = rsSearch!eadd
Text5.Text = rsSearch!edob
Text7.Text = rsSearch!egen
Text8.Text = rsSearch!ebgrp
Text9.Text = rsSearch!edoj
'and soon
rsSearch.MoveNext
Wend
End Sub
Private Sub Form_Load()
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from emp", con1, adOpenStatic, adLockOptimistic
rs1.MoveFirst
Text1.Text = rs1!eid
Text2.Text = rs1!enm
Text3.Text = rs1!epost
Text4.Text = rs1!econt
Text6.Text = rs1!eadd
Text5.Text = rs1!edob
Text7.Text = rs1!egen
Text8.Text = rs1!ebgrp
Text9.Text = rs1!edoj
End Sub
Private Sub cmdfirst_Click()
rs1.MoveFirst
Text1.Text = rs1!eid
Text2.Text = rs1!enm
Text3.Text = rs1!epost
Text4.Text = rs1!econt
Text6.Text = rs1!eadd
Text5.Text = rs1!edob
Text7.Text = rs1!egen
Text8.Text = rs1!ebgrp
Text9.Text = rs1!edoj
End Sub

Function TBenableTrue()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text8.Enabled = False
Text6.Enabled = True
Text7.Enabled = False
Text9.Enabled = False
Text4.Enabled = True
Text5.Enabled = False
End Function

Function TBenableFalse()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text8.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text9.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Function
