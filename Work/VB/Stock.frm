VERSION 5.00
Begin VB.Form stockfrm 
   BackColor       =   &H80000012&
   Caption         =   "Stock"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10515
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   4898
      TabIndex        =   0
      Top             =   2910
      Width           =   10455
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
         Height          =   510
         Left            =   7560
         TabIndex        =   6
         Top             =   3480
         Width           =   2000
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
         Height          =   510
         Left            =   2880
         TabIndex        =   5
         Top             =   3480
         Width           =   2000
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
         Height          =   510
         Left            =   3840
         TabIndex        =   4
         Top             =   2520
         Width           =   5040
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
         Height          =   510
         Left            =   3840
         TabIndex        =   3
         Top             =   1560
         Width           =   5000
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
         Left            =   7440
         TabIndex        =   2
         Top             =   600
         Width           =   2000
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
         Left            =   2880
         TabIndex        =   1
         Top             =   600
         Width           =   2000
      End
      Begin VB.Label Label8 
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
         Left            =   7200
         TabIndex        =   23
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Stock Date :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   3480
         Width           =   2205
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Stock Price :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   795
         TabIndex        =   11
         Top             =   3480
         Width           =   2000
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Stock Quantity :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   795
         TabIndex        =   10
         Top             =   2520
         Width           =   3000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Stock Type :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   795
         TabIndex        =   9
         Top             =   1560
         Width           =   3000
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Plant ID :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   5280
         TabIndex        =   8
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Stock ID :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   795
         TabIndex        =   7
         Top             =   600
         Width           =   2000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   5895
      TabIndex        =   13
      Top             =   7440
      Width           =   8340
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
         Left            =   6960
         TabIndex        =   20
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
         Left            =   4800
         TabIndex        =   19
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
         Left            =   5880
         TabIndex        =   18
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
         Left            =   2640
         TabIndex        =   17
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
         Left            =   3720
         TabIndex        =   16
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
         Left            =   1560
         TabIndex        =   15
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
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10575
      Left            =   -240
      Picture         =   "Stock.frx":0000
      ScaleHeight     =   10515
      ScaleWidth      =   20670
      TabIndex        =   21
      Top             =   0
      Width           =   20730
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK DETAILS"
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
         Height          =   1575
         Left            =   5258
         TabIndex        =   22
         Top             =   600
         Width           =   9735
      End
   End
End
Attribute VB_Name = "stockfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As Connection
Dim rs1 As Recordset

Dim con12 As Connection
Dim rs12 As Recordset

Private Sub cmdminus_Click()
Set con12 = New Connection
Set rs12 = New Recordset
Y = Val(InputBox("Enter Stock ID"))
If Y <> 0 Then
s = "select * from stock where sid=" & Y
rs12.Open s, con1, adOpenStatic, adLockOptimistic
quan = Val(InputBox("Enter Quantity"))
If quan <> 0 Then
If quan > rs12!squantity Then
MsgBox ("Required Quantity Exceeded!")
Else
ch = MsgBox("Really Wanna Made Purchase?", vbQuestion + vbYesNo, "PURCHASE")
If (ch = vbYes) Then
    rs12!squantity = rs12!squantity - quan
    rs12.Update
Else
End If
End If
End If
End If
End Sub

Private Sub Form_Load()
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from stock", con1, adOpenStatic, adLockOptimistic
rs1.MoveFirst
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End Sub
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
n = Text1.Text
s = "Select * from stock where sid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
Y = MsgBox("Are you sure to Add this record?", vbQuestion + vbYesNo, "ADD")
If (Y = vbNo) Then
Else
    rs1.AddNew
    rs1!sid = Text1.Text
    rs1!pid = Text2.Text
    rs1!stype = Text3.Text
    rs1!squantity = Text4.Text
    rs1!sprice = Text5.Text
    rs1!stockdate = Text6.Text
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
    rs1.Open "select * from stock", con1, adOpenStatic, adLockOptimistic
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
s = "delete from stock where sid=" & n
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
rs1.Open "select * from stock", con1, adOpenStatic, adLockOptimistic
rs1.MoveFirst
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End If
End Sub
Private Sub cmdlast_Click()
rs1.MoveLast
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End Sub

Private Sub cmdnext_Click()
rs1.MoveNext
If (rs1.EOF = True) Then
    rs1.MovePrevious
End If
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End Sub

Private Sub cmdprev_Click()
rs1.MovePrevious
If (rs1.BOF = True) Then
 rs1.MoveNext
End If
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End Sub

Private Sub cmdsearch_Click()
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
n = Val(InputBox("Enter Employee ID"))
s = "Select * from stock where sid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
MsgBox "Record Not Found."
End If
While Not rsSearch.EOF
Text1.Text = rsSearch!sid
Text2.Text = rsSearch!pid
Text3.Text = rsSearch!stype
Text4.Text = rsSearch!squantity
Text5.Text = rsSearch!sprice
Text6.Text = rsSearch!stockdate
'and soon
rsSearch.MoveNext
Wend
End Sub

Private Sub cmdfirst_Click()
rs1.MoveFirst
Text1.Text = rs1!sid
Text2.Text = rs1!pid
Text3.Text = rs1!stype
Text4.Text = rs1!squantity
Text5.Text = rs1!sprice
Text6.Text = rs1!stockdate
End Sub
