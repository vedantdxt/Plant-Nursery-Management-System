VERSION 5.00
Begin VB.Form plantfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Frame2"
      Height          =   5535
      Left            =   1058
      TabIndex        =   9
      Top             =   2700
      Width           =   18135
      Begin VB.PictureBox imgPLANT 
         BackColor       =   &H80000007&
         Height          =   4500
         Left            =   1080
         ScaleHeight     =   4440
         ScaleMode       =   0  'User
         ScaleWidth      =   4440
         TabIndex        =   22
         Top             =   480
         Width           =   4500
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
         Left            =   11000
         TabIndex        =   15
         Top             =   4800
         Width           =   4000
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
         Left            =   11000
         TabIndex        =   14
         Top             =   3936
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
         Height          =   500
         Left            =   11000
         TabIndex        =   13
         Top             =   3072
         Width           =   3000
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
         Left            =   11000
         TabIndex        =   12
         Top             =   2160
         Width           =   6000
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
         Left            =   11000
         TabIndex        =   11
         Top             =   1344
         Width           =   6000
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
         Left            =   11000
         TabIndex        =   10
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Plant Price :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   21
         Top             =   4800
         Width           =   4935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Plant Age :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   20
         Top             =   3936
         Width           =   4935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Plant Type :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   19
         Top             =   3072
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Plant Scientific Name :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   18
         Top             =   2208
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Plant Common Name :"
         BeginProperty Font 
            Name            =   "Segoe UI Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   17
         Top             =   1344
         Width           =   4935
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   6720
         TabIndex        =   16
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   5018
      TabIndex        =   0
      Top             =   8040
      Width           =   10215
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
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   600
         Width           =   1095
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
         Height          =   495
         Left            =   7560
         TabIndex        =   7
         Top             =   600
         Width           =   1095
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
         Left            =   5400
         TabIndex        =   6
         Top             =   600
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
         Left            =   3600
         TabIndex        =   5
         Top             =   600
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
         Left            =   2520
         TabIndex        =   4
         Top             =   600
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
         Left            =   6480
         TabIndex        =   3
         Top             =   600
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
         Left            =   1440
         TabIndex        =   2
         Top             =   600
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
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   10935
      Left            =   0
      Picture         =   "Plants.frx":0000
      ScaleHeight     =   10875
      ScaleWidth      =   20670
      TabIndex        =   23
      Top             =   0
      Width           =   20730
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PLANT DETAILS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1335
         Left            =   4358
         TabIndex        =   24
         Top             =   720
         Width           =   11535
      End
   End
End
Attribute VB_Name = "plantfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As Connection
Dim rs1 As Recordset

Private Sub cmddel_Click()
Dim n As Integer
Dim s As String
Dim rsDel As Recordset
Set rsDel = New Recordset
n = Val(Text1.Text)
s = "delete from plant where pid=" & n
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
rs1.Open "select * from plant", con1, adOpenStatic, adLockOptimistic
End If
End Sub

Private Sub cmdedit_Click()
If cmdedit.Caption = "Edit" Then
cb = TBenableTrue()
cmdedit.Caption = "Save"
Else
c = TBenableFalse()
cmdedit.Caption = "Edit"
rs1!pprice = Text6.Text
rs1!Page = Text5.Text
rs1.Update
MsgBox ("Changes Saved!")
End If
End Sub

Private Sub cmdadd_Click()
If cmdadd.Caption = "Add" Then
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
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
s = "Select * from plant where pid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
Y = MsgBox("Are you sure to Add this record?", vbQuestion + vbYesNo, "ADD")
If (Y = vbNo) Then
Else
    rs1.AddNew
    rs1!pid = Text1.Text
    rs1!pcnm = Text2.Text
    rs1!psnm = Text3.Text
    rs1!ptype = Text4.Text
    rs1!Page = Text5.Text
    rs1!pprice = Text6.Text
    rs1!pphoto = InputBox("Enter Plant Picture Absolute Address")
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
    rs1.Open "select * from plant", con1, adOpenStatic, adLockOptimistic
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

Private Sub Form_Load()
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from plant", con1, adOpenStatic, adLockOptimistic
rs1.MoveFirst
ds = data_scroll()
End Sub
Private Sub cmdfirst_Click()
rs1.MoveFirst
ds = data_scroll()
End Sub
Private Sub cmdlast_Click()
rs1.MoveLast
ds = data_scroll()
End Sub
Private Sub cmdnext_Click()
rs1.MoveNext
If (rs1.EOF = True) Then
    rs1.MovePrevious
End If
ds = data_scroll()
End Sub
Private Sub cmdprev_Click()
rs1.MovePrevious
If (rs1.BOF = True) Then
    rs1.MoveNext
End If
ds = data_scroll()
End Sub
Private Sub cmdsearch_Click()
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
n = Val(InputBox("Enter Plant ID"))
s = "Select * from plant where pid=" & n
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
MsgBox "Record Not Found."
End If
While Not rsSearch.EOF
imgsrc = rsSearch!pphoto
imgPLANT.Picture = LoadPicture(imgsrc)
        imgPLANT.ScaleMode = 3
        imgPLANT.AutoRedraw = True
        imgPLANT.PaintPicture imgPLANT.Picture, _
        0, 0, imgPLANT.ScaleWidth, imgPLANT.ScaleHeight, _
        0, 0, imgPLANT.Picture.Width / 26.46, _
        imgPLANT.Picture.Height / 26.46
        imgPLANT.Picture = imgPLANT.Image
Text1.Text = rsSearch!pid
Text2.Text = rsSearch!pcnm
Text3.Text = rsSearch!psnm
Text4.Text = rsSearch!ptype
Text5.Text = rsSearch!Page
Text6.Text = rsSearch!pprice
'and soon
rsSearch.MoveNext
Wend
End Sub

Function data_scroll()
imgsrc = rs1!pphoto
imgPLANT.Picture = LoadPicture(imgsrc)
        imgPLANT.Picture = LoadPicture(imgsrc)
        imgPLANT.ScaleMode = 3
        imgPLANT.AutoRedraw = True
        imgPLANT.PaintPicture imgPLANT.Picture, _
        0, 0, imgPLANT.ScaleWidth, imgPLANT.ScaleHeight, _
        0, 0, imgPLANT.Picture.Width / 26.46, _
        imgPLANT.Picture.Height / 26.46
        imgPLANT.Picture = imgPLANT.Image
Text1.Text = rs1!pid
Text2.Text = rs1!pcnm
Text3.Text = rs1!psnm
Text4.Text = rs1!ptype
Text5.Text = rs1!Page
Text6.Text = rs1!pprice
End Function

Function TBenableTrue()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = True
Text6.Enabled = True
End Function
Function TBenableFalse()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
End Function

