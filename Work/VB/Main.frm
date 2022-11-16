VERSION 5.00
Begin VB.MDIForm mainfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   Picture         =   "Main.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu billfrm 
      Caption         =   "BILLING INFO"
   End
   Begin VB.Menu customer 
      Caption         =   "CUSTOMER DETAILS"
   End
   Begin VB.Menu plants 
      Caption         =   "PLANTS DETAILS"
   End
   Begin VB.Menu stock 
      Caption         =   "STOCK DETAILS"
   End
   Begin VB.Menu employee 
      Caption         =   "EMPLOYEE DETAILS"
   End
   Begin VB.Menu logout 
      Caption         =   "LOG OUT"
   End
   Begin VB.Menu about 
      Caption         =   "ABOUT US"
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As Connection
Dim rs1 As Recordset

Private Sub about_Click()
f = formhide()
aboutfrm.Show
End Sub

Private Sub billfrm_Click()
f = formhide()
invoicefrm.Show
End Sub

Private Sub customer_Click()
f = formhide()
custfrm.Show
End Sub

Private Sub employee_Click()
f = formhide()
empfrm.Show
End Sub

Private Sub logout_Click()
f = formhide()
loginfrm.Show
End Sub
Private Sub MDIForm_Load()
Set con1 = New Connection
Set rs1 = New Recordset
con1.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs1.Open "select * from stock", con1, adOpenStatic, adLockOptimistic

Dim s, alert As String
cnt = 0
alert = ""
Dim rsSearch As Recordset
Set rsSearch = New Recordset
s = "Select * from stock where squantity between 5 and 0"
rsSearch.Open s, con1, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
End If
While Not rsSearch.EOF
cnt = cnt + 1
If rsSearch.RecordCount = 1 Then
alert = alert & rsSearch!sid & ""
ElseIf rsSearch.RecordCount = cnt Then
alert = alert & rsSearch!sid & ""
Else
alert = alert & rsSearch!sid & ","
End If
'and soon
rsSearch.MoveNext
Wend
If alert <> "" Then
Y = MsgBox(alert & " Has Less than 6 Quanties Left!", vbCritical = vbOKOnly, "ALERT!")
End If
End Sub

Private Sub plants_Click()
f = formhide()
plantfrm.Show
End Sub

Function formhide()
custfrm.Hide
plantfrm.Hide
aboutfrm.Hide
stockfrm.Hide
empfrm.Hide
invoicefrm.Hide
End Function

Private Sub stock_Click()
f = formhide()
stockfrm.Show
End Sub
