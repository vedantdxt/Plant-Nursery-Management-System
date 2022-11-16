VERSION 5.00
Begin VB.Form invoicefrm 
   Caption         =   "Invoice"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Print_BTN 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   16560
      TabIndex        =   63
      Top             =   9600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7320
      TabIndex        =   62
      Top             =   9600
      Width           =   1500
   End
   Begin VB.CommandButton Purchase_BTN 
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   14160
      TabIndex        =   33
      Top             =   9600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton AddPlant_BTN 
      Caption         =   "Add Plant"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11040
      TabIndex        =   15
      Top             =   9600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton AddCust_BTN 
      Caption         =   "Add Customer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3600
      TabIndex        =   14
      Top             =   9600
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Caption         =   "SUNSHINE NURSERY"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   15120
      TabIndex        =   5
      Top             =   1500
      Width           =   3255
      Begin VB.Label NurseryAddress_LB 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2835
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   16335
      Begin VB.Label PAMT_LB7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   61
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Label PQTY_LB7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   60
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   59
         Top             =   3840
         Width           =   1005
      End
      Begin VB.Label PNM_LB7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   58
         Top             =   3840
         Width           =   8775
      End
      Begin VB.Label PID_LB7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   3840
         Width           =   795
      End
      Begin VB.Label PAMT_LB3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   56
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label PAMT_LB6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   55
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label PAMT_LB5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   54
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label PAMT_LB4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   53
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label PQTY_LB6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   52
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label PQTY_LB5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   51
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label PQTY_LB4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   50
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label PQTY_LB3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   49
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   48
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   47
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   46
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   45
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label PPRICE_LB2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   44
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label PNM_LB6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   43
         Top             =   3360
         Width           =   8775
      End
      Begin VB.Label PNM_LB5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   2880
         Width           =   8775
      End
      Begin VB.Label PNM_LB4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   41
         Top             =   2400
         Width           =   8775
      End
      Begin VB.Label PNM_LB3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   40
         Top             =   1920
         Width           =   8775
      End
      Begin VB.Label PID_LB6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label PID_LB5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label PID_LB4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label PID_LB3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label TAMT_LB 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   14880
         TabIndex        =   32
         Top             =   4680
         Width           =   1005
      End
      Begin VB.Label TQTY_LB 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   13080
         TabIndex        =   31
         Top             =   4680
         Width           =   1005
      End
      Begin VB.Label TUNITCOST_LB 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   11280
         TabIndex        =   30
         Top             =   4680
         Width           =   1005
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   29
         Top             =   4680
         Width           =   1995
      End
      Begin VB.Line Line6 
         X1              =   16320
         X2              =   0
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label PAMT_LB2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   21
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label PAMT_LB1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   20
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label PQTY_LB2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   19
         Top             =   1440
         Width           =   1000
      End
      Begin VB.Label PQTY_LB1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   13080
         TabIndex        =   18
         Top             =   960
         Width           =   1000
      End
      Begin VB.Label PPRICE_LB1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   11280
         TabIndex        =   17
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "DESCRIPTION"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   300
         Width           =   2000
      End
      Begin VB.Label PNM_LB2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   8775
      End
      Begin VB.Label PNM_LB1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   960
         Width           =   8775
      End
      Begin VB.Label PID_LB2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label PID_LB1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   795
      End
      Begin VB.Line Line5 
         X1              =   1320
         X2              =   1320
         Y1              =   120
         Y2              =   5160
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "PLANT ID"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1005
      End
      Begin VB.Line Line4 
         X1              =   16320
         X2              =   0
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   14640
         TabIndex        =   8
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "QUANTITY"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12960
         TabIndex        =   7
         Top             =   300
         Width           =   1500
      End
      Begin VB.Line Line3 
         X1              =   14520
         X2              =   14520
         Y1              =   120
         Y2              =   5160
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "UNIT COST"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11040
         TabIndex        =   6
         Top             =   300
         Width           =   1500
      End
      Begin VB.Line Line2 
         X1              =   12840
         X2              =   12840
         Y1              =   120
         Y2              =   5280
      End
      Begin VB.Line Line1 
         X1              =   10680
         X2              =   10680
         Y1              =   120
         Y2              =   5160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BILLED TO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2040
      TabIndex        =   2
      Top             =   1500
      Width           =   5055
      Begin VB.Label CustNUM_LB 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2520
         TabIndex        =   27
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label Label11 
         Caption         =   "CUSTOMER CONTACT :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   2000
      End
      Begin VB.Label CustADD_LB 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "CUSTOMER ADDRESS :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label Label8 
         Caption         =   "CUSTOMER NAME :"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label CustID_LB 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1995
      End
   End
   Begin VB.Label PURDATE_LB 
      Alignment       =   2  'Center
      Caption         =   "00-00-0000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   3960
      TabIndex        =   35
      Top             =   3720
      Width           =   1500
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "PURCHASED ON : "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   34
      Top             =   3720
      Width           =   2000
   End
   Begin VB.Label BILLNO_LB 
      Alignment       =   2  'Center
      Caption         =   "000011"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Bno_LB 
      Alignment       =   2  'Center
      Caption         =   "INVOICE NUMBER :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "INVOICE"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "invoicefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs As Recordset
Dim currentqty As Integer
Dim LB As Integer
Dim CustID As Integer
Dim LCPID As Integer
Private Sub AddPlant_BTN_Click()
c = CLB()
Set con = New Connection
Set rs = New Recordset
con.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs.Open "select * from plant", con, adOpenStatic, adLockOptimistic
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
n = Val(InputBox("Enter Plant ID"))
If n <> 0 Then

    s = "Select * from plant where pid=" & n
    rsSearch.Open s, con, adOpenStatic, adLockOptimistic
    If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
        MsgBox "Plant Record Not Found."
    Else
        While Not rsSearch.EOF
        Me.Controls("PID_LB" & LB).Caption = rsSearch!pid
        Me.Controls("PNM_LB" & LB).Caption = rsSearch!pcnm
        Me.Controls("PPRICE_LB" & LB).Caption = rsSearch!pprice
        'and soon
        rsSearch.MoveNext
        Wend

        c = qty()

        quan = Val(InputBox("Enter Quantity" & vbNewLine & "Current Quantity = " & currentqty, "Quantity"))
        If quan < currentqty Then
            Me.Controls("PQTY_LB" & LB).Caption = quan
            Me.Controls("PAMT_LB" & LB).Caption = Val(Me.Controls("PPRICE_LB" & LB)) * Val(Me.Controls("PQTY_LB" & LB).Caption)
            c = Calculate()
        Else
            MsgBox ("Requested stock is not available")
            Me.Controls("PID_LB" & LB).Caption = ""
            Me.Controls("PNM_LB" & LB).Caption = ""
            Me.Controls("PPRICE_LB" & LB).Caption = ""
            LB = LB - 1
        End If
    End If
End If
End Sub


Private Sub cmdsearch_Click()
searchchoice = InputBox("1 for search by BID" & vbNewLine & "2 for search by CID")
If searchchoice = 1 Then
    cs = searchbid()
ElseIf searchchoice = 2 Then
    cs = searchcust()
End If
End Sub



Private Sub Print_BTN_Click()
PrintForm
End Sub

Private Sub Purchase_BTN_Click()
If LB <> 0 Then
ch = MsgBox("Really Wanna Made Purchase?", vbQuestion + vbYesNo, "PURCHASE")
If (ch = vbYes) Then
For plantqty = 1 To LB Step 1
Set con2 = New Connection
Set rs2 = New Recordset
Y = Val(Me.Controls("PID_LB" & plantqty).Caption)
s = "select * from stock where pid=" & Y
con2.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs2.Open s, con2, adOpenStatic, adLockOptimistic
quan = Val(Me.Controls("PQTY_LB" & plantqty))
rs2!squantity = rs2!squantity - quan
rs2.Update
rs2.Close
con2.Close

Set con4 = New Connection
Set rs4 = New Recordset
    s = "select * from custplant"
    con4.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
    rs4.Open s, con4, adOpenStatic, adLockOptimistic
    rs4.AddNew
    rs4!cid = CustID
    rs4!pid = Val(Me.Controls("PID_LB" & plantqty))
    rs4.Update
    rs4.Close
    con4.Close
Next

Set con3 = New Connection
Set rs3 = New Recordset
    s = "select * from bill where bid=" & Val(BILLNO_LB.Caption)
    con3.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
    rs3.Open s, con3, adOpenStatic, adLockOptimistic
    rs3.AddNew
    rs3!bid = BILLNO_LB.Caption
    rs3!cid = CustID
    rs3!amount = Val(TAMT_LB.Caption)
    rs3!purdate = Format(Now, "Short Date")
    rs3!purquantity = Val(TQTY_LB.Caption)
    rs3.Update
    rs3.Close
    con3.Close
MsgBox ("Purchase Successfull!")
Print_BTN.Visible = True
End If
Else
ch = MsgBox("No Purchase Made!", vbOKOnly, "Error!")
End If
End Sub
Private Sub Form_Load()
LB = 0
PURDATE_LB.Caption = Now()
Set con = New Connection
Set rs = New Recordset
con.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs.Open "select * from cust", con, adOpenStatic, adLockOptimistic
NurseryAddress_LB.Caption = "Plot No. 7, Sunshine Nursery," & vbNewLine & "Sunny Roadways, Destiny Lane," & vbNewLine & "Nashik - 422024" & vbNewLine & "Email - nursery@sunshine.com" & vbNewLine & "Website - www.sunshinenursery.com" & vbNewLine & "Landline - (0253)-426426" & vbNewLine & "Mobile No. - (+91)-9773693690"
c = getbillno()
'c = getcpid()
End Sub
Private Sub AddCust_BTN_Click()
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
CustID = Val(InputBox("Enter Customer ID"))
If CustID <> 0 Then
    s = "Select * from cust where cid=" & CustID
    rsSearch.Open s, con, adOpenStatic, adLockOptimistic
    If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
    MsgBox "Customer Record Not Found."
    Else
        While Not rsSearch.EOF
        CustID_LB.Caption = rsSearch!cname
        CustADD_LB.Caption = rsSearch!cadd
        CustNUM_LB.Caption = rsSearch!ccont
        'and soon
        rsSearch.MoveNext
        Wend
        AddPlant_BTN.Visible = True
    End If
End If
End Sub

Function qty()
Set con = New Connection
Set rs = New Recordset
con.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rs.Open "select * from stock", con, adOpenStatic, adLockOptimistic
Dim s, nm As String
Dim rsSearch As Recordset
Set rsSearch = New Recordset
s = "Select * from stock where pid=" & Val(Me.Controls("PID_LB" & LB).Caption)
rsSearch.Open s, con, adOpenStatic, adLockOptimistic
If (rsSearch.EOF = True) And (rsSearch.BOF = True) Then
MsgBox "Record Not Found."
End If
While Not rsSearch.EOF
currentqty = rsSearch!squantity
'and soon
rsSearch.MoveNext
Wend
End Function

Function CLB()
If PID_LB1.Caption = "" Then
LB = "1"
ElseIf PID_LB2.Caption = "" Then
LB = "2"
ElseIf PID_LB3.Caption = "" Then
LB = "3"
ElseIf PID_LB4.Caption = "" Then
LB = "4"
ElseIf PID_LB5.Caption = "" Then
LB = "5"
ElseIf PID_LB6.Caption = "" Then
LB = "6"
ElseIf PID_LB7.Caption = "" Then
LB = "7"
End If
End Function

Function getbillno()
Set conLR = New Connection
Set rsLR = New Recordset
conLR.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rsLR.Open "select * from bill", conLR, adOpenStatic, adLockOptimistic

Dim rsSearchLR As Recordset
Set rsSearchLR = New Recordset
s = "Select * from bill"
rsSearchLR.Open s, con, adOpenStatic, adLockOptimistic
rsSearchLR.MoveLast
BILLNO_LB.Caption = rsSearchLR!bid + 1
End Function

Function getcpid()
Set conLCPID = New Connection
Set rsLCPID = New Recordset
conLCPID.Open "provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Work\VB\24-25.accdb"
rsLCPID.Open "select * from custplant", conLCPID, adOpenStatic, adLockOptimistic

Dim rsSearchLCPID As Recordset
Set rsSearchLCPID = New Recordset
s = "Select * from custplant"
rsSearchLCPID.Open s, con, adOpenStatic, adLockOptimistic
rsSearchLCPID.MoveLast
LCPID = rsSearchLCPID!cpid + 1
End Function

Function Calculate()
TUNITCOST_LB.Caption = Val(TUNITCOST_LB.Caption) + Val(Me.Controls("PPRICE_LB" & LB).Caption)
TQTY_LB.Caption = Val(TQTY_LB.Caption) + Val(Me.Controls("PQTY_LB" & LB).Caption)
TAMT_LB.Caption = Val(TAMT_LB.Caption) + Val(Me.Controls("PAMT_LB" & LB).Caption)
 Purchase_BTN.Visible = True
End Function

Function searchcust()
Dim s, nm As String
Dim rsCS As Recordset
Set rsCS = New Recordset
CSID = Val(InputBox("Enter Customer ID"))
If CSID <> 0 Then
    s = "Select * from custplant where cid=" & CSID
    rsCS.Open s, con, adOpenStatic, adLockOptimistic
    If (rsCS.EOF = True) And (rsCS.BOF = True) Then
    MsgBox "Customer Record Not Found."
    Else
        While Not rsCS.EOF
        CSPID = rsCS!pid
        rsCS.MoveNext
        Wend
    End If
End If

Dim sts As String
Dim rsCPS As Recordset
Set rsCPS = New Recordset
    sts = "Select * from plant where pid=" & CSPID
    rsCPS.Open sts, con, adOpenStatic, adLockOptimistic
    If (rsCPS.EOF = True) And (rsCPS.BOF = True) Then
    MsgBox "Record Not Found."
    Else
        While Not rsCPS.EOF
        pid = rsCPS!pid
        pcnm = rsCPS!pcnm
        pprice = rsCPS!pprice
        rsCPS.MoveNext
        Wend
    End If
    
Dim stsmnt As String
Dim rsCDS As Recordset
Set rsCDS = New Recordset
    stsmnt = "Select * from cust where cid=" & CSID
    rsCDS.Open stsmnt, con, adOpenStatic, adLockOptimistic
    If (rsCDS.EOF = True) And (rsCDS.BOF = True) Then
    MsgBox "Record Not Found."
    Else
        While Not rsCDS.EOF
        MsgBox ("Customer Details" & vbNewLine & vbNewLine & "Customer ID : " & rsCDS!cid & vbNewLine & "Customer Name : " & rsCDS!cname & vbNewLine & "Customer Address : " & rsCDS!cadd & vbNewLine & "Customer Contact Number : " & rsCDS!ccont & vbNewLine & vbNewLine & "Plant Details" & vbNewLine & vbNewLine & "Plant ID : " & pid & vbNewLine & "Plant Name : " & pcnm & vbNewLine & "Plant Price : " & pprice)
        rsCDS.MoveNext
        Wend
    End If
End Function

Function searchbid()
Dim s, nm As String
Dim rsBS As Recordset
Set rsBS = New Recordset
bid = Val(InputBox("Enter Bill ID"))
If bid <> 0 Then
    s = "Select * from bill where bid=" & bid
    rsBS.Open s, con, adOpenStatic, adLockOptimistic
    If (rsBS.EOF = True) And (rsBS.BOF = True) Then
    MsgBox "Bill Record Not Found."
    Else
        While Not rsBS.EOF
        BILLNO_LB.Caption = rsBS!bid
        BCustid = rsBS!cid
        TAMT_LB.Caption = rsBS!amount
        PURDATE_LB.Caption = rsBS!purdate
        TQTY_LB.Caption = rsBS!purquantity
        rsBS.MoveNext
        Wend
    End If
End If

Dim q As String
Dim rsBCS As Recordset
Set rsBCS = New Recordset
    q = "Select * from cust where cid=" & BCustid
    rsBCS.Open q, con, adOpenStatic, adLockOptimistic
    If (rsBCS.EOF = True) And (rsBCS.BOF = True) Then
    MsgBox "Customer Record Not Found."
    Else
        While Not rsBCS.EOF
        CustID_LB.Caption = rsBCS!cname
        CustADD_LB = rsBCS!cadd
        CustNUM_LB = rsBCS!ccont
        rsBCS.MoveNext
        Wend
    End If
    
Dim qu As String
Dim rsBCPS As Recordset
Set rsBCPS = New Recordset
    qu = "Select * from custplant where cid=" & BCustid
    rsBCPS.Open qu, con, adOpenStatic, adLockOptimistic
    If (rsBCPS.EOF = True) And (rsBCPS.BOF = True) Then
    MsgBox "Record Not Found."
    Else
        While Not rsBCPS.EOF
        BCPid = rsBCPS!pid
        rsBCPS.MoveNext
        Wend
    End If
    
Dim query As String
Dim rsBPS As Recordset
Set rsBPS = New Recordset
    query = "Select * from plant where pid=" & BCPid
    rsBPS.Open query, con, adOpenStatic, adLockOptimistic
    If (rsBPS.EOF = True) And (rsBPS.BOF = True) Then
    MsgBox "Record Not Found."
    Else
        While Not rsBPS.EOF
        PID_LB1.Caption = rsBPS!pid
        PNM_LB1.Caption = rsBPS!pcnm
        PPRICE_LB1.Caption = TAMT_LB.Caption / 2
        PQTY_LB1.Caption = TQTY_LB.Caption
        PAMT_LB1.Caption = TAMT_LB.Caption
        TUNITCOST_LB.Caption = TAMT_LB.Caption / 2
        rsBPS.MoveNext
        Wend
    End If
End Function

