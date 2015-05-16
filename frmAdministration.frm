VERSION 5.00
Begin VB.Form frmAdministration 
   Caption         =   "תפריט מנהלה"
   ClientHeight    =   4320
   ClientLeft      =   4680
   ClientTop       =   2985
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdDetails 
      Caption         =   "פרטי בית המלון"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdBooking 
      Caption         =   "הזמנות"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdRooms 
      Caption         =   "חדרים"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdShifts 
      Caption         =   "משמרות"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdWorkers 
      Caption         =   "עובדים"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrefix 
      Caption         =   "קידומות"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCountry 
      Caption         =   "מדינות"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCustomers 
      Caption         =   "לקוחות"
      BeginProperty Font 
         Name            =   "Guttman Frnew"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblTitel 
      Caption         =   "מנהלה - תפריט ראשי"
      BeginProperty Font 
         Name            =   "Guttman Calligraphic"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "frmAdministration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBooking_Click()
    frmBooking.Show
End Sub

Private Sub cmdCountry_Click()
    frmCountry.Show
End Sub

Private Sub cmdCustomers_Click()
    frmCustomers.Show
End Sub

Private Sub cmdDetails_Click()
    frmDetails.Show
End Sub

Private Sub cmdPrefix_Click()
    frmPrefix.Show
End Sub

Private Sub cmdRooms_Click()
    frmRooms.Show
End Sub

Private Sub cmdShifts_Click()
    frmShifts.Show
End Sub

Private Sub cmdWorkers_Click()
    frmWorkers.Show
End Sub


