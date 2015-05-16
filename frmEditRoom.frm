VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditRoom 
   Caption         =   "עריכת חדר"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCanUpdate 
      Caption         =   "בטל עדכון"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "עדכן"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "שמור"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox optIsSea 
      Caption         =   "נוף לים"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   2640
      Width           =   255
   End
   Begin VB.CheckBox optIsBed 
      Caption         =   "מיטה נוספת"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "RPrice"
      DataSource      =   "adcRooms"
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtMaxGuestAmount 
      DataField       =   "RGuestAmount"
      DataSource      =   "adcRooms"
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtRoomNumber 
      DataField       =   "RNumber"
      DataSource      =   "adcRooms"
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adcRooms 
      Height          =   330
      Left            =   240
      Top             =   3840
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Rooms"
      Caption         =   "חדרים"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblTextPrice 
      Caption         =   "מחיר ללילה"
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblTextIsSea 
      Caption         =   "נוף לים"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblTextOntherBed 
      Caption         =   "מיטה נוספת"
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblTextMaxGuestAmount 
      Caption         =   "מספר אורחים מקסימלי"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblTextRoomNumber 
      Caption         =   "מספר חדר"
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblTitel 
      Caption         =   "עריכת חדר"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmEditRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    frmRooms.Show
    adcRooms.Refresh
    frmRooms.Refresh
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If optIsBed.Value = 1 Then
        adcRooms.Recordset(2) = True
    Else
        adcRooms.Recordset(2) = False
    End If
    If optIsSea.Value = 1 Then
        adcRooms.Recordset(3) = True
    Else
        adcRooms.Recordset(3) = True
    End If
    adcRooms.Recordset.Update
    adcRooms.Refresh
End Sub

Private Sub Form_Load()
    If frmRooms.cmdEdit = True Then
        adcRooms.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
        adcRooms.CommandType = adCmdText
        adcRooms.RecordSource = "SELECT * FROM Rooms WHERE RNumber = " & frmRooms.adcRooms.Recordset(0) & ""
        adcRooms.Refresh
        Unload frmRooms
        If adcRooms.Recordset(2) = True Then
            optIsBed.Value = 1
        Else
            optIsBed.Value = 0
        End If
        If adcRooms.Recordset(3) = True Then
            optIsSea.Value = 1
        Else
            optIsSea.Value = 0
        End If
    End If
    If frmRooms.cmdAdd = True Then
        adcRooms.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
        adcRooms.CommandType = adCmdText
        adcRooms.RecordSource = "SELECT * FROM Rooms"
        adcRooms.Refresh
        Unload frmRooms
        adcRooms.Recordset.AddNew
    End If
End Sub

Private Sub txtRoomNumber_LostFocus()
    If IsNumeric(txtRoomNumber.Text) Then
        MsgBox "מספר חדר מורכב אך ורק ממספרים"
        txtRoomNumber.SetFocus
    End If
End Sub
