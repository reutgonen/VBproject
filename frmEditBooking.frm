VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditBooking 
   Caption         =   "עריכת הזמנה"
   ClientHeight    =   5475
   ClientLeft      =   5265
   ClientTop       =   2790
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdClose 
      Caption         =   "סגור"
      Height          =   375
      Left            =   8040
      TabIndex        =   45
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtConnectToRoomBooking 
      DataField       =   "RBBookingCode"
      DataSource      =   "adcRoomBooking"
      Height          =   285
      Left            =   3960
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adcRoomBooking 
      Height          =   330
      Left            =   4560
      Top             =   5040
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "RoomsBooking"
      Caption         =   "חדרים - הזמנות"
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
   Begin VB.CommandButton cmdCalc 
      Caption         =   "חשב"
      Height          =   255
      Left            =   1680
      TabIndex        =   43
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtConnectToRooms 
      DataField       =   "RNumber"
      DataSource      =   "adcRooms"
      Height          =   285
      Left            =   3240
      TabIndex        =   42
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSumNights 
      Height          =   285
      Left            =   2520
      TabIndex        =   41
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adcRooms 
      Height          =   330
      Left            =   2400
      Top             =   5040
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton cmdRoom 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   39
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   38
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   7080
      TabIndex        =   35
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRooms 
      Caption         =   "הצג אפשרויות"
      Height          =   435
      Left            =   6360
      TabIndex        =   33
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdChoseCustomer 
      Caption         =   "בחר"
      Height          =   255
      Left            =   5760
      TabIndex        =   32
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "שמור"
      Height          =   495
      Left            =   1440
      TabIndex        =   31
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   30
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "עדכן"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   29
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdCanUpdate 
      Caption         =   "בטל עדכון"
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "בחר"
      Height          =   255
      Left            =   5760
      TabIndex        =   27
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdCheckIn 
      Caption         =   "בחר"
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adcBooking 
      Height          =   330
      Left            =   120
      Top             =   5040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "הזמנה"
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
   Begin VB.CheckBox optIsPaid 
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox optIsCan 
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   3600
      Width           =   255
   End
   Begin VB.Frame framePayment 
      Caption         =   "תשלום"
      Height          =   615
      Left            =   240
      TabIndex        =   21
      Top             =   1680
      Width           =   3375
      Begin VB.OptionButton optKesh 
         Caption         =   "מזומן"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optCredit 
         Caption         =   "כרטיס אשראי"
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtValidityCredit 
      DataField       =   "BCreditCardValidity"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtCreditNum 
      DataField       =   "BCreditCard"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   2400
      TabIndex        =   19
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtPrice 
      DataField       =   "BPrice"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   2400
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtCheckOut 
      DataField       =   "BDateCheckOut"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtCheckIn 
      DataField       =   "BDateCheckIn"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtNumOfCustomers 
      DataField       =   "BGuestAmount"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   6360
      TabIndex        =   15
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtId 
      DataField       =   "BIDCustomer"
      DataSource      =   "adcBooking"
      Height          =   285
      Left            =   6480
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      DataField       =   "BCode"
      DataSource      =   "adcBooking"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblYourRoom 
      Caption         =   "החדרים שלך"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblTitel 
      Caption         =   "עריכת הזמנה"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblTextCan 
      Caption         =   "בוטל"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblTextPaid 
      Caption         =   "שולם"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblTextValidityCredit 
      Caption         =   "תוקף"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label c 
      Caption         =   "אופן התשלום"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label lblTextCreditNum 
      Caption         =   "מספר כרטיס אשראי"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblTextRooms 
      Caption         =   "חדרים"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblTextPrice 
      Caption         =   "סכום לתשלום"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblTextCheckOut 
      Caption         =   "תאריך עזיבה"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTextCheckIn 
      Caption         =   "תאריך הגעה"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTextNumOfCustomers 
      Caption         =   "מספר אורחים"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTextId 
      Caption         =   "ת.ז. לקוח"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTextCode 
      Caption         =   "קוד הזמנה"
      Height          =   255
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim code As Integer
Dim sumNight As Integer

Private Sub cmdBack_Click()
    Load frmBooking
    adcBooking.Refresh
    frmBooking.Refresh
    frmBooking.Show
    Unload Me
End Sub

Private Sub cmdCalc_Click()
    Call calcprice
End Sub

Private Sub cmdCanUpdate_Click()
    Unload Me
End Sub

Private Sub cmdCheckIn_Click()
    frmCalendar.Show
End Sub

Private Sub cmdCheckOut_Click()
    frmCalendar.Show
End Sub

Private Sub cmdChoseCustomer_Click()
    frmChoseCustomer.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRooms_Click()
    Load frmFreeRoomsList
    frmFreeRoomsList.Show
End Sub

Private Sub cmdSave_Click()
    If optCredit = True Then
        adcBooking.Recordset(6) = "כרטיס אשראי"
    Else
        adcBooking.Recordset(6) = "מזומן"
    End If
    If optIsPaid.Value = 1 Then
        adcBooking.Recordset(9) = True
    Else
        adcBooking.Recordset(9) = False
    End If
    If optIsCan.Value = 1 Then
        adcBooking.Recordset(10) = True
    Else
        adcBooking.Recordset(10) = False
    End If
    adcBooking.Recordset.Update
    Dim index As Integer
    For index = 0 To 5
        If cmdRoom(index).Visible = True Then
            adcRoomBooking.Recordset.AddNew
            adcRoomBooking.Recordset(0) = txtCode.Text
            adcRoomBooking.Recordset(1) = index + 1
            adcRoomBooking.Recordset.Update
        End If
    Next
    cmdUpdate.Enabled = True
    cmdBack.Enabled = True
    cmdCanUpdate = False
    cmdSave.Enabled = False
    txtCode.Enabled = False
    txtId.Enabled = False
    txtNumOfCustomers.Enabled = False
    txtCheckIn.Enabled = False
    txtCheckOut.Enabled = False
    framePayment.Enabled = False
    txtCreditNum.Enabled = False
    txtValidityCredit.Enabled = False
    optIsCan.Enabled = False
    optIsPaid.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
    cmdSave.Enabled = True
    cmdUpdate.Enabled = False
    cmdBack.Enabled = False
    cmdCanUpdate.Enabled = True
    txtCode.Enabled = True
    txtId.Enabled = True
    txtNumOfCustomers.Enabled = True
    txtCheckIn.Enabled = True
    txtCheckOut.Enabled = True
    framePayment.Enabled = True
    txtCreditNum.Enabled = True
    txtValidityCredit.Enabled = True
    optIsCan.Enabled = True
    optIsPaid.Enabled = True
End Sub



Private Sub Form_Load()
    If frmBooking.cmdEdit = True Then
        adcBooking.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
        adcBooking.CommandType = adCmdText
        adcBooking.RecordSource = "Select * From Booking Where BCode = " & frmBooking.adcBooking.Recordset(0) & ""
        adcBooking.Refresh
        If adcBooking.Recordset(5) = "כרטיס אשראי" Then
            optCredit = True
        Else
            optKesh = True
        End If
        If adcBooking.Recordset(8) = True Then
            optIsPaid.Value = 1
        Else
            optIsPaid.Value = 0
        End If
        If adcBooking.Recordset(9) = True Then
            optIsCan.Value = 1
        Else
            optIsCan.Value = 0
        End If
    End If
    If frmBooking.cmdNew = True Then
        adcBooking.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
        adcBooking.CommandType = adCmdText
        adcBooking.RecordSource = "Select * From Booking"
        adcBooking.Refresh
        adcBooking.Recordset.MoveLast
        code = adcBooking.Recordset(0)
        adcBooking.Recordset.AddNew
        txtCode.Text = code + 1
    End If
    Unload frmBooking
End Sub


Private Sub txtNumOfCustomers_LostFocus()
    If IsNumeric(txtNumOfCustomers.Text) = False Then
        MsgBox "אין להזין תווים במספר אורחים"
        txtNumOfCustomers.SelStart = 0
        txtNumOfCustomers.SelLength = Len(txtNumOfCustomers.Text)
        txtNumOfCustomers.SetFocus
    End If
End Sub

Private Sub calcprice()
    Dim i As Integer
    Dim calcprice As Integer
    calcprice = 0
    sumNight = txtSumNights.Text
    For i = 0 To 5
        If cmdRoom(i).Visible = True Then
            adcRooms.Recordset.MoveFirst
            While adcRooms.Recordset.EOF = False
                If adcRooms.Recordset(0) = i + 1 Then
                    calcprice = calcprice + adcRooms.Recordset(4) * sumNight
                End If
                adcRooms.Recordset.MoveNext
            Wend
        End If
    Next
    txtPrice.Text = calcprice
End Sub

