VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChoseCustomer 
   Caption         =   "בחירת לקוח"
   ClientHeight    =   5250
   ClientLeft      =   4905
   ClientTop       =   2505
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9825
   Begin VB.CommandButton cmdOk 
      Caption         =   "אישור"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "חפש"
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   7560
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   7200
      TabIndex        =   2
      Text            =   "בחר"
      Top             =   960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DGCustomers 
      Bindings        =   "frmChoseCustomer.frx":0000
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1037
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1037
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adcCustomers 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      RecordSource    =   "Customer"
      Caption         =   "לקוחות"
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
   Begin VB.Label lblTextFilter 
      Caption         =   "חפש לפי"
      Height          =   255
      Left            =   8640
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmChoseCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilter As String
Dim numFilter As String

Private Sub cmdFilter_Click()
    strFilter = txtFilter.Text
    numFilter = cmbFilter.ListIndex
    If numFilter = 1 Then
        adcCustomers.Recordset.Filter = "CID = '" & strFilter & "'"
    Else
        If numFilter = 2 Then
            adcCustomers.Recordset.Filter = "CFirstName = '" & strFilter & "'"
        Else
            If numFilter = 3 Then
                adcCustomers.Recordset.Filter = "CLastName = '" & strFilter & "'"
            Else
                If numFilter = 0 Then
                    adcCustomers.Refresh
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    frmEditBooking.txtId.Text = adcCustomers.Recordset(0)
    Unload Me
End Sub

Private Sub Form_Load()
    cmbFilter.AddItem ("בחר")
    cmbFilter.AddItem ("תעודת זהות")
    cmbFilter.AddItem ("שם פרטי")
    cmbFilter.AddItem ("שם משפחה")
    If Len(frmEditBooking.txtId.Text) <> 0 Then
        adcCustomers.Recordset.Find "CID ='" & frmEditBooking.txtId.Text & "'", , adSearchForward, 1
    End If
End Sub

Private Sub cmdPrevious_Click()
    adcCustomers.Recordset.MovePrevious
    If adcCustomers.Recordset.BOF Then
        adcCustomers.Recordset.MoveLast
    End If
End Sub

Private Sub cmdFirst_Click()
    adcCustomers.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcCustomers.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
    adcCustomers.Recordset.MoveNext
    If adcCustomers.Recordset.EOF Then
        adcCustomers.Recordset.MoveFirst
    End If
End Sub
