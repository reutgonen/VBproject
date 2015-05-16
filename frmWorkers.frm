VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmWorkers 
   Caption         =   "עובדים"
   ClientHeight    =   5205
   ClientLeft      =   3885
   ClientTop       =   2790
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   11340
   Begin VB.CommandButton cmdMenu 
      Caption         =   "תפריט ראשי"
      Height          =   375
      Left            =   10440
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Height          =   375
      Left            =   10560
      TabIndex        =   6
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "עריכה"
      Height          =   375
      Left            =   10080
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "הוסף"
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   4080
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DGWorkers 
      Bindings        =   "frmWorkers.frx":0000
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc adcWorkers 
      Height          =   375
      Left            =   240
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Enabled         =   0
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Workers"
      Caption         =   "עובדים"
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
   Begin VB.Label lblTitel 
      Caption         =   "עובדים"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmWorkers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean

Private Sub cmdAdd_Click()
    frmEditWorker.Show
    flag = False
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    adcWorkers.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcWorkers.Recordset.MoveLast
End Sub

Private Sub cmdMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub cmdNext_Click()
    adcWorkers.Recordset.MoveNext
    If adcWorkers.Recordset.EOF Then
        adcWorkers.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdPrevious_Click()
    adcWorkers.Recordset.MovePrevious
    If adcWorkers.Recordset.BOF Then
        adcWorkers.Recordset.MoveLast
    End If
End Sub

Private Sub cmdUp_Click()
    frmEditWorker.Show
    flag = True
End Sub


