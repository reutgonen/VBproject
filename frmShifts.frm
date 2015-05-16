VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmShifts 
   Caption         =   "משמרות"
   ClientHeight    =   4725
   ClientLeft      =   5265
   ClientTop       =   2985
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7515
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DGShifts 
      Bindings        =   "frmShifts.frx":0000
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3836
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
   Begin MSAdodcLib.Adodc adcShifts 
      Height          =   330
      Left            =   360
      Top             =   4200
      Visible         =   0   'False
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
      RecordSource    =   "Shifts"
      Caption         =   "משמרות"
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
      Caption         =   "משמרות"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmShifts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNext_Click()
    adcShifts.Recordset.MoveNext
    If adcShifts.Recordset.EOF Then
        adcShifts.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdPrevious_Click()
    adcShifts.Recordset.MovePrevious
    If adcShifts.Recordset.BOF Then
        adcShifts.Recordset.MoveLast
    End If
End Sub

Private Sub cmdFirst_Click()
    adcShifts.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcShifts.Recordset.MoveLast
End Sub

