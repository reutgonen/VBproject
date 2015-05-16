VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRooms 
   Caption         =   "חדרים"
   ClientHeight    =   4500
   ClientLeft      =   6255
   ClientTop       =   2400
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdEdit 
      Caption         =   "עריכת חדר"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DGRooms 
      Bindings        =   "frmRooms.frx":0000
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3625
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
   Begin MSAdodcLib.Adodc adcRooms 
      Height          =   330
      Left            =   360
      Top             =   3840
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
   Begin VB.Label lblTitel 
      Caption         =   "חדרים"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEdit_Click()
    frmEditRoom.Show
    Unload Me
End Sub

Private Sub cmdNext_Click()
    adcCustomers.Recordset.MoveNext
    If adcRooms.Recordset.EOF Then
        adcRooms.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdPrevious_Click()
    adcRooms.Recordset.MovePrevious
    If adcRooms.Recordset.BOF Then
        adcRooms.Recordset.MoveLast
    End If
End Sub

Private Sub cmdFirst_Click()
    adcRooms.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcRooms.Recordset.MoveLast
End Sub

