VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPrefix 
   Caption         =   "קידומת פלאפון"
   ClientHeight    =   4545
   ClientLeft      =   6060
   ClientTop       =   2985
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6375
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "הוסף"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "עדכן"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "בטל עדכון"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "תפריט ראשי"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPrefix 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DGPrefix 
      Bindings        =   "frmPrefix.frx":0000
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
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
   Begin MSAdodcLib.Adodc adcPrefix 
      Height          =   330
      Left            =   480
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "Prefix"
      Caption         =   "קידומת"
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
      Caption         =   "קידומת"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrefix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim index As Integer
Dim str As String
Dim flag As Boolean

Private Sub cmdAdd_Click()
    str = txtPrefix.Text
    adcPrefix.Recordset.MoveFirst
    flag = False
    While flag = False And adcPrefix.Recordset.EOF = False
        If InStr(adcPrefix.Recordset(1), str) = 1 Then
            flag = True
        End If
        adcPrefix.Recordset.MoveNext
    Wend
    adcPrefix.Recordset.MovePrevious
    If flag = False Then
        adcPrefix.Recordset.MoveLast
        index = adcPrefix.Recordset(0)
        adcPrefix.Recordset.AddNew
        adcPrefix.Recordset(0) = index + 1
        adcPrefix.Recordset(1) = str
    Else
        MsgBox ("קידומת זו כבר נמצאת ברשימת הקידומות")
    End If
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdMenu_Click()
    frmMainMenu.Show
    Unload Me
End Sub

Private Sub cmdPrevious_Click()
    adcPrefix.Recordset.MovePrevious
    If adcPrefix.Recordset.BOF Then
        adcPrefix.Recordset.MoveLast
    End If
End Sub

Private Sub cmdCancel_Click()
    adcPrefix.Recordset.CancelUpdate
End Sub

Private Sub cmdFirst_Click()
    adcPrefix.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcPrefix.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
    adcPrefix.Recordset.MoveNext
    If adcPrefix.Recordset.EOF Then
        adcPrefix.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdUp_Click()
    adcPrefix.Recordset.Update
End Sub

