VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCountry 
   Caption         =   "������"
   ClientHeight    =   4365
   ClientLeft      =   4680
   ClientTop       =   2790
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7005
   Begin VB.TextBox txtCountry 
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdMenu 
      Caption         =   "����� ����"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��� �����"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "����"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adcCounty 
      Height          =   330
      Left            =   120
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
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
      Enabled         =   0
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Country"
      Caption         =   "������"
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
   Begin MSDataGridLib.DataGrid AGCountry 
      Bindings        =   "frmCountry.frx":0000
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
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
   Begin VB.Label lblTitel 
      Caption         =   "������"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim index As Integer
Dim str As String
Dim flag As Boolean

Private Sub cmdAdd_Click()
    str = txtCountry.Text
    adcCounty.Recordset.MoveFirst
    flag = False
    While flag = False And adcCounty.Recordset.EOF = False
        If InStr(adcCounty.Recordset(1), str) = 1 Then
            flag = True
        End If
        adcCounty.Recordset.MoveNext
    Wend
    adcCounty.Recordset.MovePrevious
    If flag = False Then
        adcCounty.Recordset.MoveLast
        index = adcCounty.Recordset(0)
        adcCounty.Recordset.AddNew
        adcCounty.Recordset(0) = index + 1
        adcCounty.Recordset(1) = str
    Else
        MsgBox ("����� �� ��� ����� ������ �������")
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
    adcCounty.Recordset.MovePrevious
    If adcCounty.Recordset.BOF Then
        adcCounty.Recordset.MoveLast
    End If
End Sub

Private Sub cmdCancel_Click()
    adcCounty.Recordset.CancelUpdate
End Sub

Private Sub cmdFirst_Click()
    adcCounty.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
    adcCounty.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
    adcCounty.Recordset.MoveNext
    If adcCounty.Recordset.EOF Then
        adcCounty.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdUp_Click()
    adcCounty.Recordset.Update
End Sub

