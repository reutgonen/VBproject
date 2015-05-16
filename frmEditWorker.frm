VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditWorker 
   Caption         =   "עריכת עובד"
   ClientHeight    =   4770
   ClientLeft      =   5475
   ClientTop       =   2985
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdCalEndDate 
      Caption         =   "בחר"
      Height          =   255
      Left            =   1800
      TabIndex        =   37
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdCalStartDate 
      Caption         =   "בחר"
      Height          =   255
      Left            =   1800
      TabIndex        =   36
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdCalBirthday 
      Caption         =   "בחר"
      Height          =   255
      Left            =   4920
      TabIndex        =   35
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtFinishDate 
      DataField       =   "WDateFinish"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   2400
      TabIndex        =   33
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtStartDate 
      DataField       =   "WDateStart"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtId 
      DataField       =   "WID"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   5160
      TabIndex        =   19
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtFName 
      DataField       =   "WFirstName"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   5160
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtLName 
      DataField       =   "WLastName"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   5160
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtGender 
      DataField       =   "WGender"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "WPhone"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtBirthDate 
      DataField       =   "WBirthDate"
      DataSource      =   "adcWorkers"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtEMail 
      DataField       =   "WMail"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   5160
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtStreet 
      DataField       =   "Wstreet"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtCity 
      DataField       =   "WCity"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtNumer 
      DataField       =   "WNumber"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "שמור"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtCountry 
      DataField       =   "WCity"
      DataSource      =   "adcWorkers"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "עדכן"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdCanUpdate 
      Caption         =   "בטל עדכון"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox cmbGender 
      Height          =   315
      ItemData        =   "frmEditWorker.frx":0000
      Left            =   5160
      List            =   "frmEditWorker.frx":0002
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtGenderCode 
      DataField       =   "GCode"
      DataSource      =   "adcGender"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox cmbPhone 
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtPrefix 
      DataField       =   "PCode"
      DataSource      =   "adcPhone"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adcPhone 
      Height          =   330
      Left            =   4200
      Top             =   4320
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
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
   Begin MSAdodcLib.Adodc adcGender 
      Height          =   330
      Left            =   2400
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "Gender"
      Caption         =   "מגדר"
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
   Begin MSAdodcLib.Adodc adcWorkers 
      Height          =   330
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "עריכת עובד"
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
      Left            =   2880
      TabIndex        =   34
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblFinishDate 
      Caption         =   "תאריך סיום"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTextStartDate 
      Caption         =   "תאריך התחלה"
      Height          =   255
      Left            =   3720
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTextId 
      Caption         =   "תעודת זהות"
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblTextFName 
      Caption         =   "שם פרטי"
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblTextLName 
      Caption         =   "שם משפחה"
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTextGender 
      Caption         =   "מגדר"
      Height          =   255
      Left            =   6840
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTextPhone 
      Caption         =   "מספר טלפון"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTextEMail 
      Caption         =   "אי-מייל"
      Height          =   255
      Left            =   6840
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblTextBDate 
      Caption         =   "תאריך לידה"
      Height          =   255
      Left            =   6840
      TabIndex        =   23
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTextCity 
      Caption         =   "עיר"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblTextStreet 
      Caption         =   "רחוב"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblTextNumber 
      Caption         =   "מספר"
      Height          =   255
      Left            =   3720
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num As String
Dim flag As Boolean
Dim str As String
Dim index  As Integer
Dim strId1 As String
Dim strId2 As String

Private Sub cmbGender_click()
    adcWorkers.Recordset(3) = cmbGender.ListIndex + 1
End Sub

Private Sub cmdBack_Click()
    Load frmCustomers
    frmWorkers.Show
    Unload Me
End Sub

Private Sub cmdCalBirthday_Click()
    frmCalendar.Show
End Sub

Private Sub cmdCalEndDate_Click()
    frmCalendar.Show
End Sub

Private Sub cmdCalStartDate_Click()
    frmCalendar.Show
End Sub

Private Sub cmdCanUpdate_Click()
    adcWorkers.Recordset.CancelUpdate
End Sub

Private Sub cmdSave_Click()
    adcWorkers.Recordset(4).Value = cmbPhone.List(cmbPhone.ListIndex) & txtPhoneNum
    adcWorkers.Recordset.Update
    cmdBack.Enabled = True
    cmdUpdate.Enabled = True
    cmdSave.Enabled = False
    cmdCanUpdate.Enabled = False
    txtId.Enabled = False
    txtFName.Enabled = False
    txtLName.Enabled = False
    cmbGender.Enabled = False
    txtPhoneNum.Enabled = False
    cmbPhone.Enabled = False
    txtBirthDate.Enabled = False
    txtEMail.Enabled = False
    txtCity.Enabled = False
    txtStreet.Enabled = False
    txtNumer.Enabled = False
    txtStartDate.Enabled = False
    txtFinishDate.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
    cmdSave.Enabled = True
    cmdBack.Enabled = False
    cmdUpdate.Enabled = False
    cmdCanUpdate.Enabled = False
    txtId.Enabled = True
    txtFName.Enabled = True
    txtLName.Enabled = True
    cmbGender.Enabled = True
    txtPhoneNum.Enabled = True
    cmbPhone.Enabled = True
    txtBirthDate.Enabled = True
    txtEMail.Enabled = True
    txtCity.Enabled = True
    txtStreet.Enabled = True
    txtNumer.Enabled = True
    txtStartDate.Enabled = True
    txtFinishDate.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    flag = False
    adcGender.Recordset.MoveFirst
    While Not (adcGender.Recordset.EOF)
        cmbGender.AddItem (adcGender.Recordset(1))
        adcGender.Recordset.MoveNext
    Wend
    While Not (adcPhone.Recordset.EOF)
        cmbPhone.AddItem (adcPhone.Recordset(1))
        adcPhone.Recordset.MoveNext
    Wend
    adcWorkers.Recordset.MoveFirst
    num = frmWorkers.DGWorkers.Row
    For i = 1 To Val(num)
        adcWorkers.Recordset.MoveNext
    Next
    If frmWorkers.cmdUp = True Then
        cmbGender.ListIndex = adcWorkers.Recordset(3).Value - 1
        str = Left(adcWorkers.Recordset(4), 3)
        index = 0
        adcPhone.Recordset.MoveFirst
        While flag = False And adcPhone.Recordset.EOF = False
            If InStr(adcPhone.Recordset(1).Value, str) = 1 Then
                flag = True
                index = adcPhone.Recordset(0)
            Else
                adcPhone.Recordset.MoveNext
            End If
        Wend
        cmbPhone.ListIndex = index - 1
        txtPhoneNum = Right(adcWorkers.Recordset(4), 7)
    Else
        If frmWorkers.cmdAdd = True Then
            adcWorkers.Recordset.AddNew
        End If
    End If
    Unload frmWorkers
End Sub





Private Sub txtFName_lostFocus()
    Dim strName As String
    strName = txtFName.Text
    If frmMainMenu.isNameOk(strName) = False Then
        MsgBox "שם אינו יכול להכיל תווים שאינם אותיות"
        txtFName.SelStart = 0
        txtFName.SelLength = Len(strName)
        txtFName.SetFocus
    End If
End Sub

Private Sub txtId_Change()
    strId1 = txtId.Text
    If Len(strId1) <> 0 Then
        If IsNumeric(strId1) = False Then
            MsgBox "תעודת זהות לא יכולה להכיל תווים שאינם מספרים"
            txtId.SelStart = 0
            txtId.SelLength = Len(strId1)
            txtId.SetFocus
        End If
    End If
End Sub

Private Sub txtId_LostFocus()
    strId2 = txtId.Text
    If (frmMainMenu.isIdCorrect(strId2) = False) Then
        MsgBox "תעודת הזהות אינה תקינה"
        txtId.SelStart = 0
        txtId.SelLength = Len(strId2)
        txtId.SetFocus
    End If
End Sub

Private Sub txtLName_lostFocus()
    Dim strName As String
    strName = txtFName.Text
    If frmMainMenu.isNameOk(strName) = False Then
        MsgBox "שם אינו יכול להכיל תווים שאינם אותיות"
        txtFName.SelStart = 0
        txtFName.SelLength = Len(strName)
        txtFName.SetFocus
    End If
End Sub

Private Sub txtPhoneNum_lostFocus()
    Dim strPhone As String
    strPhone = txtPhoneNum.Text
    If frmMainMenu.isAllNumbers(strPhone) = False Then
        MsgBox "מספר טלפון לא יכול להכיל תווים שאינם מספרים"
        txtPhoneNum.SelStart = 0
        txtPhoneNum.SelLength = Len(strPhone)
        txtPhoneNum.SetFocus
    End If
End Sub

Private Sub txtEMail_lostFocus()
    Dim strEmail As String
    strEmail = txtEMail.Text
    If frmMainMenu.isEmailCorrect(strEmail) = False Then
        MsgBox "המייל שהוזן אינו תקין"
        txtEMail.SelStart = 0
        txtEMail.SelLength = Len(strEmail)
        txtEMail.SetFocus
    End If
End Sub

Private Sub txtCity_lostfocus()
    Dim strCity As String
    strCity = txtCity.Text
    If frmMainMenu.isNameOk(strCity) = False Then
        MsgBox "עיר אינה יכולה להכיל תווים שאינם אותיות"
        txtCity.SelStart = 0
        txtCity.SelLength = Len(strCity)
        txtCity.SetFocus
    End If
End Sub
