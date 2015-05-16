VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditCustomer 
   Caption         =   "עריכת לקוח"
   ClientHeight    =   5265
   ClientLeft      =   4875
   ClientTop       =   2790
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9390
   Begin VB.CommandButton cmdSave 
      Caption         =   "שמור"
      Height          =   495
      Left            =   2760
      TabIndex        =   35
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "חזור"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   34
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "עדכן"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   33
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdCanUpdate 
      Caption         =   "בטל עדכון"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   32
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtBirthdayShow 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      TabIndex        =   31
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtBirthDate 
      DataField       =   "CBirthDate"
      DataSource      =   "adcCustomer"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "בחר"
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtPrefix 
      DataField       =   "PCode"
      DataSource      =   "adcPhone"
      Height          =   285
      Left            =   120
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adcPhone 
      Height          =   330
      Left            =   6120
      Top             =   4440
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
   Begin VB.ComboBox cmbPhone 
      Height          =   315
      Left            =   4920
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   285
      Left            =   5640
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtGenderCode 
      DataField       =   "GCode"
      DataSource      =   "adcGender"
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adcGender 
      Height          =   330
      Left            =   4320
      Top             =   4440
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
   Begin VB.ComboBox cmbGender 
      Height          =   315
      ItemData        =   "frmEditCustomer.frx":0000
      Left            =   5160
      List            =   "frmEditCustomer.frx":0002
      TabIndex        =   23
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtCountry 
      DataField       =   "CCountry"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adcCountry 
      Height          =   330
      Left            =   2520
      Top             =   4440
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
      RecordSource    =   "Country"
      Caption         =   "מדינות"
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
   Begin VB.ComboBox cmbCountry 
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   1320
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc adcCustomer 
      Height          =   330
      Left            =   360
      Top             =   4440
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
   Begin VB.TextBox txtNumer 
      DataField       =   "CNum"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtCity 
      DataField       =   "CCity"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtStreet 
      DataField       =   "CStreet"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtCountry11 
      DataField       =   "CNName"
      DataSource      =   "adcCountry"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtEMail 
      DataField       =   "CMail"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "CPhoneNum"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtGender 
      DataField       =   "CGender"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLName 
      DataField       =   "CLastName"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtFName 
      DataField       =   "CFirstName"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtId 
      DataField       =   "CID"
      DataSource      =   "adcCustomer"
      Height          =   285
      Left            =   5160
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblTitel 
      Caption         =   "עריכת לקוח"
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
      TabIndex        =   28
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblTextNumber 
      Caption         =   "מספר"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTextStreet 
      Caption         =   "רחוב"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTextCity 
      Caption         =   "עיר"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTextCountry 
      Caption         =   "מדינה"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblTextBDate 
      Caption         =   "תאריך לידה"
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTextEMail 
      Caption         =   "אי-מייל"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblTextPhone 
      Caption         =   "מספר טלפון"
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTextGender 
      Caption         =   "מגדר"
      Height          =   255
      Left            =   6840
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblTextLName 
      Caption         =   "שם משפחה"
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTextFName 
      Caption         =   "שם פרטי"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblTextId 
      Caption         =   "תעודת זהות"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditCustomer"
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
Dim idEdit As String

Private Sub cmbCountry_Click()
    adcCustomer.Recordset(7) = cmbCountry.ListIndex + 1
End Sub

Private Sub cmbGender_click()
    adcCustomer.Recordset(3) = cmbGender.ListIndex + 1
End Sub

Private Sub cmdBack_Click()
    Load frmCustomers
    frmCustomers.Show
    frmCustomers.Refresh
    Unload Me
End Sub

Private Sub cmdCanUpdate_Click()
    adcCustomer.Recordset.CancelUpdate
End Sub

Private Sub cmdDate_Click()
    frmCalendar.Show
End Sub

Private Sub cmdSave_Click()
    Dim id As String
    Dim flag2 As String
    adcCustomer.Recordset(4).Value = cmbPhone.List(cmbPhone.ListIndex) & txtPhoneNum
    id = txtId.Text
    flag2 = True
    Load frmCustomers
    frmCustomers.adcCustomers.Recordset.MoveFirst
    While (frmCustomers.adcCustomers.Recordset.EOF = False And flag2 = True)
        If frmCustomers.adcCustomers.Recordset(0) <> idEdit Then
            If InStr(frmCustomers.adcCustomers.Recordset(0), id) = 1 Then
                flag2 = False
            End If
        End If
        frmCustomers.adcCustomers.Recordset.MoveNext
    Wend
    Unload frmCustomers
    If flag2 = True Then
        adcCustomer.Recordset.Update
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
        cmbCountry.Enabled = False
        txtCity.Enabled = False
        txtStreet.Enabled = False
        txtNumer.Enabled = False
    Else
        MsgBox "תעודת הזהות כבר קיימת במערכת"
    End If
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
    cmbCountry.Enabled = True
    txtCity.Enabled = True
    txtStreet.Enabled = True
    txtNumer.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strDate As String
    Dim DateBirth As String
    flag = False
    adcCountry.Recordset.MoveFirst
    While Not (adcCountry.Recordset.EOF)
        cmbCountry.AddItem (adcCountry.Recordset(1))
        adcCountry.Recordset.MoveNext
    Wend
    adcGender.Recordset.MoveFirst
    While Not (adcGender.Recordset.EOF)
        cmbGender.AddItem (adcGender.Recordset(1))
        adcGender.Recordset.MoveNext
    Wend
    While Not (adcPhone.Recordset.EOF)
        cmbPhone.AddItem (adcPhone.Recordset(1))
        adcPhone.Recordset.MoveNext
    Wend
    adcCustomer.Recordset.MoveFirst
    num = frmCustomers.DGCustomers.Row
    For i = 1 To Val(num)
        adcCustomer.Recordset.MoveNext
    Next
    If frmCustomers.cmdUp = True Then
        cmbCountry.ListIndex = adcCustomer.Recordset(7).Value - 1
        cmbGender.ListIndex = adcCustomer.Recordset(3).Value - 1
        txtBirthdayShow.Text = adcCustomer.Recordset(5)
        str = Left(adcCustomer.Recordset(4), 3)
        idEdit = adcCustomer.Recordset(0)
        DateBirth = adcCustomer.Recordset(5)
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
        txtPhoneNum = Right(adcCustomer.Recordset(4), 7)
    Else
        If frmCustomers.cmdAdd = True Then
            adcCustomer.Recordset.AddNew
        End If
        idEdit = ""
    End If
    
    Unload frmCustomers
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
