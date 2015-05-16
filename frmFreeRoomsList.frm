VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFreeRoomsList 
   Caption         =   "חדרים פנויים - רשימה"
   ClientHeight    =   4905
   ClientLeft      =   4215
   ClientTop       =   2265
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   12375
   Begin VB.CommandButton cmdCancel 
      Caption         =   "חזור"
      Height          =   375
      Left            =   720
      TabIndex        =   26
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "אישור"
      Height          =   375
      Left            =   720
      TabIndex        =   25
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton cmdRoom 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   3240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox txtCheckOutDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtCheckInDate 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtBookNum 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adcBookingRooms 
      Height          =   330
      Left            =   3600
      Top             =   4440
      Visible         =   0   'False
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
      Connect         =   $"frmFreeRoomsList.frx":0000
      OLEDBString     =   $"frmFreeRoomsList.frx":0098
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RoomsBooking"
      Caption         =   "הזמנות - חדרים"
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
   Begin MSDataGridLib.DataGrid DGRoomDates 
      Bindings        =   "frmFreeRoomsList.frx":0130
      Height          =   3015
      Left            =   5400
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
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
   Begin MSAdodcLib.Adodc adcRoomsDates 
      Height          =   330
      Left            =   360
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      Caption         =   "חדרים - תאריכים"
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
   Begin VB.ComboBox cmbYear 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7800
      TabIndex        =   3
      Text            =   "שנה"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbMonth 
      Enabled         =   0   'False
      Height          =   315
      Left            =   9240
      TabIndex        =   2
      Text            =   "חודש"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cmbRoom 
      Height          =   315
      Left            =   10560
      TabIndex        =   1
      Text            =   "חדר"
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox lstDates 
      Height          =   2985
      Left            =   9840
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblTextChoose 
      Caption         =   "- הבחירה שלך"
      Height          =   255
      Left            =   4080
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblTextClose 
      Caption         =   "- תפוס"
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblTextFree 
      Caption         =   "- פנוי"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblRed 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblYellow 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblTextFreeRooms 
      Caption         =   "חדרים פנויים"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblTextOutDate 
      Caption         =   "תאריך עזיבה"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblTextCheckInDate 
      Caption         =   "תאריך הגעה"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblTextBookNum 
      Caption         =   "מספר הזמנה"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblTextMyBooking 
      Caption         =   "ההזמנה שלי"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "frmFreeRoomsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag28 As Boolean
Dim year As Integer
Dim month As Integer
Dim arr(1 To 6, 1 To 31) As Boolean
Dim flagLoad As Boolean
Dim sumNight As Integer

Private Sub cmbMonth_click()
If flagLoad = False Then
    lstDates.Clear
    month = cmbMonth.Text
    cmbYear.Enabled = True
    If year <> 0 Then
        If year Mod 4 <> 0 Then
            flag28 = True
        Else
            If year / 400 = 0 Then
                flag28 = False
            Else
                If year / 100 = 0 Then
                    flag28 = True
                Else
                    flag28 = False
                End If
            End If
        End If
        Dim day As Integer
        Select Case month
            Case 1, 3, 5, 7, 8, 10, 12
                day = 31
            Case 4, 6, 9, 11
                day = 30
               Case 2
                If flag28 = True Then
                    day = 28
                Else
                    day = 29
                End If
            End Select
        Dim i As Integer
        Dim bookMonth As Integer
        Dim bookDay As Integer
        Dim outMonth As Integer
        Dim outDay As Integer
        Dim flag As Boolean
        Dim j As Integer
        For i = 1 To day
            flag = True
            adcRoomsDates.Recordset.MoveFirst
            While adcRoomsDates.Recordset.EOF = False
                bookMonth = Mid(adcRoomsDates.Recordset(1), 4, 2)
                outMonth = Mid(adcRoomsDates.Recordset(2), 4, 2)
                    If bookMonth = month Then
                        bookDay = Left(adcRoomsDates.Recordset(1), 2)
                        If Mid(adcRoomsDates.Recordset(1), 4, 2) = Mid(adcRoomsDates.Recordset(2), 4, 2) Then
                            For j = bookDay To Left(adcRoomsDates.Recordset(2), 2) - 1
                                If i = j Then
                                    flag = False
                                End If
                            Next
                        Else
                            For j = bookDay To day
                                If i = j Then
                                    flag = False
                                End If
                            Next
                        End If
                    Else
                        If outMonth = month Then
                            outDay = Left(adcRoomsDates.Recordset(2), 2)
                            For j = 1 To outDay
                                If j = i Then
                                    flag = False
                                End If
                            Next
                        End If
                    End If
                adcRoomsDates.Recordset.MoveNext
            Wend
            If flag = True Then
                lstDates.AddItem (i)
            End If
        Next
    lstDates.Refresh
    End If
End If
End Sub

Private Sub cmbRoom_Click()
    cmbMonth.Enabled = True
    Dim roomNum As Integer
    roomNum = cmbRoom.Text
    adcRoomsDates.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
    adcRoomsDates.CommandType = adCmdText
    adcRoomsDates.RecordSource = "Select RBRoomNumber,BDateCheckIn,BDateCheckOut From Booking b, RoomsBooking rb Where b.BCode = rb.RBBookingCode and rb.RBRoomNumber = " & roomNum & " "
    adcRoomsDates.Refresh
End Sub

Private Sub cmbYear_Click()
If flagLoad = False Then
    lstDates.Clear
    year = cmbYear.Text
    If year Mod 4 <> 0 Then
        flag28 = True
    Else
        If year / 400 = 0 Then
            flag28 = False
        Else
            If year / 100 = 0 Then
                flag28 = True
            Else
                flag28 = False
            End If
        End If
    End If
    Dim day As Integer
    Select Case month
        Case 1, 3, 5, 7, 8, 10, 12
            day = 31
        Case 4, 6, 9, 11
            day = 30
        Case 2
            If flag28 = True Then
                day = 28
            Else
                day = 29
            End If
    End Select
    Dim i As Integer
    Dim bookMonth As Integer
    Dim bookDay As Integer
    Dim outMonth As Integer
    Dim outDay As Integer
    Dim flag As Boolean
    Dim j As Integer
    For i = 1 To day
        flag = True
        adcRoomsDates.Recordset.MoveFirst
        While adcRoomsDates.Recordset.EOF = False
            bookMonth = Mid(adcRoomsDates.Recordset(1), 4, 2)
            outMonth = Mid(adcRoomsDates.Recordset(2), 4, 2)
                If bookMonth = month Then
                    bookDay = Left(adcRoomsDates.Recordset(1), 2)
                    If Mid(adcRoomsDates.Recordset(1), 4, 2) = Mid(adcRoomsDates.Recordset(2), 4, 2) Then
                        For j = bookDay To Left(adcRoomsDates.Recordset(2), 2) - 1
                            If i = j Then
                                flag = False
                            End If
                        Next
                    Else
                        For j = bookDay To day
                            If i = j Then
                                flag = False
                            End If
                        Next
                    End If
                Else
                    If outMonth = month Then
                        outDay = Left(adcRoomsDates.Recordset(2), 2)
                        For j = 1 To outDay
                            If j = i Then
                                flag = False
                            End If
                        Next
                    End If
                End If
            adcRoomsDates.Recordset.MoveNext
        Wend
        If flag = True Then
            lstDates.AddItem (i)
        End If
    Next
    lstDates.Refresh
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim index As Integer
    For index = 0 To 5
        If cmdRoom(index).BackColor = vbYellow Then
            frmEditBooking.cmdRoom(index).Visible = True
        End If
    Next
    frmEditBooking.txtSumNights.Text = sumNight
    Unload Me
End Sub

Private Sub cmdRoom_Click(index As Integer)
    If cmdRoom(index).BackColor <> vbYellow Then
        cmdRoom(index).BackColor = vbYellow
    Else
        cmdRoom(index).BackColor = vbGreen
    End If
End Sub

Private Sub Form_Load()
    flagLoad = True
    txtBookNum.Text = frmEditBooking.txtCode.Text
    txtCheckInDate.Text = frmEditBooking.txtCheckIn.Text
    txtCheckOutDate.Text = frmEditBooking.txtCheckOut.Text
    Dim index As Integer
    Dim index2 As Integer
    For index = 1 To 6
        For index2 = 1 To 31
            arr(index, index2) = False
        Next
    Next
    For index = 1 To 6
        cmbRoom.AddItem (index)
    Next
    For index = 1 To 12
        cmbMonth.AddItem (index)
    Next
    For index = 2015 To 2030
        cmbYear.AddItem (index)
    Next
    year = Right(txtCheckInDate.Text, 4)
    cmbYear.ListIndex = year - 2015
    month = Mid(txtCheckInDate.Text, 4, 2)
    cmbMonth.ListIndex = month - 1
    Dim room As Integer
    If year <> 0 Then
        If year Mod 4 <> 0 Then
            flag28 = True
        Else
            If year / 400 = 0 Then
                flag28 = False
            Else
                If year / 100 = 0 Then
                    flag28 = True
                Else
                    flag28 = False
                End If
            End If
        End If
        Dim day As Integer
        Select Case month
            Case 1, 3, 5, 7, 8, 10, 12
                day = 31
            Case 4, 6, 9, 11
                day = 30
            Case 2
                If flag28 = True Then
                    day = 28
                Else
                    day = 29
                End If
        End Select
        Dim i As Integer
        Dim bookMonth As Integer
        Dim bookDay As Integer
        Dim outMonth As Integer
        Dim outDay As Integer
        Dim flag As Boolean
        Dim j As Integer
        adcRoomsDates.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBhotel.mdb;Persist Security Info=False"
        adcRoomsDates.CommandType = adCmdText
        adcRoomsDates.RecordSource = "Select RBRoomNumber,BDateCheckIn,BDateCheckOut From Booking b, RoomsBooking rb Where b.BCode = rb.RBBookingCode "
        adcRoomsDates.Refresh
        For room = 1 To 6
            adcRoomsDates.Recordset.Filter = "RBRoomNumber = '" & room & "' "
            For i = 1 To day
                flag = True
                adcRoomsDates.Recordset.MoveFirst
                While adcRoomsDates.Recordset.EOF = False
                    bookMonth = Mid(adcRoomsDates.Recordset(1), 4, 2)
                    outMonth = Mid(adcRoomsDates.Recordset(2), 4, 2)
                        If bookMonth = month Then
                            bookDay = Left(adcRoomsDates.Recordset(1), 2)
                            If Mid(adcRoomsDates.Recordset(1), 4, 2) = Mid(adcRoomsDates.Recordset(2), 4, 2) Then
                                For j = bookDay To Left(adcRoomsDates.Recordset(2), 2) - 1
                                    If i = j Then
                                        flag = False
                                    End If
                                Next
                            Else
                                For j = bookDay To day
                                    If i = j Then
                                        flag = False
                                    End If
                                Next
                            End If
                        Else
                           If outMonth = month Then
                                outDay = Left(adcRoomsDates.Recordset(2), 2)
                                For j = 1 To outDay
                                    If j = i Then
                                        flag = False
                                    End If
                                Next
                            End If
                        End If
                    adcRoomsDates.Recordset.MoveNext
                Wend
                If flag = True Then
                    arr(room, i) = True
                End If
            Next
            adcRoomsDates.Refresh
        Next
        Dim flagFree As Boolean
        bookDay = Left(txtCheckInDate, 2)
        outDay = Left(txtCheckOutDate, 2)
        If outDay > bookDay Then
            For index = 1 To 6
                flagFree = True
                For index2 = bookDay To outDay
                    If arr(index, index2) = False Then
                        flagFree = False
                    End If
                Next
                If flagFree = True Then
                    cmdRoom(index - 1).BackColor = vbGreen
                Else
                    cmdRoom(index - 1).BackColor = vbRed
                    cmdRoom(index - 1).Enabled = False
                End If
            Next
            sumNight = outDay - bookDay
        Else
            For index = 1 To 6
                flagFree = True
                For index2 = bookDay To day
                    If arr(index, index2) = False Then
                        flagFree = False
                    End If
                Next
                If flagFree = True Then
                    cmdRoom(index - 1).BackColor = vbGreen
                Else
                    cmdRoom(index - 1).BackColor = vbRed
                    cmdRoom(index - 1).Enabled = False
                End If
            Next
            sumNight = day - bookDay + outDay
        End If
   End If
   flagLoad = False
End Sub
