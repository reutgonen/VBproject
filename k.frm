VERSION 5.00
Begin VB.Form frmCalendar 
   Caption         =   "לוח שנה"
   ClientHeight    =   3915
   ClientLeft      =   6660
   ClientTop       =   2595
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   5610
   Begin VB.PictureBox CalendarBirthDay 
      Height          =   2655
      Left            =   480
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "אישור"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DateDay As String
Dim DateMonth As String
Dim DateYear As String
Dim birthCustomer As Boolean
Dim birthWorker As Boolean
Dim startWorker As Boolean
Dim endWorker As Boolean
Dim checkIn As Boolean
Dim checkOut As Boolean


Private Sub cmdOk_Click()
    If birthCustomer = True Then
        frmEditCustomer.txtBirthDate = CalendarBirthDay.Value
        frmEditCustomer.txtBirthdayShow = CalendarBirthDay.Value
        Unload Me
    Else
        If birthWorker = True Then
            frmEditWorker.txtBirthDate = CalendarBirthDay.Value
            Unload Me
        Else
            If startWorker = True Then
                frmEditWorker.txtStartDate = CalendarBirthDay.Value
                Unload Me
            Else
                If endWorker = True Then
                    frmEditWorker.txtFinishDate = CalendarBirthDay.Value
                    Unload Me
                Else
                    If checkIn = True Then
                        frmEditBooking.txtCheckIn = CalendarBirthDay.Value
                        Unload Me
                    Else
                        If checkOut = True Then
                            frmEditBooking.txtCheckOut = CalendarBirthDay.Value
                            Unload Me
                        End If
                    End If
                End If
            End If
        End If
    End If
                
End Sub

Private Sub Form_Load()
    birthCustomer = False
    birthWorker = False
    startWorker = False
    endWorker = False
    checkIn = False
    checkOut = False
    Dim previousDate As String
    If frmEditCustomer.cmdDate = True Then
        birthCustomer = True
        If Len(frmEditCustomer.adcCustomer.Recordset(5)) <> 0 Then
            previousDate = frmEditCustomer.adcCustomer.Recordset(5)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
    If frmEditWorker.cmdCalBirthday = True Then
        birthWorker = True
        If Len(frmEditWorker.adcWorkers.Recordset(5)) <> 0 Then
            previousDate = frmEditWorker.adcWorkers.Recordset(5)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
    If frmEditWorker.cmdCalStartDate = True Then
        startWorker = True
        If Len(frmEditWorker.adcWorkers.Recordset(10)) <> 0 Then
            previousDate = frmEditWorker.adcWorkers.Recordset(10)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
    If frmEditWorker.cmdCalEndDate = True Then
        endWorker = True
        If Len(frmEditWorker.adcWorkers.Recordset(11)) <> 0 Then
            previousDate = frmEditWorker.adcWorkers.Recordset(11)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
    If frmEditBooking.cmdCheckIn = True Then
        checkIn = True
        If Len(frmEditBooking.adcBooking.Recordset(3)) <> 0 Then
            previousDate = frmEditBooking.adcBooking.Recordset(3)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
    If frmEditBooking.cmdCheckOut = True Then
        checkOut = True
        If Len(frmEditBooking.adcBooking.Recordset(4)) <> 0 Then
            previousDate = frmEditBooking.adcBooking.Recordset(4)
            DateDay = Left(previousDate, 2)
            DateMonth = Mid(previousDate, 4, 2)
            DateYear = Right(previousDate, 4)
            CalendarBirthDay.Year = DateYear
            CalendarBirthDay.Month = DateMonth
            CalendarBirthDay.Day = DateDay
        End If
    End If
End Sub
