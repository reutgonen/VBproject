VERSION 5.00
Begin VB.Form frmMainMenu 
   Caption         =   "בית הלון ""גוננוס"" - תפריט ראשי"
   ClientHeight    =   5025
   ClientLeft      =   4680
   ClientTop       =   2790
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8100
   Begin VB.CommandButton Command2 
      Caption         =   "check"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtCheckId 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdministration 
      Caption         =   "מנהלה"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdWorker 
      Caption         =   "עובדים"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "הפקת דו""חות"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdMeal 
      Caption         =   "ארוחות"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCustomer 
      Caption         =   "טיפול באורחים"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBooking 
      Caption         =   "ביצוע הזמנה"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblTitel 
      Caption         =   "בית מלון ""גוננוס"" - תפריט ראשי"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCustomer_Click()
    frmCustomers.Show
End Sub

Private Sub cmdAdministration_Click()
    frmAdministration.Show
End Sub

Public Function isIdCorrect(str As String)
    Dim digBekuret As Integer
    Dim dig As Integer
    Dim sum As Integer
    Dim num As Integer
    Dim i As Integer
    sum = 0
    If Len(str) <> 9 Then
        isIdCorrect = False
    Else
        digBekuret = Val(Right(str, 1))
        str = Left(str, 8)
        For i = 0 To 7
            dig = Val(Right(str, 1))
            str = Left(str, Len(str) - 1)
            If i Mod 2 = 0 Then
                If dig >= 5 Then
                    sum = sum + ((dig * 2) Mod 10) + Int(((dig * 2) / 10))
                Else
                    sum = sum + dig * 2
                End If
            Else
                sum = sum + dig
            End If
        Next
        If sum Mod 10 = 0 Then
            num = 10
        Else
            num = sum Mod 10
        End If
        If digBekuret = 10 - num Then
            isIdCorrect = True
        Else
            isIdCorrect = False
        End If
    End If
End Function

Public Function isNameOk(str As String)
    Dim tav As String
    Dim flag As Boolean
    flag = True
    While Len(str) <> 0
        tav = Right(str, 1)
        If Not ((tav >= "a" And tav <= "z") Or (tav >= "A" And tav <= "Z") Or (tav >= "א" And tav <= "ת") Or (tav = " ") Or (tav = "-")) Then
            flag = False
        End If
        str = Left(str, Len(str) - 1)
    Wend
    isNameOk = flag
End Function

Public Function isAllNumbers(str As String)
    Dim tav As String
    Dim flag As Boolean
    flag = True
    While Len(str) <> 0
        tav = Right(str, 1)
        If IsNumeric(tav) = False Then
            flag = False
        End If
        str = Left(str, Len(str) - 1)
    Wend
    isAllNumbers = flag
End Function

Public Function isEmailCorrect(str As String)
    Dim tav As String
    Dim flag As Boolean
    Dim counter As Integer
    Dim str1 As String
    Dim str2 As String
    flag = True
    str1 = str
    While Len(str1) <> 0
        tav = Right(str1, 1)
        If ((IsNumeric(tav) = True) Or Not (tav > "a" And tav < "z") Or Not (tav > "A" And tav < "Z") Or tav = "@" Or tav = "_" Or tav = "-" Or tav = ".") = False Then
            flag = False
        End If
        str1 = Left(str1, Len(str1) - 1)
    Wend
    counter = 0
    str2 = str
    While Len(str2) <> 0
        tav = Right(str2, 1)
        If tav = "@" Then
            counter = counter + 1
        End If
        str2 = Left(str2, Len(str2) - 1)
    Wend
    If counter <> 1 Then
        flag = False
    End If
    tav = Right(str, 1)
    If tav = "." Or tav = "-" Or tav = "_" Or tav = "@" Then
        flag = False
    End If
    tav = Left(str, 1)
    If tav = "." Or tav = "-" Or tav = "_" Or tav = "@" Then
        flag = False
    End If
    isEmailCorrect = flag
End Function

Private Sub Command2_Click()
    Command2.BackColor = vbBlue
End Sub
