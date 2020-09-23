VERSION 5.00
Begin VB.Form frmMath 
   Caption         =   "Math Practice"
   ClientHeight    =   5100
   ClientLeft      =   1050
   ClientTop       =   1470
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5100
   ScaleWidth      =   8190
   Begin VB.CheckBox chkDisplay 
      Caption         =   "&Display summary infomation"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Operations"
      Height          =   1695
      Left            =   2700
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
      Begin VB.OptionButton optDivision 
         Caption         =   "D&ivision"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1300
         Width           =   900
      End
      Begin VB.OptionButton optMultipy 
         Caption         =   "&Multilication "
         Height          =   495
         Left            =   150
         TabIndex        =   18
         Top             =   840
         Width           =   1275
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "&Addition"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
      Begin VB.OptionButton optSub 
         Caption         =   "&Subtraction"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Levels"
      Height          =   1575
      Left            =   600
      TabIndex        =   16
      Top             =   2400
      Width           =   1935
      Begin VB.OptionButton optLevel1 
         Caption         =   "Level &1   (1-10)"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optLevel2 
         Caption         =   "Level &2   (10-100)"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1215
      Left            =   5400
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
      Begin VB.Label lblIncorrect 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblCorrect 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Correct:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         Caption         =   "Incorrect:"
         Height          =   240
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   7455
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify Answer"
         Default         =   -1  'True
         Height          =   495
         Left            =   5640
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgIcon 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4800
         Top             =   840
         Width           =   615
      End
      Begin VB.Image imgOperator 
         Height          =   495
         Left            =   1560
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblNum1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblNum2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Image imgEqual 
         Height          =   480
         Left            =   3120
         Picture         =   "math.frx":0000
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image imgDivision 
      Height          =   480
      Left            =   2800
      Picture         =   "math.frx":0442
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMulpulication 
      Height          =   480
      Left            =   2160
      Picture         =   "math.frx":0884
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMinus 
      Height          =   480
      Left            =   1560
      Picture         =   "math.frx":0CC6
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlus 
      Height          =   480
      Left            =   960
      Picture         =   "math.frx":1108
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgHappy 
      Height          =   480
      Left            =   360
      Picture         =   "math.frx":154A
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim num1 As Integer
Dim num2 As Integer

Private Sub chkDisplay_Click()
    If chkDisplay.Value = vbChecked Then
        fraInfo.Visible = True
    Else
        fraInfo.Visible = False
    End If
    
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdVerify_Click()
    Const box As String = vbOKOnly + vbInformation + vbDefaultButton1 + vbApplicationModal
    Dim userAnswer As Integer, rightAnswer As Integer, numb As Integer
    Static numc As Integer, numw As Integer
    userAnswer = Val(txtAnswer.Text)
    If optAdd.Value = True Then
        rightAnswer = num1 + num2
    Else
        rightAnswer = num1 - num2
    End If
    If optMultipy.Value = True Then
        rightAnswer = num1 * num2
    ElseIf optDivision.Value = True Then
        rightAnswer = num1 / num2
    End If
    Select Case userAnswer
        Case Is = rightAnswer
            imgIcon.Picture = imgHappy.Picture
            numc = numc + 1
            txtAnswer.Text = ""
            Call RandomNumbers
    Case Else  'wrong answer
        imgIcon = LoadPicture()
        numw = numw + 1
        numb = MsgBox("try again", box, "Math Application") 'message box prompt when wrong answer
        'highlight text
        txtAnswer.SelStart = 0
        txtAnswer.SelLength = Len(txtAnswer.Text)
    End Select
    txtAnswer.SetFocus
    lblCorrect.Caption = numc
    lblIncorrect.Caption = numw
End Sub

Private Sub cmdVerify_GotFocus()

        txtAnswer.SelStart = 0
        txtAnswer.SelLength = Len(txtAnswer.Text)
End Sub
Private Sub Form_Load()
    fraInfo.Visible = False
    optLevel1.Value = True
    optAdd.Value = True
    frmMath.Top = (Screen.Height - frmMath.Height) / 2
    frmMath.Left = (Screen.Width - frmMath.Width) / 2

End Sub

Private Sub RandomNumbers()
    Dim temp As Integer
    
    Randomize
    If optLevel1.Value = True Then  'generate random numbers
    ' Level 1 option button is selected
       num1 = Int((10 - 1 + 1) * Rnd + 1)
       num2 = Int((10 - 1 + 1) * Rnd + 1)
    Else        ' Level 2 option button is  selected
        num1 = Int((100 - 10 + 1) * Rnd + 1)
        num2 = Int((100 - 10 + 1) * Rnd + 1)
    End If
    If optSub.Value = True And num2 > num1 Then 'swap numbers
        temp = num1
        num2 = num1
        num2 = temp
    End If
    If optDivision.Value = True And num2 > num1 Then
        temp = num1
        num2 = num1
        num2 = temp
    End If
    lblNum1.Caption = num1
    lblNum2.Caption = num2
End Sub
 
Private Sub optAdd_Click()
    Call RandomNumbers
   imgOperator = imgPlus.Picture
End Sub

Private Sub optDivision_Click()
    Call RandomNumbers
    imgOperator = imgDivision.Picture
    txtAnswer.SetFocus
End Sub

Private Sub optLevel1_Click()
    Call RandomNumbers

End Sub

Private Sub optLevel2_Click()
    Call RandomNumbers

End Sub

Private Sub optMultipy_Click()
    Call RandomNumbers
    imgOperator = imgMulpulication.Picture

End Sub

Private Sub optSub_Click()
    Call RandomNumbers
    imgOperator = imgMinus.Picture

End Sub
