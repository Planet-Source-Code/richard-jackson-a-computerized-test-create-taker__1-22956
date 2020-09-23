VERSION 5.00
Begin VB.Form frmDisplayGrade 
   Caption         =   "Results"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintScore 
      Caption         =   "Print Score and Answers"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picAnswers 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display Score and Answers"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox picScore 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmDisplayGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDisplay_Click()

    Dim i As Integer
    Dim x As Integer
    Dim numOfQuest As Integer
    
    'clear pic box and display score
    picScore.Cls
    picScore.Print "Out of " & numOfQ & " questions"
    picScore.Print "you answered " & numCorrect & " correctly."
    picScore.Print "Your score is " & userScore & "."
    
    
    numOfQuest = numOfQ
    x = 0
    For i = 1 To numOfQuest
        x = x + 1
        'display question numbers and X's if answer is incorrect
        If usersAnswer(i) = RTrim(questTest(i).correctAns) Then
            picAnswers.Print i & " - " & usersAnswer(i),
        Else
            picAnswers.Print "X" & i & " - " & usersAnswer(i) & "X",
        End If
        'allow 4 answers to be displayed per line
        If x = 4 Then
            picAnswers.Print
            x = 0
        End If
     Next i
End Sub

Private Sub cmdPrintScore_Click()

    Dim i As Integer
    Dim x As Integer
    Dim numOfQuest As Integer
    
    'print score on printer
    Printer.FontSize = 12
    Printer.Print "Out of " & numOfQ & " questions"
    Printer.Print "you answered " & numCorrect & " correctly."
    Printer.Print "Your score is " & userScore & "."
    
    Printer.Print: Printer.Print
    
    numOfQuest = numOfQ
    x = 0
    For i = 1 To numOfQuest
        x = x + 1
        'print question number and X's on wrong answer to printer
        If usersAnswer(i) = RTrim(questTest(i).correctAns) Then
            Printer.Print i & " - " & usersAnswer(i),
        Else
            Printer.Print "X" & i & " - " & usersAnswer(i) & "X",
        End If
        'allow 4 answers on each line
        If x = 4 Then
            Printer.Print
            x = 0
        End If
     Next i
     
     Printer.EndDoc
     
End Sub

Private Sub Form_Load()

    Show
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
End Sub
