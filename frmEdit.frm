VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Edit Question From Test Bank"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton cmdMakeChanges 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   2160
      TabIndex        =   28
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox txtQuestion 
      Height          =   615
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtOptA 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question Type"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   5655
      Begin VB.OptionButton optMultiple 
         Caption         =   "Multiple Choice"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optTrueFalse 
         Caption         =   "True/False"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtOptB 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox txtOptC 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox txtOptD 
      Height          =   405
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Correct Answer Selection"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   5640
      Width           =   6015
      Begin VB.OptionButton optA 
         Caption         =   "Option A"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optB 
         Caption         =   "Option B"
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optC 
         Caption         =   "Option C"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optD 
         Caption         =   "Option D"
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Question:"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "250 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblLength 
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label10 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label11 
      Caption         =   "Option A:"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Option B:"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Option C:"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Option D:"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    'unload form
    Unload Me
    
End Sub

Private Sub cmdMakeChanges_Click()
    
    Dim i As Integer
    Dim x As Integer
    
    'check for blank fields
    If txtQuestion = "" Or txtOptA = "" Or txtOptB = "" Then
       MsgBox "One of your fields is blank.", , "Warning"
       Exit Sub
    End If
    'update question in bank database
    With frmInstructor.datBank.Recordset
        .Edit
        .Fields("question").Value = txtQuestion
        questHold(frmInstructor.lisTestBank.ListIndex + 1).quest = txtQuestion
        If optMultiple.Value = True Then
            .Fields("type").Value = "M"
            questHold(frmInstructor.lisTestBank.ListIndex + 1).theType = "M"
        Else
            .Fields("type").Value = "T"
            questHold(frmInstructor.lisTestBank.ListIndex + 1).theType = "T"
        End If
        .Fields("opt1").Value = txtOptA
        questHold(frmInstructor.lisTestBank.ListIndex + 1).answerA = txtOptA
        .Fields("opt2").Value = txtOptB
        questHold(frmInstructor.lisTestBank.ListIndex + 1).answerB = txtOptB
        If txtOptC = "" Then
            .Fields("opt3").Value = " "
            questHold(frmInstructor.lisTestBank.ListIndex + 1).answerC = " "
        Else
            .Fields("opt3").Value = txtOptC
            questHold(frmInstructor.lisTestBank.ListIndex + 1).answerC = txtOptC
        End If
        If txtOptD = "" Then
            .Fields("opt4").Value = " "
            questHold(frmInstructor.lisTestBank.ListIndex + 1).answerD = " "
        Else
            .Fields("opt4").Value = txtOptD
            questHold(frmInstructor.lisTestBank.ListIndex + 1).answerD = txtOptD
        End If
        If optA.Value = True Then
            .Fields("answer") = "A"
            questHold(frmInstructor.lisTestBank.ListIndex + 1).correctAns = "A"
        Else
            If optB.Value = True Then
                .Fields("answer") = "B"
                questHold(frmInstructor.lisTestBank.ListIndex + 1).correctAns = "B"
            Else
                If optC.Value = True Then
                    .Fields("answer") = "C"
                    questHold(frmInstructor.lisTestBank.ListIndex + 1).correctAns = "C"
                Else
                    .Fields("answer") = "D"
                    questHold(frmInstructor.lisTestBank.ListIndex + 1).correctAns = "D"
                End If
            End If
        End If
    
        .Update
    End With
    
    'save current number of questions in list box
    'clear list box
    x = frmInstructor.lisTestBank.ListCount
    frmInstructor.lisTestBank.Clear
    
    'load list box with newly edited question
    For i = 1 To x
        frmInstructor.lisTestBank.AddItem questHold(i).quest
    Next i
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    'make current form visible
    Show
    
    'center form on screen
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2

End Sub

Private Sub optMultiple_Click()

    'if multipe is selected clear first 2 answer option text
    'boxes in case they have been changed to true and false
    If optMultiple.Value = True Then
        txtOptA.Text = ""
        txtOptB.Text = ""
    End If
    
End Sub

Private Sub optTrueFalse_Click()
    
    'if true/false is selected set first 2 answer option text
    'boxes to true and false, and set focus on optionA
    If optTrueFalse.Value = True Then
        txtOptA.Text = "True"
        txtOptB.Text = "False"
        optA.SetFocus
    End If
        
End Sub

Private Sub txtOptA_Change()

    'display number of characters in text box
    Label4.Caption = Len(txtOptA.Text)
    
End Sub

Private Sub txtOptA_KeyPress(KeyAscii As Integer)
    
    'if return is pressed move to next field
    If KeyAscii = 13 Then
        txtOptB.SetFocus
        Exit Sub
    End If
    
    'if text is larger than 50 then do not allow additional
    'characters
    If Len(txtOptA.Text) > 49 Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtOptB_Change()
    
    'display number of characters in text box
    Label5.Caption = Len(txtOptB.Text)

End Sub

Private Sub txtOptB_KeyPress(KeyAscii As Integer)
    
    'if return is pressed move to next field
    If KeyAscii = 13 Then
        txtOptC.SetFocus
        Exit Sub
    End If
    
    'if text is larger than 50 then do not allow additional
    'characters
    If Len(txtOptB.Text) > 49 Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtOptC_Change()

    'display number of characters in text box
    Label6.Caption = Len(txtOptC.Text)

End Sub

Private Sub txtOptC_KeyPress(KeyAscii As Integer)
    
    'if return is pressed move to next field
    If KeyAscii = 13 Then
        txtOptD.SetFocus
        Exit Sub
    End If
    
    'if text is larger than 50 then do not allow additional
    'characters
    If Len(txtOptC.Text) > 49 Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtOptD_Change()
    
    'display number of characters in text box
    Label7.Caption = Len(txtOptD.Text)
    
End Sub

Private Sub txtOptD_KeyPress(KeyAscii As Integer)
    
    'if return is pressed move to next field
    If KeyAscii = 13 Then
        optA.SetFocus
        Exit Sub
    End If
    
    'if text is larger than 50 then do not allow additional
    'characters
    If Len(txtOptD.Text) > 49 Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub txtQuestion_Change()
        
    'display number of characters in text box
    lblLength.Caption = Len(txtQuestion.Text)

End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)

    'if return is pressed move to next field
    If KeyAscii = 13 Then
        optMultiple.SetFocus
        Exit Sub
    End If
    
    'if text is larger than 250 then do not allow additional
    'characters
    If Len(txtQuestion.Text) > 249 Then
        If Not KeyAscii = 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub

