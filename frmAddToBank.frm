VERSION 5.00
Begin VB.Form frmAddToBank 
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   495
      Left            =   5040
      TabIndex        =   26
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddAnother 
      Caption         =   "Add and Create Another"
      Height          =   495
      Left            =   3000
      TabIndex        =   25
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddClose 
      Caption         =   "Add and Return"
      Height          =   495
      Left            =   960
      TabIndex        =   24
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Correct Answer Selection"
      Height          =   855
      Left            =   720
      TabIndex        =   19
      Top             =   5640
      Width           =   6015
      Begin VB.OptionButton optD 
         Caption         =   "Option D"
         Height          =   255
         Left            =   4560
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optC 
         Caption         =   "Option C"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optB 
         Caption         =   "Option B"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optA 
         Caption         =   "Option A"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtOptD 
      Height          =   405
      Left            =   1800
      TabIndex        =   12
      Top             =   5040
      Width           =   4935
   End
   Begin VB.TextBox txtOptC 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox txtOptB 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Question Type"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
      Begin VB.OptionButton optTrueFalse 
         Caption         =   "True/False"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optMultiple 
         Caption         =   "Multiple Choice"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtOptA 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox txtQuestion 
      Height          =   615
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label14 
      Caption         =   "Option D:"
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Option C:"
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Option B:"
      Height          =   375
      Left            =   600
      TabIndex        =   28
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Option A:"
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "50 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label lblLength 
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "250 Characters Max                             Length:"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Question:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmAddToBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddAnother_Click()

    'check for blank fields
    If txtQuestion = "" Or txtOptA = "" Or txtOptB = "" Then
       MsgBox "One of your fields is blank.", , "Warning"
       Exit Sub
    End If
       
    'if newDB remove dummy record from database
    If newDB Then
        frmInstructor.datBank.Recordset.MoveFirst
        frmInstructor.datBank.Recordset.Delete
        newDB = False
    End If
        
    'call sub to add record
    Call addrecord
    
    'blank all fields
    txtQuestion.Text = ""
    txtOptA.Text = ""
    txtOptB.Text = ""
    txtOptC.Text = ""
    txtOptD.Text = ""
    optMultiple.Value = True
    optA.Value = True
    
End Sub

Private Sub cmdAddClose_Click()

    'check for blank fields
    If txtQuestion = "" Or txtOptA = "" Or txtOptB = "" Then
       MsgBox "One of your fields is blank.", , "Warning"
       Exit Sub
    End If
       
    'delete dummy record if database is new
    If newDB Then
        frmInstructor.datBank.Recordset.MoveFirst
        frmInstructor.datBank.Recordset.Delete
        newDB = False
    End If
        
    'call sub to add record
    Call addrecord
    
    'unload form
    Unload Me
        
End Sub

Private Sub addrecord()

    Dim i As Integer
    
    i = frmInstructor.lisTestBank.ListCount + 1
    frmInstructor.lisTestBank.AddItem (txtQuestion)
    
    'add new question to test bank database, and load question
    'hold array correct answer and question type
    With frmInstructor.datBank.Recordset
        .AddNew
        
        .Fields("question").Value = txtQuestion
        If optMultiple.Value = True Then
            .Fields("type").Value = "M"
            questHold(i).theType = "M"
        Else
            .Fields("type").Value = "T"
            questHold(i).theType = "T"
        End If
        .Fields("opt1").Value = txtOptA
        .Fields("opt2").Value = txtOptB
        If txtOptC = "" Then
            .Fields("opt3").Value = " "
        Else
            .Fields("opt3").Value = txtOptC
        End If
        If txtOptD = "" Then
            .Fields("opt4").Value = " "
        Else
            .Fields("opt4").Value = txtOptD
        End If
        If optA.Value = True Then
            .Fields("answer") = "A"
            questHold(i).correctAns = "A"
        Else
            If optB.Value = True Then
                .Fields("answer") = "B"
                questHold(i).correctAns = "B"
            Else
                If optC.Value = True Then
                    .Fields("answer") = "C"
                    questHold(i).correctAns = "C"
                Else
                    .Fields("answer") = "D"
                    questHold(i).correctAns = "D"
                End If
            End If
        End If
        
        .Update
        .MoveFirst
        
    End With
        
    'load question and answers to questHold array
    With questHold(i)
        .quest = txtQuestion
        .answerA = txtOptA
        .answerB = txtOptB
        .answerC = txtOptC
        .answerD = txtOptD
    End With
    
End Sub

Private Sub cmdCancel_Click()
    
    'unload form
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    'make current form visible
    Show
    
    'center form on screen
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    'set caption to include current bank
    Caption = "Add Question To Bank: " & currentBank
    
    'set type and answer selections
    optMultiple.Value = True
    optA.Value = True
    
End Sub

Private Sub optMultiple_Click()

    'if multiple is selected clear first 2 answer option text
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
