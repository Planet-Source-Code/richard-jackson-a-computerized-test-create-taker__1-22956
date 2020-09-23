VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInstructor 
   Caption         =   "Computerized Testing Program - Instructor Screen"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data datBank 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdSaveTest 
      Caption         =   "Save Test"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10680
      TabIndex        =   39
      Top             =   7320
      Width           =   975
   End
   Begin VB.CheckBox chkAllowGoBack 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Allow Changes To Answers That Have Already Been Answered"
      Height          =   375
      Left            =   8040
      TabIndex        =   38
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Test Background Color"
      Height          =   975
      Left            =   8040
      TabIndex        =   33
      Top             =   6960
      Width           =   2415
      Begin VB.OptionButton optGreen 
         BackColor       =   &H0000FF00&
         Caption         =   "Green"
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optBlue 
         BackColor       =   &H00FF0000&
         Caption         =   "Blue"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optRed 
         BackColor       =   &H000000FF&
         Caption         =   "Red"
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optGrey 
         Caption         =   "Grey"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtMinutes 
      Height          =   285
      Left            =   9360
      TabIndex        =   31
      Top             =   6090
      Width           =   615
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "q"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "p"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdReviewQuestion 
      Caption         =   "View Question In Template"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   28
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveTestQuest 
      Caption         =   "Remove Test Question"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   27
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox chkTimed 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Timed Test"
      Height          =   255
      Left            =   8040
      TabIndex        =   26
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ListBox lisTestQuests 
      Height          =   2205
      Left            =   600
      TabIndex        =   25
      Top             =   5760
      Width           =   5895
   End
   Begin VB.TextBox txtOption4 
      Height          =   285
      Left            =   840
      TabIndex        =   23
      Top             =   5040
      Width           =   8535
   End
   Begin VB.TextBox txtOption3 
      Height          =   285
      Left            =   840
      TabIndex        =   22
      Top             =   4680
      Width           =   8535
   End
   Begin VB.TextBox txtOption2 
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Top             =   4320
      Width           =   8535
   End
   Begin VB.CommandButton cmdAddToTest 
      Caption         =   "Add Question To Test"
      Height          =   495
      Left            =   9960
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearTemplate 
      Caption         =   "Clear Template"
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtQuestTemplate 
      Height          =   735
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Question Type"
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   3375
      Begin VB.OptionButton optTrueFalse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "True/False"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optMultiple 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Multiple Choice"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdEditBank 
      Caption         =   "Edit Selected Question"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2160
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMoveQuestion 
      Caption         =   "Move Question Into Test Question Template"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemoveFromBank 
      Caption         =   "Remove Question From Current Bank"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddToBank 
      Caption         =   "Add New Question To Current Bank"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtQuestDisp 
      Height          =   1095
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.ListBox lisTestBank 
      Height          =   1620
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Answer Choices With Correct Answer Selected"
      Height          =   1815
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Width           =   9375
      Begin VB.OptionButton opt2Template 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Option1"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton opt4Template 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Option3"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   255
      End
      Begin VB.OptionButton opt3Template 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Option2"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox txtOption1 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   8535
      End
      Begin VB.OptionButton opt1template 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Minutes:"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   8400
      TabIndex        =   32
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lblCurTest 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Current Test:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Test Question Template"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   9
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblQuestType 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Question Type:"
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Test Questions: Click On Question To View It Completely"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblTestBank 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Test Bank: Untitled"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   2
      Left            =   0
      Top             =   5520
      Width           =   12060
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   1
      Left            =   0
      Top             =   2760
      Width           =   12015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000A&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   2775
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuCreate 
         Caption         =   "&Create New Test"
      End
      Begin VB.Menu mnuHypen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Existing Test"
      End
      Begin VB.Menu mnuHyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Test"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Test &As"
      End
      Begin VB.Menu mnuHyphen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Current Test and Key"
      End
      Begin VB.Menu mnuPrintScores 
         Caption         =   "Print Test Scores"
      End
      Begin VB.Menu mnuHyphen4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuLogin 
      Caption         =   "&Login Info"
      Begin VB.Menu mnuCreateStudent 
         Caption         =   "Create Student Account"
      End
      Begin VB.Menu mnuEditStudent 
         Caption         =   "Edit Student Password"
      End
      Begin VB.Menu mnuHyphen5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateTeacher 
         Caption         =   "Create Teacher Account"
      End
      Begin VB.Menu mnuEditPassword 
         Caption         =   "Edit Your Password"
      End
   End
   Begin VB.Menu mnuTestQuestion 
      Caption         =   "Test Question Banks"
      Begin VB.Menu mnuCreateBank 
         Caption         =   "Create New Question Bank"
      End
      Begin VB.Menu mnuHyphen7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenBank 
         Caption         =   "Open Existing Question Bank"
      End
   End
End
Attribute VB_Name = "frmInstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim testHasChanged As Boolean
Dim createCancel As Boolean
Dim questNum As Integer

Private Sub chkAllowGoBack_Click()

    'since check box has changed so has the test
    testHasChanged = True
    
End Sub

Private Sub chkTimed_Click()
    
    'since check box has changed so has the test
    testHasChanged = True
    
End Sub

Private Sub cmdAddToBank_Click()

    'bring up the form to add a new question to question bank
    Load frmAddToBank
    
End Sub

Private Sub cmdAddToTest_Click()
        
    'check for more than 100 questions
    If questNum > 100 Then
        MsgBox "You have already reached the maximum limit of 100.  " & _
               "Remove an existing question if you want to add a new one." _
               , , "Warning!"
        Exit Sub
    End If
    
    'check template for blank values
    If (optMultiple.Value = False And optTrueFalse = False) Or _
       (opt1template.Value = False And opt2Template.Value = False _
        And opt3Template = False And opt4Template.Value = False) Or _
        (txtQuestTemplate.Text = "") Then
        
        MsgBox "One of your fields is blank!", , "Warning!"
        Exit Sub
    End If
    
    'check to see if a current test is already in use, if not
    'create a new one
    
    If currentTest = "" Then
        Call createTest
        questNum = 0
        lblCurTest.Caption = "Current Test: " & currentTest
    End If
               
    testHasChanged = True
    
    'add question to list box and question test array
    lisTestQuests.AddItem txtQuestTemplate
    
    With questTest(lisTestQuests.ListCount)
        .quest = txtQuestTemplate
        .answerA = txtOption1
        .answerB = txtOption2
        .answerC = txtOption3
        .answerD = txtOption4
        
        If optMultiple Then
            .theType = "M"
        Else
            .theType = "T"
        End If
        
        If opt1template Then
            .correctAns = "A"
        Else
            If opt2Template Then
                .correctAns = "B"
            Else
                If opt3Template Then
                    .correctAns = "C"
                Else
                    .correctAns = "D"
                End If
            End If
        End If
        
    End With
    
    
End Sub
Private Sub SaveTest()

    Dim recordNum As Integer
    
    'make sure all files are closed
    Close
    'double check to see if there is at least 1 question
    If lisTestQuests.ListCount > 0 Then
        'select file for output
        Open currentTest For Output As #1
        'write array to file
        For recordNum = 1 To lisTestQuests.ListCount
            Write #1, questTest(recordNum).quest
            Write #1, questTest(recordNum).answerA
            Write #1, questTest(recordNum).answerB
            Write #1, questTest(recordNum).answerC
            Write #1, questTest(recordNum).answerD
            Write #1, questTest(recordNum).correctAns
            Write #1, questTest(recordNum).theType
        Next recordNum
        Close #1
        
        'open file for test layout
        Open Left(currentTest, Len(currentTest) - 3) & "lyt" For Output As #1
        'write the layout options
        If chkTimed Then
            Write #1, "T"
        Else
            Write #1, "F"
        End If
        If chkTimed And Val(txtMinutes) < 1 Then
            Write #1, 1
        Else
            Write #1, Val(txtMinutes)
        End If
        If chkAllowGoBack Then
            Write #1, "T"
        Else
            Write #1, "F"
        End If
        If optGrey.Value = True Then
            Write #1, "G"
        Else
            If optRed.Value = True Then
                Write #1, "R"
            Else
                If optBlue.Value = True Then
                    Write #1, "B"
                Else
                    Write #1, "GN"
                End If
            End If
        End If
        Close #1
        
        'reset test has changed to false
        testHasChanged = False
    End If
        
End Sub
Private Sub createTest()
    
    Dim foundTest As Boolean
    Dim i As Integer
    Dim userResponse As Integer
    
    On Error GoTo dlgError
    
    'if test had changed then ask to save current test before
    'creating a new one
    If testHasChanged Then
        userResponse = MsgBox("The current test has changed since you last saved it. " & _
                              "Do you want to save it before you create a new one?", _
                              vbYesNoCancel, "Warning!")
        If userResponse = vbYes Then
            Call SaveTest
        Else
            If userResponse = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    'get name for new test
    With dlgFile
        .CancelError = True
        .FileName = ""
        .DialogTitle = "Choose A Name For The Test"
        .Flags = 2
        .Filter = "Test Files| *.tst"
        .ShowSave
        
        If Dir(.FileName) <> "" Then
            Kill .FileName
        End If
                
        currentTest = .FileName
        lblCurTest.Caption = "Current Test: " & currentTest
        mnuSave.Enabled = True
        mnuSaveAs.Enabled = True
        cmdReviewQuestion.Enabled = True
        cmdRemoveTestQuest.Enabled = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
        cmdSaveTest.Enabled = True
        
        testHasChanged = False
        
    End With
    
    'add test name to test table in login database
    foundTest = False
    datLogin.RecordSource = "Test"
    datLogin.Refresh
    With datLogin.Recordset
        .MoveFirst
        Do Until .EOF Or foundTest
            If currentTest = RTrim(.Fields("TestName").Value) Then
                foundTest = True
            End If
            .MoveNext
        Loop
        If Not foundTest Then
            .AddNew
            .Fields("TestName").Value = currentTest
            .Update
        End If
    End With
    
    'clear the list box and question test array
    lisTestQuests.Clear
    For i = 1 To 100
        questTest(i).answerA = ""
        questTest(i).answerB = ""
        questTest(i).answerC = ""
        questTest(i).answerD = ""
        questTest(i).correctAns = ""
        questTest(i).quest = ""
        questTest(i).theType = ""
    Next i
    
    'set defaults for new test
    chkTimed.Value = False
    txtMinutes.Text = ""
    chkAllowGoBack.Value = False
    optGrey = True
        
dlgError:
    
    On Error GoTo 0
    Exit Sub
    
End Sub
    
Private Sub cmdClearTemplate_Click()

    'clear all controls in template
    optMultiple.Value = False
    optTrueFalse.Value = False
    opt1template.Value = False
    opt2Template.Value = False
    opt3Template.Value = False
    opt4Template.Value = False
    txtQuestTemplate.Text = ""
    txtOption1.Text = ""
    txtOption2.Text = ""
    txtOption3.Text = ""
    txtOption4.Text = ""
        
End Sub

Private Sub cmdDown_Click()

    Dim questTemp As questionSet
    Dim i As Integer
    Dim listEntries As Integer
    Dim curListSpot As Integer
    
    'make sure question is selected
    If lisTestQuests.Text = "" Then
        MsgBox "Click on a question first.", , "Attention"
    Else
        'make sure selection is not already in the lowest position
        If lisTestQuests.ListIndex = lisTestQuests.ListCount - 1 Then
            Exit Sub
        End If
        
        'testhaschanged needs to be set to true
        testHasChanged = True
        
        'switch question in array with the one below it
        questTemp = questTest(lisTestQuests.ListIndex + 1)
        questTest(lisTestQuests.ListIndex + 1) = questTest(lisTestQuests.ListIndex + 2)
        questTest(lisTestQuests.ListIndex + 2) = questTemp
        
        'clear listbox and reload it
        listEntries = lisTestQuests.ListCount
        curListSpot = lisTestQuests.ListIndex
        lisTestQuests.Clear
        For i = 1 To listEntries
            lisTestQuests.AddItem questTest(i).quest
        Next i
        'select the question that was moved
        lisTestQuests.ListIndex = curListSpot + 1
    End If
    
End Sub

Private Sub cmdEditBank_Click()

    Dim i As Integer
    
    'make sure question is selected
    If lisTestBank.Text <> "" Then
        With datBank.Recordset
            .MoveFirst
            'move to current question in database
            If lisTestBank.ListIndex > 0 Then
                For i = 1 To lisTestBank.ListIndex
                    .MoveNext
                Next i
            End If
            'load question to be edited into edit form
            frmEdit.txtQuestion = RTrim(.Fields("question").Value)
            If .Fields("type").Value = "M" Then
                frmEdit.optMultiple = True
            Else
                frmEdit.optTrueFalse = True
            End If
            frmEdit.txtOptA.Text = RTrim(.Fields("opt1"))
            frmEdit.txtOptB.Text = RTrim(.Fields("opt2"))
            frmEdit.txtOptC.Text = RTrim(.Fields("opt3"))
            frmEdit.txtOptD.Text = RTrim(.Fields("opt4"))
            If RTrim(.Fields("answer").Value) = "A" Then
                frmEdit.optA = True
            Else
                If RTrim(.Fields("answer")) = "B" Then
                    frmEdit.optB = True
                Else
                    If RTrim(.Fields("answer")) = "C" Then
                        frmEdit.optC = True
                    Else
                        frmEdit.optD = True
                    End If
                End If
            End If
        End With
        
        Load frmEdit
        
    Else
    
        MsgBox "Select a question to edit", , "Warning!"
    
    End If
        
End Sub

Private Sub cmdMoveQuestion_Click()

    'check to see if a question from bank is clicked on, if so
    'move question, question type, options, and correct anwser
    'to template
    
    If lisTestBank.Text = "" Then
        MsgBox "Click on a question first.", , "Attention"
    Else
        With questHold(lisTestBank.ListIndex + 1)
        txtQuestTemplate.Text = .quest
        txtOption1.Text = .answerA
        txtOption2.Text = .answerB
        txtOption3.Text = .answerC
        txtOption4.Text = .answerD
        If .correctAns = "A" Then
            opt1template.Value = True
            opt2Template.Value = False
            opt3Template.Value = False
            opt4Template.Value = False
        Else
            If .correctAns = "B" Then
                opt2Template.Value = True
                opt1template.Value = False
                opt3Template.Value = False
                opt4Template.Value = False
            Else
                If .correctAns = "C" Then
                    opt3Template.Value = True
                    opt1template.Value = False
                    opt2Template.Value = False
                    opt4Template.Value = False
                Else
                    opt4Template.Value = True
                    opt1template.Value = False
                    opt2Template.Value = False
                    opt3Template.Value = False
                End If
            End If
        End If
        If .theType = "M" Then
            optMultiple.Value = True
        Else
            optTrueFalse.Value = True
            txtOption3.Text = ""
            txtOption4.Text = ""
            opt3Template.Value = False
            opt4Template.Value = False
        End If
        End With
        
    End If
        
End Sub

Private Sub cmdRemoveFromBank_Click()
    Dim questRem As String
    
    'check to see if the question to be removed from the bank
    'has been clicked on
    If lisTestBank.Text = "" Then
        MsgBox "You have not selected a question to remove.", , "Warning!"
    Else
        
        'remove question from database
        questRem = questHold(lisTestBank.ListIndex + 1).quest
        With datBank.Recordset
        .MoveFirst
        .FindFirst ("question = '" & questRem & "'")
        .Delete
        .MoveFirst
        End With
        
        'remove question from bank list
        lisTestBank.RemoveItem lisTestBank.ListIndex
                
    End If
    
End Sub

Private Sub cmdRemoveTestQuest_Click()

    Dim i As Integer
    
    'make sure quesition is selected
    If lisTestQuests.Text = "" Then
        MsgBox "Click on a question first.", , "Attention"
    Else
        'delete current question from array
        testHasChanged = True
        For i = lisTestQuests.ListIndex + 1 To lisTestQuests.ListCount - 1
            questTest(i) = questTest(i + 1)
        Next i
        'remove question from listbox
        lisTestQuests.RemoveItem lisTestQuests.ListIndex
    End If
    
End Sub

Private Sub cmdReviewQuestion_Click()
    
    'make sure a question is selected
    If lisTestQuests.Text = "" Then
        MsgBox "Click on a question first.", , "Attention"
    Else
        'load question into template section
        With questTest(lisTestQuests.ListIndex + 1)
        txtQuestTemplate.Text = .quest
        txtOption1.Text = .answerA
        txtOption2.Text = .answerB
        txtOption3.Text = .answerC
        txtOption4.Text = .answerD
        If .correctAns = "A" Then
            opt1template.Value = True
            opt2Template.Value = False
            opt3Template.Value = False
            opt4Template.Value = False
        Else
            If .correctAns = "B" Then
                opt2Template.Value = True
                opt1template.Value = False
                opt3Template.Value = False
                opt4Template.Value = False
            Else
                If .correctAns = "C" Then
                    opt3Template.Value = True
                    opt1template.Value = False
                    opt2Template.Value = False
                    opt4Template.Value = False
                Else
                    opt4Template.Value = True
                    opt1template.Value = False
                    opt2Template.Value = False
                    opt3Template.Value = False
                End If
            End If
        End If
        If .theType = "M" Then
            optMultiple.Value = True
        Else
            optTrueFalse.Value = True
            txtOption3.Text = ""
            txtOption4.Text = ""
            opt3Template.Value = False
            opt4Template.Value = False
        End If
        End With
        
    End If
        
End Sub

Private Sub cmdSaveTest_Click()
    
    'make sure test has changed and there is at least 1 question
    If testHasChanged And lisTestQuests.ListCount > 0 Then
        Call SaveTest
    End If
    
End Sub

Private Sub cmdUp_Click()

    Dim questTemp As questionSet
    Dim i As Integer
    Dim listEntries As Integer
    Dim curListSpot As Integer
    
    'make sure a question has been selected
    If lisTestQuests.Text = "" Then
        MsgBox "Click on a question first.", , "Attention"
    Else
        'don't perform sub if the selected question is already
        'at the top position
        If lisTestQuests.ListIndex = 0 Then
            Exit Sub
        End If
        
        'set testHasChanged to true since it has changed
        testHasChanged = True
        
        'switch question position in arrray with the 1 above it
        questTemp = questTest(lisTestQuests.ListIndex)
        questTest(lisTestQuests.ListIndex) = questTest(lisTestQuests.ListIndex + 1)
        questTest(lisTestQuests.ListIndex + 1) = questTemp
        
        'clear and reload list box
        listEntries = lisTestQuests.ListCount
        curListSpot = lisTestQuests.ListIndex
        lisTestQuests.Clear
        For i = 1 To listEntries
            lisTestQuests.AddItem questTest(i).quest
        Next i
        'highlight the moved question
        lisTestQuests.ListIndex = curListSpot - 1
    End If
    
End Sub

Private Sub Form_Load()

    'make form visible and unload login form
    Show
    Unload frmLogin
    
    With datLogin
        .DatabaseName = App.Path & "\login.mdb"
        .RecordSource = "Login"
        .Refresh
    End With
    
    'set caption of form to include user ID
    Caption = Caption & " - User ID : " & userCode
    
    'if logged in with default system password disable menu
    'option to edit password
    If UCase(userCode) = "SYSTEM999" Then
        mnuEditPassword.Enabled = False
    End If
    
    mnuSave.Enabled = False
    mnuSaveAs.Enabled = False
    
    'disable Add to bank, remove from bank, and edit from bank
    'command buttons
    cmdAddToBank.Enabled = False
    cmdRemoveFromBank.Enabled = False
    cmdEditBank.Enabled = False
    
    'reset newDb to false
    newDB = False
    
    'reset test has changed to false
    testHasChanged = False
      
End Sub

Private Sub lisTestBank_Click()
    
    'load question from bank list to large question display
    txtQuestDisp.Text = lisTestBank.Text
    
    'set the label to display correct question type
    Select Case questHold(lisTestBank.ListIndex + 1).theType
        Case "M"
            lblQuestType.Caption = "Question Type: " & "Multiple Choice"
        Case "T"
            lblQuestType.Caption = "Question Type: " & "True/False"
    End Select
        
End Sub


Private Sub mnuCreate_Click()
    
    Call createTest

End Sub

Private Sub mnuCreateBank_Click()
    
    Dim NewBank As Database, MyWS As Workspace
    Dim T1 As TableDef
    Dim T1Flds(1 To 7) As Field
    Dim T1Idx As Index
    Dim myRec As Recordset
    Dim checkDIR As String
    
    On Error GoTo DialogError
    
    'open save dialog
    With dlgFile
        .CancelError = True
        .DialogTitle = "Choose A Name For The Test Bank"
        .Flags = cdlOFNOverwritePrompt
        .Filter = "Question Bank Files|*.mdb"
        .ShowSave
        
        datBank.Database.Close
        
        'delete the file chosen if it exists
        If Dir(.FileName) <> "" Then
            Kill .FileName
        End If
        
        'set currentBank with file name
        currentBank = .FileName
        
        'set test bank label with file name
        lblTestBank.Caption = "Test Bank: '" & .FileName & "'"
        
        'set Add to bank, remove from bank, and edit from bank
        'command buttons
        cmdAddToBank.Enabled = True
        cmdRemoveFromBank.Enabled = True
        cmdEditBank.Enabled = True
        
        'create new question bank database
        Set MyWS = DBEngine.Workspaces(0)
        Set NewBank = MyWS.CreateDatabase(.FileName, dbLangGeneral)
        
        Set T1 = NewBank.CreateTableDef("bank")
        
        Set T1Flds(1) = T1.CreateField("question", dbText, 250)
        Set T1Flds(2) = T1.CreateField("type", dbText, 1)
        Set T1Flds(3) = T1.CreateField("opt1", dbText, 50)
        Set T1Flds(4) = T1.CreateField("opt2", dbText, 50)
        Set T1Flds(5) = T1.CreateField("opt3", dbText, 50)
        Set T1Flds(6) = T1.CreateField("opt4", dbText, 50)
        Set T1Flds(7) = T1.CreateField("answer", dbText, 50)
        
        T1.Fields.Append T1Flds(1)
        T1.Fields.Append T1Flds(2)
        T1.Fields.Append T1Flds(3)
        T1.Fields.Append T1Flds(4)
        T1.Fields.Append T1Flds(5)
        T1.Fields.Append T1Flds(6)
        T1.Fields.Append T1Flds(7)
        
        Set T1Idx = T1.CreateIndex("question")
        T1Idx.Primary = True
        T1Idx.Unique = True
        T1Idx.Required = True
        Set T1Flds(1) = T1Idx.CreateField("question")
        T1Idx.Fields.Append T1Flds(1)
        T1.Indexes.Append T1Idx
        NewBank.TableDefs.Append T1
        
        'add a dummy record to prevent any db errors
        Set myRec = T1.OpenRecordset
        myRec.AddNew
        myRec("question") = "dummy"
        myRec("type") = "D"
        myRec("opt1") = "dummy"
        myRec("opt2") = "dummy"
        myRec("opt3") = "dummy"
        myRec("opt4") = "dummy"
        myRec("answer") = "d"
        myRec.Update
        myRec.Close
        
        'close new database
        NewBank.Close
                
        'open new database with data control
        datBank.DatabaseName = .FileName
        datBank.RecordSource = "bank"
        datBank.Refresh
        
        'reset questHold array to null
        For x = 1 To 200
            questHold(x).answerA = ""
            questHold(x).answerB = ""
            questHold(x).answerC = ""
            questHold(x).answerD = ""
            questHold(x).correctAns = ""
            questHold(x).quest = ""
            questHold(x).theType = ""
        Next x
        
        'set newDB flag to true
        newDB = True
        
        'clear test bank list box
        lisTestBank.Clear
        
        'clear large question display
        txtQuestDisp.Text = ""
        
        'load form to add questions to bank
        Load frmAddToBank
                
    End With
    
DialogError:
    On Error GoTo 0
    Exit Sub
    
End Sub

Private Sub mnuCreateStudent_Click()

    'load form and set caption to student
    Load frmAccount
    frmAccount.Caption = "Student"
    
End Sub

Private Sub mnuCreateTeacher_Click()

    'load form and set caption to teacher
    Load frmAccount
    frmAccount.Caption = "Teacher"
    
End Sub

Private Sub mnuEditPassword_Click()

    'edit currently logged on user's password
    Load frmTeachPass
    'search for user in database and display current(old) password
    With frmInstructor.datLogin.Recordset
        .MoveFirst
        Do Until .EOF
            If loggedUser = .Fields("UserID").Value Then
                frmTeachPass.picOutput.Cls
                frmTeachPass.picOutput.Print "Old Password:  "; .Fields("Password").Value
                Exit Do
            End If
            .MoveNext
        Loop
    End With
                
End Sub

Private Sub mnuEditStudent_Click()

    Load frmStudPass
    
End Sub

Private Sub mnuExit_Click()
    
    Dim response As Integer
    
    'before exiting check to see if test has changed and give
    'user option to save or not or cancel exit
    If testHasChanged Then
        response = MsgBox("Test has changed, do you want to save before exiting program?", vbYesNoCancel _
                   , "Warning!")
        If response = vbYes Then
            Call SaveTest
        Else
            If response = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    'stop application
    End
    
End Sub


Private Sub mnuOpen_Click()

    Dim i As Integer
    Dim userResponse As Integer
    Dim timed As String
    Dim minutes As Integer
    Dim goBack As String
    Dim theColor As String
    
    On Error GoTo dlgError
    
    'if test has changed check to see if user wants to save current
    'test before opening new one
    If testHasChanged Then
        userResponse = MsgBox("The current test has changed since you last saved it. " & _
                              "Do you want to save it before you open another one?", _
                              vbYesNoCancel, "Warning!")
        If userResponse = vbYes Then
            Call SaveTest
        Else
            If userResponse = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    'select filename to open
    With dlgFile
        .CancelError = True
        .DialogTitle = "Choose Test To Open"
        .Filter = "Test Files|*.tst"
        .Flags = 2
        .ShowOpen
        
        Open .FileName For Input As #1
        currentTest = .FileName
        lblCurTest.Caption = "Current Test: " & currentTest
        mnuSave.Enabled = True
        mnuSaveAs.Enabled = True
        cmdReviewQuestion.Enabled = True
        cmdRemoveTestQuest.Enabled = True
        cmdUp.Enabled = True
        cmdDown.Enabled = True
        cmdSaveTest.Enabled = True
        
        lisTestQuests.Clear
        
        'load test question array from test file
        Do Until EOF(1)
            i = i + 1
            Input #1, questTest(i).quest
            Input #1, questTest(i).answerA
            Input #1, questTest(i).answerB
            Input #1, questTest(i).answerC
            Input #1, questTest(i).answerD
            Input #1, questTest(i).correctAns
            Input #1, questTest(i).theType
            lisTestQuests.AddItem questTest(i).quest
        Loop
        Close #1
        
        'load test options
        Open Left(.FileName, Len(.FileName) - 3) & "lyt" For Input As #1
            Input #1, timed
            Input #1, minutes
            Input #1, goBack
            Input #1, theColor
        Close
        
        If timed = "T" Then
            chkTimed.Value = 1
            txtMinutes = minutes
        Else
            chkTimed.Value = 0
            txtMinutes = ""
        End If
        If goBack = "T" Then
            chkAllowGoBack.Value = 1
        Else
            chkAllowGoBack.Value = 0
        End If
        If theColor = "G" Then
            optGrey.Value = True
        Else
            If theColor = "R" Then
                optRed.Value = True
            Else
                If theColor = "B" Then
                    optBlue.Value = True
                Else
                    optGreen.Value = True
                End If
            End If
        End If
        
    End With
    
    'reset test has changed to false
    testHasChanged = False
        
dlgError:
    On Error GoTo 0
    Exit Sub
    
End Sub

Private Sub mnuOpenBank_Click()

    Dim i As Integer
    Dim x As Integer
    
    On Error GoTo DialogError
    
    'open the open dialog
    With dlgFile
        .CancelError = True
        .DialogTitle = "Choose The Name Of The Test Bank"
        .Flags = 2
        .Filter = "Database Files|*.mdb"
        .ShowOpen
              
        'set database,label, and currentBank with file name
        datBank.DatabaseName = .FileName
        lblTestBank.Caption = "Test Bank:  '" & .FileName & "'"
        currentBank = .FileName
    End With
    
    'enable add to bank, romove from bank, and edit bank
    'command buttons
    cmdAddToBank.Enabled = True
    cmdRemoveFromBank.Enabled = True
    cmdEditBank.Enabled = True
    
    'clear big question display and test bank list box
    txtQuestDisp.Text = ""
    lisTestBank.Clear
    
    'open database
    datBank.RecordSource = "bank"
    datBank.Refresh
    datBank.Recordset.MoveFirst
    
    'clear questHold array with null
    For x = 1 To 200
        questHold(x).answerA = ""
        questHold(x).answerB = ""
        questHold(x).answerC = ""
        questHold(x).answerD = ""
        questHold(x).correctAns = ""
        questHold(x).quest = ""
        questHold(x).theType = ""
    Next x
    
    'load questHold array with test bank database data
    Do Until datBank.Recordset.EOF
        i = i + 1
        With questHold(i)
            .quest = datBank.Recordset.Fields("question").Value
            .theType = datBank.Recordset.Fields("type").Value
            .answerA = datBank.Recordset.Fields("opt1").Value
            .answerB = datBank.Recordset.Fields("opt2").Value
            .answerC = datBank.Recordset.Fields("opt3").Value
            .answerD = datBank.Recordset.Fields("opt4").Value
            .correctAns = datBank.Recordset.Fields("answer").Value
            lisTestBank.AddItem (.quest)
        End With
        datBank.Recordset.MoveNext
    Loop
    
    If i = 1 And RTrim(datBank.Recordset.Fields("question").Value) = "dummy" Then
        newDB = True
    End If
    
    'set question type label
    lblQuestType.Caption = "Question Type:"
                
DialogError:
    On Error GoTo 0
    Exit Sub
    
End Sub

Private Sub mnuPrint_Click()

    Dim i As Integer
    Dim linecount As Integer
    
    'print the current test with font size 12
    Printer.FontSize = 12
    For i = 1 To 5
        Printer.Print
    Next i
    Printer.Print Space(80 - (Len(currentTest) / 2)); currentTest
    Printer.Print
    
    'print questions
    For i = 1 To lisTestQuests.ListCount
        Printer.FontSize = 10
        Printer.Print i & "." & questTest(i).quest
        Printer.Print
        linecount = linecount + 2
        If questTest(i).theType = "T" Then
            Printer.Print "A. " & questTest(i).answerA
            Printer.Print "B. " & questTest(i).answerB
            linecount = linecount + 3
        Else
            Printer.Print "A. " & questTest(i).answerA
            Printer.Print "B. " & questTest(i).answerB
            Printer.Print "C. " & questTest(i).answerC
            Printer.Print "D. " & questTest(i).answerD
            linecount = linecount + 5
        End If
        Printer.Print
        If linecount > 55 Then
            Printer.NewPage
            Printer.FontSize = 12
            Printer.Print: Printer.Print: Printer.Print: Printer.Print
            Printer.Print: Printer.Print: Printer.Print
            Printer.FontSize = 10
            linecount = 0
        End If
    Next i
    
    'print test key
    Printer.NewPage
    Printer.Print: Printer.Print
    Printer.Print currentTest & " - Test Key"
    For i = 1 To lisTestQuests.ListCount
        Printer.Print
        Printer.Print i & ": " & questTest(i).correctAns
    Next i
    
    Printer.EndDoc
    
End Sub

Private Sub mnuPrintScores_Click()

    Load frmTestScores
    
End Sub

Private Sub mnuSave_Click()

    'make sure test has changed and that there is at least
    '1 question in the test
    If testHasChanged And lisTestQuests.ListCount > 0 Then
        Call SaveTest
    End If
    
End Sub

Private Sub mnuSaveAs_Click()

    'make sure there are questions in the test
    If lisTestQuests.ListCount > 0 Then
    
        'select filename to save as
        With dlgFile
        .CancelError = True
        .FileName = ""
        .DialogTitle = "Choose A Name For The Test"
        .Flags = 2
        .Filter = "Test Files| *.tst"
        .ShowSave
        
        'if file did exist, delete it
        If Dir(.FileName) <> "" Then
            Kill .FileName
        End If
                
        'set currentTest to the filename selected
        'enable save and saves menu selections
        currentTest = .FileName
        mnuSave.Enabled = True
        mnuSaveAs.Enabled = True
                
        End With
        
        'call sub to actually save the test
        Call SaveTest
    End If
    
dlgError:
    On Error GoTo 0
    Exit Sub
    
    
    
End Sub

Private Sub optBlue_Click()

    'if this option is changed then test has changed
    testHasChanged = True
    
End Sub

Private Sub optGreen_Click()

    'if this option is changed then test has changed
    testHasChanged = True
    
End Sub

Private Sub optGrey_Click()

    'if this option is changed then test has changed
    testHasChanged = True
    
End Sub

Private Sub optMultiple_Click()

    'enable options 3 and 4 if option multiple choice is choosen
    opt3Template.Enabled = True
    opt4Template.Enabled = True
    
End Sub

Private Sub optRed_Click()

    'if this option is changed then test has changed
    testHasChanged = True
    
End Sub

Private Sub optTrueFalse_Click()

    'disable and clear options 3 and 4 if true/false choice is choosen
    txtOption1.Text = "True"
    txtOption2.Text = "False"
    txtOption3.Text = ""
    txtOption4.Text = ""
    opt3Template.Value = False
    opt4Template.Value = False
    opt3Template.Enabled = False
    opt4Template.Enabled = False
    
End Sub

Private Sub txtMinutes_Change()
    
    'if this option is changed then test has changed
    testHasChanged = True
    
End Sub
