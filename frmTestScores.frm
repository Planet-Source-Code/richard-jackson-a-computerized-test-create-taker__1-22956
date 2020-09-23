VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmTestScores 
   Caption         =   "Test Scores"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00FFFFC0&
      Height          =   7575
      Left            =   3240
      ScaleHeight     =   7515
      ScaleWidth      =   7035
      TabIndex        =   8
      Top             =   120
      Width           =   7095
   End
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Scores For Selected Student"
      Height          =   2415
      Left            =   360
      TabIndex        =   3
      Top             =   4200
      Width           =   2775
      Begin MSDBCtls.DBCombo cmbStudent 
         Bindings        =   "frmTestScores.frx":0000
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "UserID"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.CommandButton cmdDisplayStudent 
         Caption         =   "Display"
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintStudent 
         Caption         =   "Print"
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Scores For Selected Test"
      Height          =   2415
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      Begin MSDBCtls.DBCombo cmbTest 
         Bindings        =   "frmTestScores.frx":0017
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "TestName"
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.CommandButton cmdDisplayTest 
         Caption         =   "Display"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrintTest 
         Caption         =   "Print"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Data datScores 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data datTest 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmTestScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbStudent_Click(Area As Integer)

    'when user selects an user ID, display the user's name as
    'a tool tip
    With datLogin.Recordset
        .MoveFirst
        .FindFirst ("UserID = '" & cmbStudent & "'")
        cmbStudent.ToolTipText = RTrim(.Fields("FirstName").Value) & " " & _
                      .Fields("LastName").Value
    End With
    
End Sub

Private Sub cmdDisplayStudent_Click()

    Dim numOfTests As Integer
    Dim totalScore As Integer
    Dim average As Single
    
    'move to correct records and display info
    picDisplay.Cls
    With datScores.Recordset
        .MoveFirst
        Do Until .EOF
            If cmbStudent.Text = .Fields("ID").Value Then
                numOfTests = numOfTests + 1
                totalScore = totalScore + .Fields("Grade").Value
                picDisplay.Print "Test: " & .Fields("Test").Value; _
                                 Tab(25); "I.D. "; .Fields("ID").Value; _
                                 Tab(55); .Fields("Date").Value; _
                                 Tab(70); .Fields("Grade").Value
            End If
            .MoveNext
        Loop
        
        'calculate average if any tests existed
        If numOfTests > 0 Then
            average = totalScore / numOfTests
            picDisplay.Print
            picDisplay.Print "Average of tests for student " & cmbStudent.Text & _
                             " is " & FormatNumber(average, 1)
        Else
            MsgBox "Test Scores Not Available", , "Attention"
        End If
    End With
    
End Sub

Private Sub cmdDisplayTest_Click()

    
    Dim numOfTests As Integer
    Dim totalScore As Integer
    Dim average As Single
    
    'move to current records and display
    picDisplay.Cls
    With datScores.Recordset
        .MoveFirst
        Do Until .EOF
            If cmbTest.Text = .Fields("Test").Value Then
                numOfTests = numOfTests + 1
                totalScore = totalScore + .Fields("Grade").Value
                picDisplay.Print "Test: " & cmbTest.Text; _
                                 Tab(25); "I.D. "; .Fields("ID").Value; _
                                 Tab(55); .Fields("Date").Value; _
                                 Tab(70); .Fields("Grade").Value
            End If
            .MoveNext
        Loop
        
        'if tests existed then display average
        If numOfTests > 0 Then
            average = totalScore / numOfTests
            picDisplay.Print
            picDisplay.Print "Average of test " & cmbTest.Text & _
                             " is " & FormatNumber(average, 1)
        Else
            MsgBox "Test Scores Not Available", , "Attention"
        End If
    End With
    
End Sub

Private Sub cmdPrintStudent_Click()

    Dim numOfTests As Integer
    Dim totalScore As Integer
    Dim average As Single
    
    Printer.FontSize = 12
    'find records and print
    With datScores.Recordset
        .MoveFirst
        Do Until .EOF
            If cmbStudent.Text = .Fields("ID").Value Then
                numOfTests = numOfTests + 1
                totalScore = totalScore + .Fields("Grade").Value
                Printer.Print "Test: " & .Fields("Test").Value; _
                                 Tab(35); "I.D. "; .Fields("ID").Value; _
                                 Tab(55); .Fields("Date").Value; _
                                 Tab(70); .Fields("Grade").Value
            End If
            .MoveNext
        Loop
        
        'if tests existed then print the average
        If numOfTests > 0 Then
            average = totalScore / numOfTests
            Printer.Print
            Printer.Print "Average of tests for student " & cmbStudent.Text & _
                          " is " & FormatNumber(average, 1)
            Printer.EndDoc
        Else
            MsgBox "Test Scores Not Available", , "Attention"
        End If
    End With
    
End Sub

Private Sub cmdPrintTest_Click()

    Dim numOfTests As Integer
    Dim totalScore As Integer
    Dim average As Single
    
    Printer.FontSize = 12
    'find correct records and print
    With datScores.Recordset
        .MoveFirst
        Do Until .EOF
            If cmbTest.Text = .Fields("Test").Value Then
                numOfTests = numOfTests + 1
                totalScore = totalScore + .Fields("Grade").Value
                Printer.Print "Test: " & cmbTest.Text, _
                              "I.D. " & .Fields("ID").Value, _
                                 .Fields("Date").Value, _
                                 .Fields("Grade").Value
            End If
            .MoveNext
        Loop
        
        'if tests existed then print the average
        If numOfTests > 0 Then
            average = totalScore / numOfTests
            Printer.Print
            Printer.Print "Average of test " & cmbTest.Text & _
                        " is " & FormatNumber(average, 1)
            Printer.EndDoc
        Else
            MsgBox "Test Scores Not Available", , "Attention"
        End If
    End With
End Sub

Private Sub Form_Load()

    Show
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    datLogin.DatabaseName = App.Path & "\login.mdb"
    datLogin.RecordSource = "Login"
    datLogin.Refresh
    
    datScores.DatabaseName = App.Path & "\login.mdb"
    datScores.RecordSource = "TestScores"
    datScores.Refresh
    
    datTest.DatabaseName = App.Path & "\login.mdb"
    datTest.RecordSource = "Test"
    datTest.Refresh
    
End Sub

