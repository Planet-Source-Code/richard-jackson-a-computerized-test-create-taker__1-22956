Attribute VB_Name = "TestProgram"
'Final Project:   Computerized Test
'Name: Richard Jackson
'Purpose:   This project allows teachers to log on and create
'           and edit: tests, question banks, student accounts,
'           and teacher accounts.  Also, the teacher can view
'           print test scores and tests with keys.  Students
'           can log on and take tests, print and display scores,
'           and edit their password.
'
'TestProgram.BAS Module
'

Public usersAnswer(1 To 100) As String
Public userScore As Integer
Public numCorrect As Integer
Public numOfQ As Integer
Public loggedUser As String
Public userCode As String
Public password As String
Public currentTest As String
Public currentBank  As String
Public r As Integer, g As Integer, b As Integer
Public newDB As Boolean

Type questionSet
    quest As String * 250
    theType As String * 1
    answerA As String * 50
    answerB As String * 50
    answerC As String * 50
    answerD As String * 50
    correctAns As String * 1
End Type

Public questHold(1 To 200) As questionSet
Public questTest(1 To 100) As questionSet


    
