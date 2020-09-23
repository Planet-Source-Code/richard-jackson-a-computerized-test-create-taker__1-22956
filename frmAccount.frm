VERSION 5.00
Begin VB.Form frmAccount 
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOutput 
      Height          =   1335
      Left            =   4560
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateAccount 
      Caption         =   "Create Account"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtSSN 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtLast 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtFirst 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Social Security Number"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdCreateAccount_Click()

    Dim password As String
    Dim userid As String
    Dim passchar As Integer
    Dim i As Integer
    
    'check for any blank fields
    If txtFirst <> "" And txtLast <> "" And txtSSN <> "" Then
        'generate random password
        Randomize Timer
        For i = 1 To 5
            passchar = Int(26 * Rnd) + 65
            password = password + Chr$(passchar)
        Next i
        'obtain first 4 characters from last name and first name
        'if last name is shorter than 4 characters
        If Len(txtLast) < 4 Then
            userid = UCase(txtLast & Left(txtFirst, 4 - Len(txtLast)))
        Else
            userid = UCase(Left(txtLast, 4))
        End If
        'create user ID
        userid = userid & Right(txtSSN, 4)
        'display the user ID and password to user
        picOutput.Cls
        picOutput.Print "User I.D. is "; userid
        picOutput.Print "Password is "; password
        'add new account to login database
        With frmInstructor.datLogin.Recordset
            .AddNew
            .Fields("UserID").Value = userid
            .Fields("LastName").Value = txtLast
            .Fields("FirstName").Value = txtFirst
            .Fields("SSN").Value = Right(txtSSN, 4)
            .Fields("Password").Value = password
            If Caption = "Student" Then
                .Fields("Instructor").Value = False
            Else
                .Fields("Instructor").Value = True
            End If
            .Update
        End With
    Else
        MsgBox "One of the fields is empty!", , "Warning!"
    End If
    
End Sub

Private Sub Form_Load()

    'display and center form
    Show
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
End Sub
