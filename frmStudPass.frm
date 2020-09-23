VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmStudPass 
   Caption         =   "Change Student Password"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSavePass 
      Caption         =   "Save New Password"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.PictureBox picOutput 
      Height          =   855
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBList cmbUser 
      Bindings        =   "frmStudPass.frx":0000
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2858
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "New Password"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Click On Student To Change"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmStudPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbUser_Click()

    Dim found As Boolean
    
    found = False
    picOutput.Cls
    'move through the database until correct user ID is found
    With datLogin.Recordset
        .MoveFirst
        Do Until .EOF
            If cmbUser.Text = .Fields("UserID").Value Then
                found = True
                'print userID, Name, and old password
                picOutput.Print "User:"; Tab(18); cmbUser.Text
                picOutput.Print "Name:"; Tab(18); RTrim(.Fields("FirstName").Value) & _
                                " " & .Fields("LastName").Value
                picOutput.Print ; "Old Password:"; Tab(18); .Fields("Password").Value
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    
End Sub

Private Sub cmdSavePass_Click()
    
    'make sure password is 5 characters long
    If Len(txtPassword) = 5 Then
        'update password in database
        With datLogin.Recordset
            .Edit
            .Fields("Password").Value = UCase(txtPassword)
            .Update
            MsgBox "Password Has Been Changed!", , "Successful!"
        End With
    Else
        MsgBox "Password must contain 5 characters!", , "Warning!"
    End If
    
End Sub

Private Sub Form_Load()

    'display and center form
    Show
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    'select only students from database with SQL
    With datLogin
        .DatabaseName = App.Path & "\login.mdb"
        .RecordSource = "SELECT UserID, Password, FirstName, LastName " & _
                                "FROM Login " & _
                                "WHERE Instructor = False"
        .Refresh
        cmbUser.ListField = "UserID"
    End With
    
End Sub
