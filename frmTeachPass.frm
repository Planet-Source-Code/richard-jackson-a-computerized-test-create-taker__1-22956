VERSION 5.00
Begin VB.Form frmTeachPass 
   Caption         =   "Change Teacher Password"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdSavePass 
      Caption         =   "Save New Password"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox picOutput 
      Height          =   615
      Left            =   600
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "New Password"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmTeachPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSavePass_Click()

    'make sure password is 5 characters long
    If Len(txtPassword) = 5 Then
        'update password in database
        With datLogin.Recordset
            .MoveFirst
            .FindFirst "UserID = '" & loggedUser & "'"
            .Edit
            .Fields("Password").Value = UCase(txtPassword)
            .Update
            MsgBox "Password Changed!!", , "Success!"
            Unload Me
        End With
    Else
        MsgBox "Password must contain 5 characters", , "Warning!!"
    End If
    
End Sub

Private Sub Form_Load()

    'display form
    Show
    txtPassword.SetFocus
    datLogin.DatabaseName = App.Path & "\login.mdb"
    datLogin.RecordSource = "Login"
    datLogin.Refresh
    
End Sub
