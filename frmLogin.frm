VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datLogin 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00C00000&
      Height          =   4815
      Left            =   120
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login Screen"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   4
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   960
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Computerized Testing Program"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
        
    End
    
End Sub

Private Sub cmdLogin_Click()

    Dim found As Boolean
    Dim instructor As Boolean
    
    'reset found flag to false
    found = False
    
    'check for blank login fields
    If txtPassword.Text = "" Or txtUserID.Text = "" Then
        MsgBox "One of your login fields is blank, pleae try again.", , "Attention"
        Exit Sub
    End If
    
    'search userID to see if it exists
    datLogin.Recordset.MoveFirst
    userCode = datLogin.Recordset.Fields("UserID").Value
    Do Until found Or datLogin.Recordset.EOF
        userCode = datLogin.Recordset.Fields("UserID").Value
        If UCase(RTrim(userCode)) = UCase(txtUserID.Text) Then
            found = True
            Exit Do
        Else
            datLogin.Recordset.MoveNext
        End If
    Loop
        
    If found Then
        'check password if found
        password = datLogin.Recordset.Fields("Password").Value
        instructor = datLogin.Recordset.Fields("Instructor").Value
        If UCase(password) = UCase(txtPassword.Text) Then
            'load appropiate form for student or instructor
            loggedUser = UCase(userCode)
            If instructor Then
                Load frmInstructor
            Else
                Load frmStudent
            End If
        Else
            MsgBox "Password is incorrect!", , "Warning!!"
        End If
    Else
        MsgBox "User ID was not found, try again.", , "Warning!"
    End If
                
End Sub

Private Sub Form_Load()
    
    'center form on screen
    frmLogin.Left = (Screen.Width - Width) / 2
    frmLogin.Top = (Screen.Height - Height) / 2
    'set color for login screen
    r = 204
    g = 195
    b = 175
    BackColor = RGB(r, g, b)
    Label1.BackColor = RGB(r, g, b)
    cmdLogin.BackColor = RGB(r, g, b)
    cmdExit.BackColor = RGB(r, g, b)
    
    With datLogin
        .DatabaseName = App.Path & "\login.mdb"
        .RecordSource = "Login"
        .Refresh
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'close login database
    datLogin.Recordset.Close
    
End Sub


