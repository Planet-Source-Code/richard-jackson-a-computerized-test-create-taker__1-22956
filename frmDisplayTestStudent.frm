VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDisplayTestStudent 
   AutoRedraw      =   -1  'True
   Caption         =   "Test Scores"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.Data datScores 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmDisplayTestStudent.frx":0000
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmDisplayTestStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim i As Integer
    
    Show
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    'set column width
    MSFlexGrid1.ColWidth(0) = 5100
    For i = 1 To 2
        MSFlexGrid1.ColWidth(i) = 900
    Next i
    
    'select testscores from the currently logged user
    With datScores
        .DatabaseName = App.Path & "\login.mdb"
        .RecordSource = "SELECT Test, Date, Grade " & _
                        "FROM TestScores " & _
                        "WHERE ID = '" & loggedUser & "'"
        .Refresh
    End With
    
End Sub
