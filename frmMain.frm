VERSION 5.00
Begin VB.Form frmMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple Users With Napster"
   ClientHeight    =   1740
   ClientLeft      =   5580
   ClientTop       =   4695
   ClientWidth     =   3660
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3660
   Begin VB.CommandButton cmdSetUser 
      Caption         =   "&Set User"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Just Set User"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run Napster"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Set User and Run Napster"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmbUser 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmMain.frx":0CCA
      Left            =   323
      List            =   "frmMain.frx":0CCC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Choose User"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&User:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strNapPath As String

Private Sub cmdRun_Click()
    Dim blnComp As Boolean
    If cmbUser.ListIndex = 0 Then
        blnComp = SetUser("") 'Create New user
    Else
        blnComp = SetUser(cmbUser.Text) 'Set user
    End If
    If blnComp = True Then
        Shell strNapPath & "\napster.exe", vbNormalFocus 'Run Napster
    End If
End Sub

Private Sub cmdSetUser_Click()
    Dim blnComp As Boolean
    If cmbUser.ListIndex = 0 Then
        blnComp = SetUser("") 'Set User to null string to cause napster to create a new user
    Else
        blnComp = SetUser(cmbUser.Text) 'Set current user
    End If
End Sub


Private Sub Form_Load()
    Dim strUsers() As String, blnComp As Boolean
    strNapPath = GetNapPath 'Get Napster path
    cmbUser.AddItem "New User"
    blnComp = GetUserNames(strUsers)
    If blnComp = True Then
        For i = LBound(strUsers) To UBound(strUsers)
            cmbUser.AddItem strUsers(i)
        Next i
    End If
    If GetCurUser <> "" Then
        cmbUser.Text = GetCurUser
    End If
End Sub
