VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "User Password"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   225
      TabIndex        =   8
      Top             =   750
      Width           =   4665
      Begin VB.ComboBox cboUser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2940
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1215
         PasswordChar    =   "•"
         TabIndex        =   9
         Top             =   780
         Width           =   2940
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   840
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4230
         Picture         =   "Login.frx":0000
         Top             =   810
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   4230
         Picture         =   "Login.frx":038A
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   30
      Top             =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change &Time(F6)"
      Height          =   435
      Left            =   4950
      TabIndex        =   0
      Top             =   870
      Width           =   1740
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   4950
      TabIndex        =   1
      Top             =   1372
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   4950
      TabIndex        =   2
      Top             =   1875
      Width           =   1740
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "hh:mm:ss"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1455
      TabIndex        =   7
      Top             =   2760
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "mm-dd-yyyy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1425
      TabIndex        =   6
      Top             =   2430
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "System Time:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "System Date:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   2430
      Width           =   1185
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "Login.frx":0714
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To begin Select a Username"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1005
      TabIndex        =   3
      Top             =   195
      Width           =   2730
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Login.frx":13DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboUser_Click()
    txtPassword.SetFocus
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub Command1_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF6 Then
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Call forminit
End Sub
Sub forminit()
    Call CenterForm(frmLogin)
    'Me.Top = 1700
    'Me.Left = 5000
    Call loadcbouser
End Sub
Sub loadcbouser()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblUserInfo"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboUser.AddItem !sUserName
            cboUser.ItemData(cboUser.NewIndex) = CLng(!iUserID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub
Private Sub cmdOK_Click()
    If cboUser = "" Then cboUser.SetFocus: Exit Sub
    If txtPassword = "" Then txtPassword.SetFocus: Exit Sub
    
    Dim rsUsers As Recordset
    Static attempt As Integer
    
    Set rsUsers = New ADODB.Recordset
    'flgFirstUse = 2
    With rsUsers
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tbluserinfo WHERE sUserName='" & cboUser & "' AND sUserPassword='" & txtPassword & "'", cn, adOpenDynamic, adLockOptimistic
        If Not .EOF Then
            Unload Me
            If !sUserTaskLevel = "Administrator" Then
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                frmMain.Text3 = Trim(!sUserFirstName) & " " & Trim(!sUserMi) & " " & Trim(!sUserLastName)
                'frmMain.Frame1.Visible = True
                '
            ElseIf !sUserTaskLevel = "Secretary" Then
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                frmMain.Text3 = Trim(!sUserFirstName) & " " & Trim(!sUserMi) & " " & Trim(!sUserLastName)

                '
            Else
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                frmMain.Text3 = Trim(!sUserFirstName) & " " & Trim(!sUserMi) & " " & Trim(!sUserLastName)

                
            End If
        Else
            attempt = attempt + 1
            MsgBox "A C C E S S   D E N I E D " & vbCrLf & _
            "Please type the correct password", vbCritical, "This is your " & attempt & " attemp"
            Call Highlight(txtPassword)
            If attempt = 3 Then
                MsgBox "This will terminate the applicatin", vbCritical, "You already used all attempt"
                End
            End If
        End If
    End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cboUser.ListIndex = -1
        txtPassword = ""
        cmdCancel.SetFocus
    End If
End Sub


Private Sub Timer1_Timer()
    Label6.Caption = Format(Date, "mmmm dd, yyyy")
    Label7.Caption = Format(Time, "Medium Time")
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub
