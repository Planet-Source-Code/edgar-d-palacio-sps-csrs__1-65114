VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About the System"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4965
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      HideSelection   =   0   'False
      Left            =   1725
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "About.frx":058A
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   -105
      TabIndex        =   2
      Top             =   3120
      Width           =   8355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is licensed to :"
      Height          =   345
      Left            =   1725
      TabIndex        =   8
      Top             =   1365
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0 Copyright 2006"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   150
      TabIndex        =   7
      Top             =   3300
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1725
      TabIndex        =   5
      Top             =   1890
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "St. Paul's School of Ormoc Foundation, Inc."
      Height          =   345
      Left            =   1725
      TabIndex        =   4
      Top             =   1575
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Designed and Developed by:Edgar D. Palacio. You can contact me at 09165296620, email:greadjen@yahoo.com"
      Height          =   615
      Left            =   1725
      TabIndex        =   3
      Top             =   675
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SPS -Computerized School Registration Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   1725
      TabIndex        =   1
      Top             =   195
      Width           =   4845
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   135
      Picture         =   "About.frx":0677
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image3 
      Height          =   915
      Left            =   135
      Picture         =   "About.frx":125E
      Stretch         =   -1  'True
      Top             =   345
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   135
      Picture         =   "About.frx":27A1
      Stretch         =   -1  'True
      Top             =   345
      Width           =   1050
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
'center the form
Private Sub Form_Activate()
    Call CenterForm(frmAbout)
    Timer1_Timer
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
    If Image2.Visible = True Then
        Image2.Visible = False
        Image3.Visible = True
        Image1.Visible = False
    ElseIf Image3.Visible = True Then
        Image3.Visible = False
        Image1.Visible = True
        Image2.Visible = False
    Else
        Image1.Visible = False
        Image2.Visible = True
        Image3.Visible = False
    End If
End Sub
