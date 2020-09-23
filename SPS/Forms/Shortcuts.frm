VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShortcuts 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   Icon            =   "Shortcuts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5865
      Top             =   105
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Shortcuts.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvShortcuts 
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5980
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shortcut Key"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   12347
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   75
      Picture         =   "Shortcuts.frx":0724
      Stretch         =   -1  'True
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Keyboard Shortcuts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   150
      Width           =   3345
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Shortcuts.frx":0AAE
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Shortcuts.frx":1778
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call loadlsvshortcuts
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Sub loadlsvshortcuts()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblKeyboard"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvShortcuts.ListItems.Clear
    
    If rs.EOF Then
        lsvShortcuts.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvShortcuts.ListItems.Add(, , !iKeyboardID, , 1)
                X.SubItems(1) = !sKeyboardName
                X.SubItems(2) = !sDescription
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub Timer1_Timer()
    If Image2.Visible = False Then
        Image2.Visible = True
        imgIcon.Visible = False
    ElseIf imgIcon.Visible = False Then
        imgIcon.Visible = True
        Image2.Visible = False
    End If
End Sub
