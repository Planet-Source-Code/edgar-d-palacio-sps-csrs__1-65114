VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Level.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1410
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
            Picture         =   "Level.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Level.frx":0726
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Level.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Level.frx":0E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Level.frx":11F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   2640
      TabIndex        =   0
      Top             =   795
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add new level"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "edit"
            Object.ToolTipText     =   "Edit existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete existing record"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Find specific record"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3885
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   6853
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Level.frx":158E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvLevel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   442
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsvLevel 
         Height          =   2805
         Left            =   135
         TabIndex        =   3
         Top             =   945
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   4948
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Level"
            Object.Width           =   3316
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Find Gr./Yr. Level:"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   495
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Level"
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
      Left            =   735
      TabIndex        =   5
      Top             =   150
      Width           =   1470
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   75
      Picture         =   "Level.frx":15AA
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   15
      Picture         =   "Level.frx":2274
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummyButton As String
Public dummySelect As Integer

Private Sub Form_Activate()
    Call loadlsvLevel
    Call forminit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtSearch = ""
        Label2.Enabled = False
        txtSearch.Enabled = False
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call PositionForm(frmLevel)
End Sub
Sub forminit()
    
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub lsvLevel_Click()
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
End Sub

Private Sub lsvLevel_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    dummyButton = "edit"
    row = lsvLevel.SelectedItem.Index
    dummySelect = lsvLevel.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblLevels "
    strSQL = strSQL & "WHERE iLevelID=" & dummySelect
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With frmLevelUpdate
        .txtLevelDummy = dummySelect
        .txtLevel = rs!sLevelName
        .txtDescription = rs!sLevelNotes
    End With
    Load frmLevelUpdate
    frmLevelUpdate.Show 1
    Set rs = Nothing
End Sub
Sub loadlsvLevel()
    Dim X As Integer
    Dim strSQL As String
    strSQL = "SELECT * FROM tblLevels"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvLevel.ListItems.Clear

    While Not rs.EOF
        Set lst = lsvLevel.ListItems.Add(, , rs(0), , 1)
        For X = 1 To 2
            lst.SubItems(X) = rs(X)
        Next X
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

Private Sub lsvLevel_KeyPress(KeyAscii As Integer)
    lsvLevel_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            dummyButton = "add"
            Load frmLevelUpdate
            frmLevelUpdate.Show 1
        Case "edit"
            lsvLevel_DblClick
        Case "delete"
            If MsgBox("Do you want to DELETE " & lsvLevel.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete Level") = vbYes Then
                cn.Execute "DELETE FROM tblLevels WHERE iLevelID=" & lsvLevel.SelectedItem.Text
                Call loadlsvLevel
            End If
        Case "search"
            Label2.Enabled = True
            txtSearch.Enabled = True
            txtSearch.SetFocus
    End Select
        
End Sub

Private Sub txtSearch_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblLevels "
    strSQL = strSQL & "WHERE sLevelName LIKE'" & txtSearch.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvLevel.ListItems.Clear
    With rs
        Do While Not rs.EOF
            Set X = lsvLevel.ListItems.Add(, , !iLevelID, , 1)
                X.SubItems(1) = !sLevelName
                X.SubItems(2) = !sLevelNotes
                .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvLevel.ListItems.Count <> 0 Then
            lsvLevel.SetFocus
        End If
    End If
End Sub

