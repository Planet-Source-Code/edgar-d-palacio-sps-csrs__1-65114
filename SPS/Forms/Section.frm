VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSection 
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
   Icon            =   "Section.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1755
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
            Picture         =   "Section.frx":038A
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
            Picture         =   "Section.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Section.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Section.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Section.frx":11F2
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
            Object.ToolTipText     =   "Add new section"
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
      TabPicture(0)   =   "Section.frx":158C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvSection"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1350
         TabIndex        =   2
         Top             =   442
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsvSection 
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Section"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Allowed"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Min. Ave."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Max. Ave."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Notes"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Find Section:"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   495
         Width           =   1110
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   75
      Picture         =   "Section.frx":15A8
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Section"
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
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   15
      Picture         =   "Section.frx":2272
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummyButton As String
Public dummySelect As Integer
Public FindItem As Boolean

Private Sub Form_Activate()
    Call loadlsvSection
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
    Call PositionForm(frmSection)
End Sub
Sub forminit()
    
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub lsvSection_Click()
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
End Sub

Private Sub lsvSection_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    dummyButton = "edit"
    row = lsvSection.SelectedItem.Index
    dummySelect = lsvSection.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblSections "
    strSQL = strSQL & "WHERE iSectionID=" & dummySelect
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With frmSectionUpdate
        .txtSectionDummy = dummySelect
        .cboLevel.ListIndex = ListFindItem(.cboLevel, CLng(rs!iLevelID))
        .txtSection = rs!sSectionName
        .txtAllowed = rs!iAllowed
        .txtMinAverage = rs!iMaxAverage
        .txtMaxAverage = rs!iMinAverage
        .txtDescription = rs!sSectionNotes
    End With
    Load frmSectionUpdate
    frmSectionUpdate.Show 1
    Set rs = Nothing
End Sub
Sub loadlsvSection()
    Dim X
    Dim strSQL As String
    Dim rsSection As New ADODB.Recordset
   
    strSQL = "SELECT * FROM qrySections"
    
    rsSection.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvSection.ListItems.Clear
    If rsSection.EOF = True Then
        lsvSection.ListItems.Clear
    Else
        With rsSection
            Do While Not rsSection.EOF
                Set X = lsvSection.ListItems.Add(, , !iSectionID, , 1)
                X.SubItems(1) = !Section
                X.SubItems(2) = !iAllowed
                X.SubItems(3) = !iMinAverage
                X.SubItems(4) = !iMaxAverage
                X.SubItems(5) = !sSectionNotes
                .MoveNext
            Loop
        End With
        Set rsSection = Nothing
    End If
End Sub

Private Sub lsvSection_KeyPress(KeyAscii As Integer)
    lsvSection_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            dummyButton = "add"
            Load frmSectionUpdate
            frmSectionUpdate.Show 1
        Case "edit"
            lsvSection_DblClick
        Case "delete"
            If MsgBox("Do you want to DELETE Section " & lsvSection.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete Section") = vbYes Then
                cn.Execute "DELETE FROM tblSections WHERE iSectionID=" & lsvSection.SelectedItem.Text
                Call loadlsvSection
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
    strSQL = "SELECT * FROM qrySections "
    strSQL = strSQL & "WHERE Section LIKE'" & txtSearch.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvSection.ListItems.Clear
    With rs
        Do While Not rs.EOF
            Set X = lsvSection.ListItems.Add(, , !iSectionID, , 1)
                X.SubItems(1) = !Section
                X.SubItems(2) = !iAllowed
                X.SubItems(3) = !iMinAverage
                X.SubItems(4) = !iMaxAverage
                X.SubItems(5) = !sSectionNotes
                .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvSection.ListItems.Count <> 0 Then
            lsvSection.SetFocus
        End If
    End If
End Sub

Function ListFindItem(lstCtrl As Control, lngSearch As Long) As Integer
   'just returns the position, does not set it
   'used to see if item is in list
   Dim intLen As Integer
   Dim intLoop As Integer
   Dim intPos As Integer

   intLen = lstCtrl.ListCount - 1
   intPos = -1
   For intLoop = 0 To intLen
      If lstCtrl.ItemData(intLoop) = lngSearch Then
         intPos = intLoop
         Exit For
      End If
   Next intLoop
   ListFindItem = intPos
End Function
