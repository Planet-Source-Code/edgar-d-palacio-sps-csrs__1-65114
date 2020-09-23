VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Fees.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1215
      Top             =   885
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
            Picture         =   "Fees.frx":038A
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
            Picture         =   "Fees.frx":0726
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fees.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fees.frx":0E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Fees.frx":11F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   4035
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
            Object.ToolTipText     =   "Add new school fee"
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
      Height          =   4080
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   7197
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Fees.frx":158E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvFees"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "Search Option"
         Height          =   810
         Left            =   3765
         TabIndex        =   6
         Top             =   315
         Width           =   1725
         Begin VB.OptionButton optLevel 
            Caption         =   "Level"
            Height          =   285
            Left            =   195
            TabIndex        =   8
            Top             =   465
            Width           =   1455
         End
         Begin VB.OptionButton optSchoolFees 
            Caption         =   "School Fees"
            Height          =   285
            Left            =   195
            TabIndex        =   7
            Top             =   225
            Value           =   -1  'True
            Width           =   1470
         End
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   900
         TabIndex        =   2
         Top             =   585
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsvFees 
         Height          =   2760
         Left            =   120
         TabIndex        =   3
         Top             =   1170
         Width           =   6330
         _ExtentX        =   11165
         _ExtentY        =   4868
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
            Text            =   "Fee Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "School Year"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Level Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Search:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   645
         Width           =   675
      End
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   5700
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL FEES WHOLE YEAR:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2865
      TabIndex        =   9
      Top             =   5700
      Width           =   2505
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   75
      Picture         =   "Fees.frx":15AA
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of School Fees"
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
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   15
      Picture         =   "Fees.frx":2274
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummyButton As String
Public dummySelect As Integer
Dim total

Private Sub Form_Activate()
    Call forminit
    Call loadlsvFees
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
    Call PositionForm(frmFees)
End Sub
Sub forminit()
    
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub lsvSchoolYear_Click()
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
End Sub

Private Sub lsvFees_Click()
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
End Sub

Private Sub lsvFees_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    
    dummyButton = "edit"
    row = lsvFees.SelectedItem.Index
    dummySelect = lsvFees.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblFees "
    strSQL = strSQL & "WHERE iFeeID=" & dummySelect
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With frmFeesUpdate
        .txtFeesDummy = dummySelect
        .cboSchoolYear.ListIndex = ListFindItem(.cboSchoolYear, CLng(rs1!iSchoolYearID))
        .cboLevel.ListIndex = ListFindItem(.cboLevel, CLng(rs1!iLevelID))
        .txtFeeName = rs1!sFeeName
        .txtAmount = rs1!cAmount
        .txtDescription = rs1!sFeeNotes
    End With
    Load frmFeesUpdate
    frmFeesUpdate.Show 1
    Set rs1 = Nothing
End Sub
Sub loadlsvFees()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryFees"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvFees.ListItems.Clear
    
    If rs.EOF Then
        lsvFees.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvFees.ListItems.Add(, , !iFeeID, , 1)
                X.SubItems(1) = !sFeeName
                X.SubItems(2) = !sSchoolYearName
                X.SubItems(3) = !sLevelName
                X.SubItems(4) = !cAmount
                X.SubItems(5) = !sFeeNotes
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub lsvFees_KeyPress(KeyAscii As Integer)
    lsvFees_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            dummyButton = "add"
            Load frmFeesUpdate
            frmFeesUpdate.Show 1
        Case "edit"
            lsvFees_DblClick
        Case "delete"
            If MsgBox("Do you want to DELETE " & lsvFees.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete School Fee") = vbYes Then
                cn.Execute "DELETE FROM tblFees WHERE iFeeID=" & lsvFees.SelectedItem.Text
                Call loadlsvFees
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
    
    
    
    If optSchoolFees.Value = True Then
        strSQL = "SELECT * FROM qryFees "
        strSQL = strSQL & "WHERE sFeeName LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvFees.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvFees.ListItems.Add(, , !iFeeID, , 1)
                    X.SubItems(1) = !sFeeName
                    X.SubItems(2) = !sSchoolYearName
                    X.SubItems(3) = !sLevelName
                    X.SubItems(4) = !cAmount
                    X.SubItems(5) = !sFeeNotes
                    .MoveNext
            Loop
        End With
        Set rs = Nothing
    ElseIf optLevel.Value = True Then
        
        Dim rsviewLevel As New ADODB.Recordset
        
        strSQL = "SELECT * FROM qryFees "
        strSQL = strSQL & "WHERE sLevelName LIKE'" & txtSearch.Text & "%'"
        
        rsviewLevel.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        If txtSearch = "" Then
            Call loadlsvFees
            lblTotal.Caption = "0.00"
        Else
            lsvFees.ListItems.Clear
            total = 0
            With rsviewLevel
                Do While Not rsviewLevel.EOF
                    Set X = lsvFees.ListItems.Add(, , !iFeeID, , 1)
                        X.SubItems(1) = !sFeeName
                        X.SubItems(2) = !sSchoolYearName
                        X.SubItems(3) = !sLevelName
                        X.SubItems(4) = !cAmount
                        X.SubItems(5) = !sFeeNotes
                        total = total + !cAmount
                        .MoveNext
                Loop
            End With
            Set rsviewLevel = Nothing
            lblTotal = Format(total, "P###,###.00")
        End If
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvFees.ListItems.Count <> 0 Then
            lsvFees.SetFocus
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

