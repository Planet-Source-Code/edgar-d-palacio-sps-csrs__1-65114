VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEnrollRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "EnrollRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Sort Option"
      Height          =   810
      Left            =   6780
      TabIndex        =   9
      Top             =   1455
      Width           =   1755
      Begin VB.OptionButton optNameSort 
         Caption         =   "Name Sort"
         Height          =   285
         Left            =   195
         TabIndex        =   11
         Top             =   465
         Width           =   1455
      End
      Begin VB.OptionButton optGenderSort 
         Caption         =   "Gender Sort"
         Height          =   285
         Left            =   195
         TabIndex        =   10
         Top             =   225
         Value           =   -1  'True
         Width           =   1470
      End
   End
   Begin VB.TextBox txtDummyEnrollmentID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4440
      TabIndex        =   8
      Top             =   1710
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Option"
      Height          =   810
      Left            =   4980
      TabIndex        =   5
      Top             =   1455
      Width           =   1755
      Begin VB.OptionButton optLastName 
         Caption         =   "Last Name"
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "Level"
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   465
         Width           =   1245
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   975
      TabIndex        =   3
      Top             =   1710
      Width           =   3360
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2640
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
            Picture         =   "EnrollRecord.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvStudentList 
      Height          =   5055
      Left            =   150
      TabIndex        =   1
      Top             =   2310
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8916
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "School Year"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Enrolled"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Age"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Payment Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Upon Enrollment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "PTA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remainnig"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "O.R. Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1935
      Top             =   915
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EnrollRecord.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EnrollRecord.frx":0ABE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   6990
      TabIndex        =   2
      Top             =   810
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Assign section"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "assign"
            Object.ToolTipText     =   "Back to enrollment form"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1770
      Width           =   675
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "EnrollRecord.frx":0E58
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment Record"
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
      TabIndex        =   0
      Top             =   150
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "EnrollRecord.frx":1B22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmEnrollRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dummySelect As Integer
Private Sub Form_Activate()
    Call forminit
    optGenderSort.Value = False
    optNameSort.Value = False
End Sub

Sub forminit()
    Call loadStudentList
    Call PositionForm3(frmEnrollRecord)
End Sub

Private Sub lsvStudentList_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    
    frmEnroll.Toolbar1.Buttons(2).Enabled = False
    frmEnroll.Toolbar1.Buttons(4).Enabled = True
    frmEnroll.Toolbar1.Buttons(5).Enabled = True
    row = lsvStudentList.SelectedItem.Index
    dummySelect = lsvStudentList.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qryEnrollment "
    strSQL = strSQL & "WHERE iEnrollmentID=" & dummySelect
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
        txtDummyEnrollmentID = dummySelect
    With frmEnroll
        .lsvStudents.Visible = False
        .txtEnrollmentIDDummy = rs1!iEnrollmentID
        .txtStudentNameDummy = rs1!istudentID
        .txtLastName = rs1!sStudentLastName
        .txtFirstName = rs1!sStudentFirstName
        .txtMiddleName = rs1!sStudentMiddlename
        .txtAge = rs1!sStudentAge
        .txtSex = rs1!sStudentSex
        .txtSchool = rs1!sSchoolLastAttended
        .dtpDateEnrolled = rs1!dtEnrolled
        .cboSchoolYear.ListIndex = ListFindItem(.cboSchoolYear, CLng(rs1!iSchoolYearID))
        .cboLevel.ListIndex = ListFindItem(.cboLevel, CLng(rs1!iLevelID))
        If rs1!sPaymentType = "Full" Then
            .optFullPayment.Value = True
        ElseIf rs1!sPaymentType = "Partial" Then
            .optPartialPayment.Value = True
        End If
        .txtUponEnrollment = rs1!cUponEnrollment
        .txtPTA = rs1!cPTA
        .txtTotal = rs1!cTotal
        .txtCashAmount = rs1!cCashAmount
        .txtChange = rs1!cChange
        .txtRemaining = rs1!cRemaining
        .txtOR = rs1!sORNumber
        .lblTotal = rs1!cAccount
        .txtDummyTotal = rs1!cAccount
    End With
    With frmAssign
        .txtEnrollmentID = rs1!istudentID
        .txtLevelIDDummy = rs1!iLevelID
        .lblLastName = rs1!sStudentLastName
        .lblFirstName = rs1!sStudentFirstName
        .lblMiddleName = rs1!sStudentMiddlename
        .lblLevel = rs1!sLevelName
        .txtGender = rs1!sStudentSex
        .txtSchoolYear = rs1!sSchoolYearName
    End With
    Set rs1 = Nothing
End Sub



Private Sub optGenderSort_Click()
    If optGenderSort.Value = True Then
        lsvStudentList.Sorted = True
        lsvStudentList.SortKey = 4
        lsvStudentList.SortOrder = lvwDescending
     End If
End Sub


Private Sub optNameSort_Click()
    If optNameSort.Value = True Then
        lsvStudentList.Sorted = True
        lsvStudentList.SortKey = 3
        lsvStudentList.SortOrder = lvwAscending
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
        Case "assign"
            lsvStudentList_DblClick
            Load frmAssign
            frmAssign.Show 1
        Case "back"
            Unload Me
    End Select
End Sub
Sub loadStudentList()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryEnrollment"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvStudentList.ListItems.Clear
    
    If rs.EOF Then
        lsvStudentList.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvStudentList.ListItems.Add(, , !iEnrollmentID, , 1)
                X.SubItems(1) = !sSchoolYearName
                X.SubItems(2) = !dtEnrolled
                X.SubItems(3) = !Name
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sStudentAge
                X.SubItems(6) = !sLevelName
                X.SubItems(7) = !sPaymentType
                X.SubItems(8) = !cUponEnrollment
                X.SubItems(9) = !cPTA
                X.SubItems(10) = !cRemaining
                X.SubItems(11) = !sORNumber
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub txtSearch_Change()
    Dim X
    Dim strSQL As String
    
    
    
    If optLastName.Value = True Then
        strSQL = "SELECT * FROM qryEnrollment "
        strSQL = strSQL & "WHERE Name LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudentList.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvStudentList.ListItems.Add(, , !iEnrollmentID, , 1)
               X.SubItems(1) = !sSchoolYearName
                X.SubItems(2) = !dtEnrolled
                X.SubItems(3) = !Name
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sStudentAge
                X.SubItems(6) = !sLevelName
                X.SubItems(7) = !sPaymentType
                X.SubItems(8) = !cUponEnrollment
                X.SubItems(9) = !cPTA
                X.SubItems(10) = !cRemaining
                X.SubItems(11) = !sORNumber
                .MoveNext
            Loop
        End With
        Set rs = Nothing
    ElseIf optLevel.Value = True Then
        
        Dim rsviewLevel As New ADODB.Recordset
        
        strSQL = "SELECT * FROM qryEnrollment "
        strSQL = strSQL & "WHERE sLevelName LIKE'" & txtSearch.Text & "%'"
        
        rsviewLevel.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
            lsvStudentList.ListItems.Clear
            With rsviewLevel
                Do While Not rsviewLevel.EOF
                Set X = lsvStudentList.ListItems.Add(, , !iEnrollmentID, , 1)
                X.SubItems(1) = !sSchoolYearName
                X.SubItems(2) = !dtEnrolled
                X.SubItems(3) = !Name
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sStudentAge
                X.SubItems(6) = !sLevelName
                X.SubItems(7) = !cUponEnrollment
                X.SubItems(8) = !cRemaining
                X.SubItems(9) = !sORNumber
                        .MoveNext
                Loop
            End With
            Set rsviewLevel = Nothing
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvStudentList.ListItems.Count <> 0 Then
            lsvStudentList.SetFocus
        End If
    End If
End Sub
