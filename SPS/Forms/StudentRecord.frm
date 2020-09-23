VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStudentRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "StudentRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2655
      Top             =   1035
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
            Picture         =   "StudentRecord.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvStudent 
      Height          =   5970
      Left            =   150
      TabIndex        =   4
      Top             =   2400
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   10530
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Birth Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Age"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Father"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Occupation"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Mother"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Occupation"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Religion"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Nationality"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Guardian"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Relationship"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Occupation"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Contact #"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   150
      TabIndex        =   1
      Top             =   1575
      Width           =   8385
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Search:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   300
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1950
      Top             =   1005
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
            Picture         =   "StudentRecord.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StudentRecord.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StudentRecord.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "StudentRecord.frx":11F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   5865
      TabIndex        =   5
      Top             =   900
      Width           =   2775
      _ExtentX        =   4895
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
            Object.ToolTipText     =   "Add new record"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Record"
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
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "StudentRecord.frx":158C
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "StudentRecord.frx":2256
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmStudentRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummyButton As String
Public dummySelect As Integer
Dim total

Private Sub Form_Activate()
    Call loadlsvStudent
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
    Call PositionForm(frmStudentRecord)
End Sub
Sub forminit()
    
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub
Private Sub lsvStudent_Click()
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
End Sub
Private Sub lsvStudent_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    
    dummyButton = "edit"
    row = lsvStudent.SelectedItem.Index
    dummySelect = lsvStudent.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblStudents "
    strSQL = strSQL & "WHERE iStudentID=" & dummySelect
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With frmStudent
        .txtStudentDummy = dummySelect
        .txtLastName = rs1!sStudentLastName
        .txtFirstName = rs1!sStudentFirstName
        .txtMiddleName = rs1!sStudentMiddlename
        .cboSex = rs1!sStudentSex
        .dtpBirthDate = rs1!dtStudentBirthDate
        .txtAge = rs1!sStudentAge
        .txtBirthPlace = rs1!sStudentBirthPlace
        .cboNationality = rs1!sStudentNationality
        .cboReligion = rs1!sStudentReligion
        .txtContactNumber = rs1!sStudentContactNumber
        .txtAddress = rs1!sStudentAddress
        .txtLastSchoolAttended = rs1!sSchoolLastAttended
        .txtFather = rs1!sFatherName
        .txtFatherOccupation = rs1!sFatherOccupation
        .txtMother = rs1!sMotherName
        .txtMotherOccupation = rs1!sMotherOccupation
        .txtGuardian = rs1!sGuardian
        .txtRelationship = rs1!sGuardianRelationship
        .txtGuardianContactNumber = rs1!sGuardianNumber
        .txtGuardianOccupation = rs1!sGuardianOccupation
    End With
    Load frmStudent
    frmStudent.Show 1
    Set rs1 = Nothing
End Sub
Sub loadlsvStudent()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblStudents"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvStudent.ListItems.Clear
    
    If rs.EOF Then
        lsvStudent.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvStudent.ListItems.Add(, , !istudentID, , 1)
                X.SubItems(1) = Trim(!sStudentLastName) + ", " + Trim(!sStudentFirstName) + " " + Trim(!sStudentMiddlename)
                X.SubItems(2) = !sStudentSex
                X.SubItems(3) = !dtStudentBirthDate
                X.SubItems(4) = !sStudentAge
                X.SubItems(5) = !sStudentContactNumber
                X.SubItems(6) = !sStudentAddress
                X.SubItems(7) = !sFatherName
                X.SubItems(8) = !sFatherOccupation
                X.SubItems(9) = !sMotherName
                X.SubItems(10) = !sMotherOccupation
                X.SubItems(11) = !sStudentReligion
                X.SubItems(12) = !sStudentNationality
                X.SubItems(13) = !sGuardian
                X.SubItems(14) = !sGuardianRelationship
                X.SubItems(15) = !sGuardianOccupation
                X.SubItems(16) = !sGuardianNumber
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub lsvStudent_KeyPress(KeyAscii As Integer)
    lsvStudent_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            dummyButton = "add"
            Load frmStudent
            frmStudent.Show 1
        Case "edit"
            lsvStudent_DblClick
        Case "delete"
            If MsgBox("Do you want to DELETE " & lsvStudent.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete Student Record") = vbYes Then
                cn.Execute "DELETE FROM tblStudents WHERE iStudentID=" & lsvStudent.SelectedItem.Text
                Call loadlsvStudent
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
    
        strSQL = "SELECT * FROM tblStudents "
        strSQL = strSQL & "WHERE sStudentLastName LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudent.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvStudent.ListItems.Add(, , !istudentID, , 1)
                X.SubItems(1) = Trim(!sStudentLastName) + ", " + Trim(!sStudentFirstName) + " " + Trim(!sStudentMiddlename)
                X.SubItems(2) = !sStudentSex
                X.SubItems(3) = !dtStudentBirthDate
                X.SubItems(4) = !sStudentAge
                X.SubItems(5) = !sStudentContactNumber
                X.SubItems(6) = !sStudentAddress
                X.SubItems(7) = !sFatherName
                X.SubItems(8) = !sFatherOccupation
                X.SubItems(9) = !sMotherName
                X.SubItems(10) = !sMotherOccupation
                X.SubItems(11) = !sStudentReligion
                X.SubItems(12) = !sStudentNationality
                X.SubItems(13) = !sGuardian
                X.SubItems(14) = !sGuardianRelationship
                X.SubItems(15) = !sGuardianOccupation
                X.SubItems(16) = !sGuardianNumber
                .MoveNext
            Loop
        End With
        Set rs = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvStudent.ListItems.Count <> 0 Then
            lsvStudent.SetFocus
        End If
    End If
End Sub



