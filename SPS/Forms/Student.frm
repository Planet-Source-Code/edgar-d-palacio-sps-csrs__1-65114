VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Student.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStudentDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   435
      TabIndex        =   43
      Top             =   870
      Visible         =   0   'False
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   1395
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Basic Information"
      TabPicture(0)   =   "Student.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtLastName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFirstName"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtMiddleName"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboSex"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboNationality"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboReligion"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtAge"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtBirthPlace"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtContactNumber"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtAddress"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtLastSchoolAttended"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dtpBirthDate"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Family Background"
      TabPicture(1)   =   "Student.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label18"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label20"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtFather"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtFatherOccupation"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtMother"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtMotherOccupation"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtGuardian"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtGuardianOccupation"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtRelationship"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtGuardianContactNumber"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin MSComCtl2.DTPicker dtpBirthDate 
         Height          =   315
         Left            =   1740
         TabIndex        =   44
         Top             =   1920
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   38821
      End
      Begin VB.TextBox txtLastSchoolAttended 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1755
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   4575
         Width           =   3945
      End
      Begin VB.TextBox txtGuardianContactNumber 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73320
         TabIndex        =   17
         Top             =   2850
         Width           =   3450
      End
      Begin VB.TextBox txtRelationship 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73320
         TabIndex        =   16
         Top             =   2505
         Width           =   3450
      End
      Begin VB.TextBox txtGuardianOccupation 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73320
         TabIndex        =   18
         Top             =   3180
         Width           =   4965
      End
      Begin VB.TextBox txtGuardian 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73320
         TabIndex        =   15
         Top             =   2175
         Width           =   3450
      End
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   -74865
         TabIndex        =   36
         Top             =   1920
         Width           =   6525
      End
      Begin VB.TextBox txtMotherOccupation 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   14
         Top             =   1545
         Width           =   4965
      End
      Begin VB.TextBox txtMother 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   13
         Top             =   1215
         Width           =   3450
      End
      Begin VB.TextBox txtFatherOccupation 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   12
         Top             =   885
         Width           =   4965
      End
      Begin VB.TextBox txtFather 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73350
         TabIndex        =   11
         Top             =   555
         Width           =   3450
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   1755
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4050
         Width           =   4890
      End
      Begin VB.TextBox txtContactNumber 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   8
         Top             =   3699
         Width           =   2355
      End
      Begin VB.TextBox txtBirthPlace 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   5
         Top             =   2631
         Width           =   4665
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   4
         Top             =   2285
         Width           =   495
      End
      Begin VB.ComboBox cboReligion 
         Height          =   315
         ItemData        =   "Student.frx":03C2
         Left            =   1755
         List            =   "Student.frx":03C9
         TabIndex        =   7
         Top             =   3338
         Width           =   2625
      End
      Begin VB.ComboBox cboNationality 
         Height          =   315
         ItemData        =   "Student.frx":03DD
         Left            =   1755
         List            =   "Student.frx":03E4
         TabIndex        =   6
         Top             =   2977
         Width           =   2625
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "Student.frx":03F2
         Left            =   1770
         List            =   "Student.frx":03FC
         TabIndex        =   3
         Top             =   1563
         Width           =   2040
      End
      Begin VB.TextBox txtMiddleName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   2
         Top             =   1217
         Width           =   3825
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   1
         Top             =   871
         Width           =   3825
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1755
         TabIndex        =   0
         Top             =   525
         Width           =   3825
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "School Last Attended:"
         Height          =   405
         Left            =   555
         TabIndex        =   42
         Top             =   4575
         Width           =   1065
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Contact Number:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   40
         Top             =   2895
         Width           =   1470
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Relationship:"
         Height          =   195
         Left            =   -74520
         TabIndex        =   39
         Top             =   2565
         Width           =   1110
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
         Height          =   195
         Left            =   -74430
         TabIndex        =   38
         Top             =   3240
         Width           =   1020
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Guardian:"
         Height          =   195
         Left            =   -74265
         TabIndex        =   37
         Top             =   2235
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
         Height          =   195
         Left            =   -74430
         TabIndex        =   35
         Top             =   1605
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Mother's Name:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   34
         Top             =   1275
         Width           =   1350
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
         Height          =   195
         Left            =   -74430
         TabIndex        =   33
         Top             =   945
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Father's Name:"
         Height          =   195
         Left            =   -74715
         TabIndex        =   32
         Top             =   615
         Width           =   1305
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   855
         TabIndex        =   31
         Top             =   4103
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Contact Number:"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   3752
         Width           =   1470
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Birth Place:"
         Height          =   195
         Left            =   630
         TabIndex        =   29
         Top             =   2684
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   195
         Left            =   1215
         TabIndex        =   28
         Top             =   2338
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         Height          =   195
         Left            =   675
         TabIndex        =   27
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Religion:"
         Height          =   195
         Left            =   870
         TabIndex        =   26
         Top             =   3398
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nationality:"
         Height          =   195
         Left            =   645
         TabIndex        =   25
         Top             =   3037
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         Height          =   195
         Left            =   1215
         TabIndex        =   24
         Top             =   1623
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
         Height          =   195
         Left            =   450
         TabIndex        =   23
         Top             =   1270
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   630
         TabIndex        =   22
         Top             =   924
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   645
         TabIndex        =   21
         Top             =   578
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4350
      Top             =   2220
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
            Picture         =   "Student.frx":040E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Student.frx":07A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   5625
      TabIndex        =   41
      Top             =   780
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Student.frx":0B42
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Personal Info"
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
      TabIndex        =   19
      Top             =   150
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Student.frx":180C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6990
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FindItem As Boolean

Private Sub Form_Activate()
    Call forminit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Sub forminit()
    SSTab1.Tab = 0
    txtLastName.SetFocus
End Sub

Private Sub Form_Load()
    Call PositionForm2(frmStudent)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
            
            
            If frmStudentRecord.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblStudents"
                
                
                If complete = True Then
                rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs
                        .AddNew
                        !sStudentLastName = txtLastName
                        !sStudentFirstName = txtFirstName
                        !sStudentMiddlename = txtMiddleName
                        !sStudentSex = cboSex
                        !dtStudentBirthDate = dtpBirthDate.Value
                        !sStudentAge = txtAge
                        !sStudentBirthPlace = txtBirthPlace
                        !sStudentNationality = cboNationality
                        !sStudentReligion = cboReligion
                        !sStudentContactNumber = txtContactNumber
                        !sStudentAddress = txtAddress
                        !sSchoolLastAttended = txtLastSchoolAttended
                        !sFatherName = txtFather
                        !sFatherOccupation = txtFatherOccupation
                        !sMotherName = txtMother
                        !sMotherOccupation = txtMotherOccupation
                        !sGuardian = txtGuardian
                        !sGuardianRelationship = txtRelationship
                        !sGuardianOccupation = txtGuardianOccupation
                        !sGuardianNumber = txtGuardianContactNumber
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add New Student Record"
                    Unload Me
                    Set rs = Nothing
                Else
                    If txtLastName = "" Then
                        MsgBox "Please enter LAST NAME", vbExclamation + vbOKOnly, "Empty Last Name"
                        txtLastName.SetFocus
                        Exit Sub
                    ElseIf txtFirstName = "" Then
                        MsgBox "Please enter FIRST NAME", vbExclamation + vbOKOnly, "Empty First Name"
                        txtFirstName.SetFocus
                        Exit Sub
                    ElseIf txtMiddleName = "" Then
                        MsgBox "Please enter MIDDLE NAME", vbExclamation + vbOKOnly, "Empty Middle Name"
                        txtMiddleName.SetFocus
                        Exit Sub
                    ElseIf cboSex = "" Then
                        MsgBox "Please select SEX/GENDER", vbExclamation + vbOKOnly, "Empty Sex"
                        cboSex.SetFocus
                        Exit Sub
                    ElseIf dtpBirthDate = "" Then
                        MsgBox "Please specify BIRTH DATE", vbExclamation + vbOKOnly, "Empty Birth Date"
                        dtpBirthDate.SetFocus
                        Exit Sub
                    ElseIf txtContactNumber = "" Then
                        MsgBox "Please enter CONTACT NUMBER", vbExclamation + vbOKOnly, "Empty Contact Number"
                        txtContactNumber.SetFocus
                        Exit Sub
                    ElseIf txtAddress = "" Then
                        MsgBox "Please enter ADDRESS", vbExclamation + vbOKOnly, "Empty Address"
                        txtAddress.SetFocus
                        Exit Sub
                    End If
                End If
                
            ElseIf frmStudentRecord.dummyButton = "edit" Then
                strSQL = "SELECT * FROM tblStudents"
                strSQL = strSQL & " WHERE iStudentID=" & txtStudentDummy
                
                If complete = True Then
                rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs3
                        !sStudentLastName = txtLastName
                        !sStudentFirstName = txtFirstName
                        !sStudentMiddlename = txtMiddleName
                        !sStudentSex = cboSex
                        !dtStudentBirthDate = dtpBirthDate.Value
                        !sStudentAge = txtAge
                        !sStudentBirthPlace = txtBirthPlace
                        !sStudentNationality = cboNationality
                        !sStudentReligion = cboReligion
                        !sStudentContactNumber = txtContactNumber
                        !sStudentAddress = txtAddress
                        !sSchoolLastAttended = txtLastSchoolAttended
                        !sFatherName = txtFather
                        !sFatherOccupation = txtFatherOccupation
                        !sMotherName = txtMother
                        !sMotherOccupation = txtMotherOccupation
                        !sGuardian = txtGuardian
                        !sGuardianRelationship = txtRelationship
                        !sGuardianOccupation = txtGuardianOccupation
                        !sGuardianNumber = txtGuardianContactNumber
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update Student Record"
                    Unload Me
                    Set rs3 = Nothing
                Else
                    If txtLastName = "" Then
                        MsgBox "Please enter LAST NAME", vbExclamation + vbOKOnly, "Empty Last Name"
                        txtLastName.SetFocus
                        Exit Sub
                    ElseIf txtFirstName = "" Then
                        MsgBox "Please enter FIRST NAME", vbExclamation + vbOKOnly, "Empty First Name"
                        txtFirstName.SetFocus
                        Exit Sub
                    ElseIf txtMiddleName = "" Then
                        MsgBox "Please enter MIDDLE NAME", vbExclamation + vbOKOnly, "Empty Middle Name"
                        txtMiddleName.SetFocus
                        Exit Sub
                    ElseIf cboSex = "" Then
                        MsgBox "Please select SEX/GENDER", vbExclamation + vbOKOnly, "Empty Sex"
                        cboSex.SetFocus
                        Exit Sub
                    ElseIf dtpBirthDate = "" Then
                        MsgBox "Please specify BIRTH DATE", vbExclamation + vbOKOnly, "Empty Birth Date"
                        dtpBirthDate.SetFocus
                        Exit Sub
                    ElseIf txtContactNumber = "" Then
                        MsgBox "Please enter CONTACT NUMBER", vbExclamation + vbOKOnly, "Empty Contact Number"
                        txtContactNumber.SetFocus
                        Exit Sub
                    ElseIf txtAddress = "" Then
                        MsgBox "Please enter ADDRESS", vbExclamation + vbOKOnly, "Empty Address"
                        txtAddress.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case "back"
            Unload Me
    End Select
End Sub

Function complete()
    If txtLastName = "" Or txtFirstName = "" Or txtMiddleName = "" Or cboSex = "" _
        Or dtpBirthDate = "" Or txtAge = "" Or txtContactNumber = "" Or txtAddress = "" Then
        complete = False
    Else
        complete = True
    End If
End Function
