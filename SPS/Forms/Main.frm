VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "St. Paul's School of Ormoc Foundation, Inc. - Computerized School Registration Software"
   ClientHeight    =   10740
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   12345
      TabIndex        =   4
      Top             =   705
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16744576
      Appearance      =   0
      StartOfWeek     =   20512769
      CurrentDate     =   38822
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   1440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2055
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1926
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":205A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "student"
            Object.ToolTipText     =   "Student Info"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "enroll"
            Object.ToolTipText     =   "Enrollment Info"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pay"
            Object.ToolTipText     =   "Payments Info"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "studentsearch"
            Object.ToolTipText     =   "View stuent info"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "enrollsearch"
            Object.ToolTipText     =   "View enrolled students"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "section"
            Object.ToolTipText     =   "Assign section"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lock"
            Object.ToolTipText     =   "System Lock"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "About SPS-CSRS"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   10425
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13573
            MinWidth        =   13123
            Text            =   "Developed by: Edgar D. Palacio"
            TextSave        =   "Developed by: Edgar D. Palacio"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   2990
            TextSave        =   "4/26/2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:58 AM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   165
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   -15
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   270
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   210
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   -45
      Picture         =   "Main.frx":25F4
      Stretch         =   -1  'True
      Top             =   3990
      Width           =   15450
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   -45
      Top             =   6840
      Width           =   15525
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   -45
      Top             =   3765
      Width           =   15525
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFNewStudent 
         Caption         =   "&New Student"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFTeacher 
         Caption         =   "&New Teacher"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFMaintenance 
         Caption         =   "&Maintenance"
         Begin VB.Menu mnuFMSchoolYear 
            Caption         =   "School Year"
            Shortcut        =   ^Y
         End
         Begin VB.Menu mnuFMQuarter 
            Caption         =   "Quarter"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mnuFMLevels 
            Caption         =   "Levels"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFMSections 
            Caption         =   "Sections"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuFMSchoolFees 
            Caption         =   "School Fees"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuFMShortcuts 
            Caption         =   "Add Keyboard Shortcuts"
            Shortcut        =   ^K
         End
      End
      Begin VB.Menu mnuFsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFAdministratorOption 
         Caption         =   "&Administrator Option"
         Begin VB.Menu mnuFAModifyUsers 
            Caption         =   "Modify Users"
            Shortcut        =   ^U
         End
      End
      Begin VB.Menu mnuFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFLogOff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuTEnrollStudent 
         Caption         =   "Enroll Student"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTPayments 
         Caption         =   "Payments"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuVStudentInfo 
         Caption         =   "Student Personal Info"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuVTeacherInfo 
         Caption         =   "Teaher's Personal Info"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVEnrolledStudents 
         Caption         =   "Enrolled Students"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuVSections 
         Caption         =   "Sections"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuVSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVToolbar 
         Caption         =   "&Toolbar"
         Begin VB.Menu mnuVTDefault 
            Caption         =   "Default"
         End
         Begin VB.Menu mnuVTAlignLeft 
            Caption         =   "Align Left"
         End
         Begin VB.Menu mnuVTAlignRight 
            Caption         =   "Align Right"
         End
         Begin VB.Menu mnuVTAlignBottom 
            Caption         =   "Align Bottom"
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTCalculator 
         Caption         =   "Calculator"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuTSystemLock 
         Caption         =   "System Lock"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuKS 
         Caption         =   "&Display Keyboard Shortcuts"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i
Dim str1 As String

Private Sub Form_Load()
    Call DBConnect
    Call forminit
    str1 = "Program created by:  Edgar D. Palacio"
    i = 0
End Sub
Sub forminit()
    Me.Show
    Load frmLogin
    frmLogin.Show 1
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim answer
    answer = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Exit Computerize School Registration System")
    If answer = vbYes Then
        Call DBClose
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub mnuFAModifyUsers_Click()
    Load frmUserInfo
    frmUserInfo.Show 1
End Sub

Private Sub mnuFExit_Click()
    Unload Me
End Sub

Private Sub mnuFLogOff_Click()
    Load frmLogin
    frmLogin.Show 1
End Sub

Private Sub mnuFMLevels_Click()
    Load frmLevel
    frmLevel.Show 1
End Sub

Private Sub mnuFMQuarter_Click()
    Load frmQuarter
    frmQuarter.Show 1
End Sub

Private Sub mnuFMSchoolFees_Click()
    Load frmFees
    frmFees.Show 1
End Sub

Private Sub mnuFMSchoolYear_Click()
    Load frmSchoolYear
    frmSchoolYear.Show 1
End Sub

Private Sub mnuFMSections_Click()
    Load frmSection
    frmSection.Show 1
End Sub

Private Sub mnuFMShortcuts_Click()
    Load frmShortcutKeys
    frmShortcutKeys.Show 1
End Sub

Private Sub mnuFNewStudent_Click()
    Load frmStudentRecord
    frmStudentRecord.Show 1
End Sub

Private Sub mnuFTeacher_Click()
    Load frmTeacher
    frmTeacher.Show 1
End Sub

Private Sub mnuHAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub

Private Sub mnuKS_Click()
    Load frmShortcuts
    frmShortcuts.Show 1
End Sub

Private Sub mnuTCalculator_Click()
    On Error GoTo Err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "Calculator Missing"
End Sub

Private Sub mnuTEnrollStudent_Click()
    Load frmEnroll
    frmEnroll.Show 1
End Sub

Private Sub mnuTPayments_Click()
    Load frmPayment
    frmPayment.Show 1
End Sub

Private Sub mnuTSystemLock_Click()
    Load frmLock
    frmLock.Show 1
End Sub

Private Sub mnuVEnrolledStudents_Click()
    Load frmEnrollRecord
    frmEnrollRecord.Show 1
End Sub

Private Sub mnuVSections_Click()
    Load frmAssign
    frmAssign.Show 1
End Sub

Private Sub mnuVStudentInfo_Click()
    Load frmStudentRecord
    frmStudentRecord.Show
End Sub

Private Sub mnuVTAlignBottom_Click()
    Toolbar1.Align = vbAlignBottom
    MonthView1.Left = 12405
    MonthView1.Top = 135
End Sub

Private Sub mnuVTAlignLeft_Click()
    Toolbar1.Align = vbAlignLeft
    MonthView1.Left = 12405
    MonthView1.Top = 135
End Sub

Private Sub mnuVTAlignRight_Click()
    MonthView1.Left = 11790
    Toolbar1.Align = vbAlignRight
    MonthView1.Top = 135
End Sub

Private Sub mnuVTDefault_Click()
    Toolbar1.Align = vbAlignTop
    MonthView1.Left = 12405
    MonthView1.Top = 705
End Sub

Private Sub mnuVTeacherInfo_Click()
    Load frmTeacher
    frmTeacher.Show 1
End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 100
    i = i + 1
    StatusBar1.Panels(1).Text = Left(str1, i)
    If i = Len(str1) Then
        i = 1
        Timer1.Interval = 3000
    End If
End Sub

Private Sub UpDown1_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
        Case "student"
            Load frmStudentRecord
            frmStudentRecord.Show 1
            
        Case "enroll"
            Load frmEnroll
            frmEnroll.Show 1
            
        Case "pay"
            Load frmPayment
            frmPayment.Show 1
            
        Case "studentsearch"
            Load frmStudentRecord
            frmStudentRecord.Show 1
            
        Case "enrollsearch"
            Load frmEnrollRecord
            frmEnrollRecord.Show 1
            
        Case "section"
            Load frmAssign
            frmAssign.Show 1
        Case "print"
            MsgBox "not available"
        Case "lock"
            Load frmLock
            frmLock.Show 1
        Case "help"
            Load frmShortcuts
            frmShortcuts.Show 1
    End Select
        
End Sub


