VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTeacher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Teacher.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1155
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
            Picture         =   "Teacher.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   4260
      TabIndex        =   22
      Top             =   825
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5085
      Left            =   135
      TabIndex        =   20
      Top             =   1560
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Teacher's Record"
      TabPicture(0)   =   "Teacher.frx":0724
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvTeacher"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtSearch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Personal Info"
      TabPicture(1)   =   "Teacher.frx":0740
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(8)=   "Label11"
      Tab(1).Control(9)=   "Label12"
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(11)=   "Label14"
      Tab(1).Control(12)=   "dtpGraduated"
      Tab(1).Control(13)=   "dtpBirthDate"
      Tab(1).Control(14)=   "txtLastName"
      Tab(1).Control(15)=   "txtFirstName"
      Tab(1).Control(16)=   "txtMiddleName"
      Tab(1).Control(17)=   "cboSex"
      Tab(1).Control(18)=   "cboCivilStatus"
      Tab(1).Control(19)=   "cboReligion"
      Tab(1).Control(20)=   "txtAge"
      Tab(1).Control(21)=   "txtContactNumber"
      Tab(1).Control(22)=   "txtAddress"
      Tab(1).Control(23)=   "txtCourse"
      Tab(1).Control(24)=   "txtIDDummy"
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "Employment Info"
      TabPicture(2)   =   "Teacher.frx":075C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command1"
      Tab(2).Control(1)=   "Text7"
      Tab(2).Control(2)=   "txtPhilHealth"
      Tab(2).Control(3)=   "txtSSS"
      Tab(2).Control(4)=   "txtTIN"
      Tab(2).Control(5)=   "txtComments"
      Tab(2).Control(6)=   "txtPosition"
      Tab(2).Control(7)=   "txtIDNumber"
      Tab(2).Control(8)=   "dtpStartContract"
      Tab(2).Control(9)=   "dtpEndContract"
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(11)=   "Label23"
      Tab(2).Control(12)=   "Label22"
      Tab(2).Control(13)=   "Label20"
      Tab(2).Control(14)=   "Label19"
      Tab(2).Control(15)=   "Label18"
      Tab(2).Control(16)=   "Label17"
      Tab(2).Control(17)=   "Label16"
      Tab(2).Control(18)=   "Label15"
      Tab(2).Control(19)=   "Label21"
      Tab(2).ControlCount=   20
      Begin VB.TextBox txtIDDummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -69420
         TabIndex        =   49
         Top             =   465
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Calculate"
         Height          =   300
         Left            =   -72645
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1890
         Width           =   1140
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1890
         Width           =   495
      End
      Begin VB.TextBox txtPhilHealth 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         TabIndex        =   19
         Top             =   4170
         Width           =   1980
      End
      Begin VB.TextBox txtSSS 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         TabIndex        =   18
         Top             =   3825
         Width           =   1980
      End
      Begin VB.TextBox txtTIN 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         TabIndex        =   17
         Top             =   3480
         Width           =   1980
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         Height          =   1020
         Left            =   -73215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2250
         Width           =   4845
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         TabIndex        =   15
         Top             =   1539
         Width           =   4605
      End
      Begin VB.TextBox txtCourse 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   10
         Top             =   4260
         Width           =   3720
      End
      Begin VB.TextBox txtIDNumber 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73215
         TabIndex        =   12
         Top             =   465
         Width           =   1980
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3600
         Width           =   4785
      End
      Begin VB.TextBox txtContactNumber 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   8
         Top             =   3255
         Width           =   2250
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   5
         Top             =   2220
         Width           =   390
      End
      Begin VB.ComboBox cboReligion 
         Height          =   315
         ItemData        =   "Teacher.frx":0778
         Left            =   -73200
         List            =   "Teacher.frx":077F
         TabIndex        =   7
         Top             =   2910
         Width           =   2520
      End
      Begin VB.ComboBox cboCivilStatus 
         Height          =   315
         ItemData        =   "Teacher.frx":0793
         Left            =   -73200
         List            =   "Teacher.frx":07A3
         TabIndex        =   6
         Top             =   2550
         Width           =   2520
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         ItemData        =   "Teacher.frx":07CA
         Left            =   -73200
         List            =   "Teacher.frx":07D4
         TabIndex        =   3
         Top             =   1515
         Width           =   1935
      End
      Begin VB.TextBox txtMiddleName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   2
         Top             =   1140
         Width           =   3720
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   1
         Top             =   810
         Width           =   3720
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -73200
         TabIndex        =   0
         Top             =   465
         Width           =   3720
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Left            =   915
         TabIndex        =   23
         Top             =   465
         Width           =   3555
      End
      Begin MSComctlLib.ListView lsvTeacher 
         Height          =   4050
         Left            =   135
         TabIndex        =   24
         Top             =   915
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7144
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Sex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Birth Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Age"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Civil Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Contact #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Address"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Course"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Position"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Start Contract"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "End Contract"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "T.I.N. #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "S.S.S. #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "PhilHealth #"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpBirthDate 
         Height          =   315
         Left            =   -73200
         TabIndex        =   4
         Top             =   1860
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpGraduated 
         Height          =   315
         Left            =   -73200
         TabIndex        =   11
         Top             =   4605
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpStartContract 
         Height          =   315
         Left            =   -73215
         TabIndex        =   13
         Top             =   813
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   38821
      End
      Begin MSComCtl2.DTPicker dtpEndContract 
         Height          =   315
         Left            =   -73215
         TabIndex        =   14
         Top             =   1176
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60555265
         CurrentDate     =   38821
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(note: calculate the year only)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71415
         TabIndex        =   50
         Top             =   1950
         Width           =   2610
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Year in service:"
         Height          =   195
         Left            =   -74670
         TabIndex        =   47
         Top             =   1950
         Width           =   1350
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "PhilHealth #:"
         Height          =   195
         Left            =   -74430
         TabIndex        =   45
         Top             =   4230
         Width           =   1110
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "S.S.S. #:"
         Height          =   195
         Left            =   -74130
         TabIndex        =   44
         Top             =   3885
         Width           =   810
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "T.I.N. #:"
         Height          =   195
         Left            =   -74070
         TabIndex        =   43
         Top             =   3540
         Width           =   750
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         Height          =   195
         Left            =   -74325
         TabIndex        =   42
         Top             =   2250
         Width           =   1005
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Position:"
         Height          =   195
         Left            =   -74055
         TabIndex        =   41
         Top             =   1592
         Width           =   735
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "End of Contract:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   40
         Top             =   1236
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Start of Contract:"
         Height          =   195
         Left            =   -74835
         TabIndex        =   39
         Top             =   873
         Width           =   1515
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Date Graduated:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   38
         Top             =   4650
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         Height          =   195
         Left            =   -74025
         TabIndex        =   37
         Top             =   4320
         Width           =   690
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "ID Number:"
         Height          =   195
         Left            =   -74340
         TabIndex        =   36
         Top             =   518
         Width           =   1020
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   -74100
         TabIndex        =   35
         Top             =   3675
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Contact Number:"
         Height          =   195
         Left            =   -74805
         TabIndex        =   34
         Top             =   3315
         Width           =   1470
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   195
         Left            =   -73740
         TabIndex        =   33
         Top             =   2265
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Birth Date:"
         Height          =   195
         Left            =   -74280
         TabIndex        =   32
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Religion:"
         Height          =   195
         Left            =   -74085
         TabIndex        =   31
         Top             =   2970
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Civil Status:"
         Height          =   195
         Left            =   -74385
         TabIndex        =   30
         Top             =   2610
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         Height          =   195
         Left            =   -73740
         TabIndex        =   29
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
         Height          =   195
         Left            =   -74505
         TabIndex        =   28
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   -74325
         TabIndex        =   27
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   -74310
         TabIndex        =   26
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Search:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   525
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   345
      Top             =   945
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
            Picture         =   "Teacher.frx":07E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Teacher.frx":0B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Teacher.frx":0F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Teacher.frx":12B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Teacher.frx":164E
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher's Personal Information"
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
      TabIndex        =   21
      Top             =   150
      Width           =   3765
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Teacher.frx":2318
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "frmTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummySelect As Integer

Private Sub Command1_Click()
    Text7 = Year(Date) - (Year(dtpStartContract.Value))
End Sub

Private Sub Form_Activate()
    Call forminit
    Call loadlsvTeacher
End Sub
Sub forminit()
    
    SSTab1.Tab = 0
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub
Private Sub Form_Load()
    Call PositionForm(frmTeacher)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        txtLastName.SetFocus
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
    End If
End Sub



Private Sub lsvTeacher_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    SSTab1.Tab = 1
    txtLastName.SetFocus
    row = lsvTeacher.SelectedItem.Index
    dummySelect = lsvTeacher.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblTeachers "
    strSQL = strSQL & "WHERE iTeacherID=" & dummySelect
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs1
        txtIDDummy = dummySelect
        txtLastName = !sTeacherLastName
        txtFirstName = !sTeacherFirstName
        txtMiddleName = !sTeacherMiddlename
        cboSex = !sTeacherSex
        dtpBirthDate = !dtTeacherBirthDate
        txtAge = !sTeacherAge
        cboCivilStatus = !sTeacherCivilStatus
        cboReligion = !sTeacherReligion
        txtContactNumber = !sTeacherContactNumber
        txtAddress = !sTeacherAddress
        txtCourse = !sTeacherCourse
        dtpGraduated = !dtTeacherGraduated
        txtIDNumber = !steacherIDNumber
        dtpStartContract = !dtStartContract
        dtpEndContract = !dtEndContract
        txtPosition = !sPosition
        txtComments = !sComments
        txtTIN = !sTIN
        txtSSS = !sSSS
        txtPhilHealth = !sPhilHealth
    End With
    Set rs1 = Nothing
End Sub

Private Sub lsvTeacher_KeyPress(KeyAscii As Integer)
    lsvTeacher_DblClick
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
        Case "add"
            Dim strSQLA As String
            Dim rsTeacherA As New ADODB.Recordset
            
            strSQLA = "SELECT * FROM tblTeachers"
            
            SSTab1.Tab = 1
            If complete = True Then
                rsTeacherA.Open strSQLA, cn, adOpenDynamic, adLockOptimistic
                
                With rsTeacherA
                    .AddNew
                    !sTeacherLastName = txtLastName
                    !sTeacherFirstName = txtFirstName
                    !sTeacherMiddlename = txtMiddleName
                    !sTeacherSex = cboSex
                    !dtTeacherBirthDate = dtpBirthDate
                    !sTeacherAge = txtAge
                    !sTeacherCivilStatus = cboCivilStatus
                    !sTeacherReligion = cboReligion
                    !sTeacherContactNumber = txtContactNumber
                    !sTeacherAddress = txtAddress
                    !sTeacherCourse = txtCourse
                    !dtTeacherGraduated = dtpGraduated
                    !steacherIDNumber = txtIDNumber
                    !dtStartContract = dtpStartContract
                    !dtEndContract = dtpEndContract
                    !sPosition = txtPosition
                    !sComments = txtComments
                    !sTIN = txtTIN
                    !sSSS = txtSSS
                    !sPhilHealth = txtPhilHealth
                    .Update
                End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new teacher record"
                    Set rsTeacherA = Nothing
                    Call textclear
                    SSTab1.Tab = 0
                    Call loadlsvTeacher
            Else
                If txtLastName = "" Then
                    MsgBox "Please type the LAST NAME", vbExclamation + vbOKOnly, "Empty Last Name"
                    txtLastName.SetFocus
                    Exit Sub
                ElseIf txtFirstName = "" Then
                    MsgBox "Please type the FIRST NAME", vbExclamation + vbOKOnly, "Empty First Name"
                    txtFirstName.SetFocus
                    Exit Sub
                ElseIf txtMiddleName = "" Then
                    MsgBox "Please type the MIDDLE NAME", vbExclamation + vbOKOnly, "Empty Middle Name"
                    txtMiddleName.SetFocus
                    Exit Sub
                ElseIf txtContactNumber = "" Then
                    MsgBox "PLease type the CONTACT NUMBER", vbExclamation + vbOKOnly, "Empty Contact Number"
                    txtContactNumber.SetFocus
                    Exit Sub
                End If
            End If
        Case "edit"
            Dim strSQLE As String
            Dim rsTeacherE As New ADODB.Recordset
            
            strSQLE = "SELECT * FROM tblTeachers"
            strSQLE = strSQLE & " WHERE iTeacherId =" & txtIDDummy
            
            If complete = True Then
            
                rsTeacherE.Open strSQLE, cn, adOpenDynamic, adLockOptimistic
                
                With rsTeacherE
                    !sTeacherLastName = txtLastName
                    !sTeacherFirstName = txtFirstName
                    !sTeacherMiddlename = txtMiddleName
                    !sTeacherSex = cboSex
                    !dtTeacherBirthDate = dtpBirthDate
                    !sTeacherAge = txtAge
                    !sTeacherCivilStatus = cboCivilStatus
                    !sTeacherReligion = cboReligion
                    !sTeacherContactNumber = txtContactNumber
                    !sTeacherAddress = txtAddress
                    !sTeacherCourse = txtCourse
                    !dtTeacherGraduated = dtpGraduated
                    !steacherIDNumber = txtIDNumber
                    !dtStartContract = dtpStartContract
                    !dtEndContract = dtpEndContract
                    !sPosition = txtPosition
                    !sComments = txtComments
                    !sTIN = txtTIN
                    !sSSS = txtSSS
                    !sPhilHealth = txtPhilHealth
                    .Update
                End With
                MsgBox "The changed you made was successfully updated", vbInformation + vbOKOnly, "Update teacher record"
                Set rsTeacherE = Nothing
                SSTab1.Tab = 0
                Call loadlsvTeacher
            Else
                If txtLastName = "" Then
                    MsgBox "Please type the LAST NAME", vbExclamation + vbOKOnly, "Empty Last Name"
                    txtLastName.SetFocus
                    Exit Sub
                ElseIf txtFirstName = "" Then
                    MsgBox "Please type the FIRST NAME", vbExclamation + vbOKOnly, "Empty First Name"
                    txtFirstName.SetFocus
                    Exit Sub
                ElseIf txtMiddleName = "" Then
                    MsgBox "Please type the MIDDLE NAME", vbExclamation + vbOKOnly, "Empty Middle Name"
                    txtMiddleName.SetFocus
                    Exit Sub
                ElseIf txtContactNumber = "" Then
                    MsgBox "PLease type the CONTACT NUMBER", vbExclamation + vbOKOnly, "Empty Contact Number"
                    txtContactNumber.SetFocus
                    Exit Sub
                End If
            End If
        Case "delete"
            If MsgBox("Do you want to DELETE " & lsvTeacher.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete Teacher Record") = vbYes Then
                cn.Execute "DELETE FROM tblTeachers WHERE iTeacherID=" & lsvTeacher.SelectedItem.Text
                Call loadlsvTeacher
            End If
        Case "search"
            Label13.Enabled = True
            txtSearch.Enabled = True
            txtSearch.SetFocus
    
    End Select
End Sub
Function complete()
    If txtLastName = "" Or txtFirstName = "" Or txtMiddleName = "" Or txtContactNumber = "" Then
        complete = False
    Else
        complete = True
    End If
End Function
Sub loadlsvTeacher()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblTeachers"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvTeacher.ListItems.Clear
    
    If rs.EOF Then
        lsvTeacher.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvTeacher.ListItems.Add(, , !iTeacherID, , 1)
                X.SubItems(1) = Trim(!sTeacherLastName) + ", " + Trim(!sTeacherFirstName) + " " + Trim(!sTeacherMiddlename)
                X.SubItems(2) = !steacherIDNumber
                X.SubItems(3) = !sTeacherSex
                X.SubItems(4) = !dtTeacherBirthDate
                X.SubItems(5) = !sTeacherAge
                X.SubItems(6) = !sTeacherCivilStatus
                X.SubItems(7) = !sTeacherContactNumber
                X.SubItems(8) = !sTeacherAddress
                X.SubItems(9) = !sTeacherCourse
                X.SubItems(10) = !sPosition
                X.SubItems(11) = !dtStartContract
                X.SubItems(12) = !dtEndContract
                X.SubItems(13) = !sTIN
                X.SubItems(14) = !sSSS
                X.SubItems(15) = !sPhilHealth
         .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Sub textclear()
    txtLastName = ""
    txtFirstName = ""
    txtMiddleName = ""
    cboSex = ""
    dtpBirthDate = Date
    txtAge = ""
    cboCivilStatus = ""
    cboReligion = ""
    txtContactNumber = ""
    txtAddress = ""
    txtCourse = ""
    dtpGraduated = Date
    txtIDNumber = ""
    dtpStartContract = Date
    dtpEndContract = Date
    txtPosition = ""
    txtComments = ""
    txtTIN = ""
    txtSSS = ""
    txtPhilHealth = ""
End Sub

Private Sub txtSearch_Change()
    Dim X
    Dim strSQL As String
    
        strSQL = "SELECT * FROM tblTeachers "
        strSQL = strSQL & "WHERE sTeacherLastName LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvTeacher.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvTeacher.ListItems.Add(, , !iTeacherID, , 1)
                X.SubItems(1) = Trim(!sTeacherLastName) + ", " + Trim(!sTeacherFirstName) + " " + Trim(!sTeacherMiddlename)
                X.SubItems(2) = !steacherIDNumber
                X.SubItems(3) = !sTeacherSex
                X.SubItems(4) = !dtTeacherBirthDate
                X.SubItems(5) = !sTeacherAge
                X.SubItems(6) = !sTeacherCivilStatus
                X.SubItems(7) = !sTeacherContactNumber
                X.SubItems(8) = !sTeacherAddress
                X.SubItems(9) = !sTeacherCourse
                X.SubItems(10) = !sPosition
                X.SubItems(11) = !dtStartContract
                X.SubItems(12) = !dtEndContract
                X.SubItems(13) = !sTIN
                X.SubItems(14) = !sSSS
                X.SubItems(15) = !sPhilHealth
                .MoveNext
            Loop
        End With
        Set rs = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvTeacher.ListItems.Count <> 0 Then
            lsvTeacher.SetFocus
        End If
    End If
End Sub
