VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEnroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Enroll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtpayment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   56
      Text            =   "Tuition Fee"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtEnrollmentIDDummy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   1050
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame Frame5 
      Caption         =   "Payment Details"
      Enabled         =   0   'False
      Height          =   3180
      Left            =   4920
      TabIndex        =   42
      Top             =   4890
      Width           =   4515
      Begin VB.TextBox txtOR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2730
         Width           =   645
      End
      Begin VB.TextBox txtRemaining 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2250
         Width           =   1725
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1770
         Width           =   1725
      End
      Begin VB.TextBox txtCashAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   6
         Top             =   1380
         Width           =   1725
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2490
         TabIndex        =   45
         Top             =   1005
         Width           =   1725
      End
      Begin VB.TextBox txtUponEnrollment 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   4
         Top             =   240
         Width           =   1725
      End
      Begin VB.TextBox txtPTA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2475
         TabIndex        =   5
         Top             =   615
         Width           =   1725
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "O.R. Number:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   855
         TabIndex        =   53
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Remaining Payable:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   330
         TabIndex        =   51
         Top             =   2310
         Width           =   1710
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Change:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1305
         TabIndex        =   49
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cash Amount:"
         Height          =   195
         Left            =   810
         TabIndex        =   47
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1485
         TabIndex        =   46
         Top             =   1065
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Upont Enrollment:"
         Height          =   195
         Left            =   495
         TabIndex        =   44
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "P.T.A."
         Height          =   195
         Left            =   1530
         TabIndex        =   43
         Top             =   675
         Width           =   510
      End
   End
   Begin MSComctlLib.ListView lsvStudents 
      Height          =   1260
      Left            =   1590
      TabIndex        =   13
      Top             =   1335
      Visible         =   0   'False
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   2223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   5644
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000010&
      Height          =   30
      Left            =   15
      TabIndex        =   35
      Top             =   6210
      Width           =   4695
   End
   Begin VB.ComboBox cboLevel 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5760
      Width           =   2790
   End
   Begin MSComCtl2.DTPicker dtpDateEnrolled 
      Height          =   285
      Left            =   1515
      TabIndex        =   0
      Top             =   5010
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   503
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38820
   End
   Begin VB.ComboBox cboSchoolYear 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5370
      Width           =   1950
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000010&
      Height          =   30
      Left            =   15
      TabIndex        =   24
      Top             =   4860
      Width           =   4695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2835
      Left            =   135
      TabIndex        =   10
      Top             =   1875
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   5001
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General Info"
      TabPicture(0)   =   "Enroll.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtLastName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFirstName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtMiddleName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAge"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSex"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSchool"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txtSchool 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   435
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2205
         Width           =   3090
      End
      Begin VB.TextBox txtSex 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1890
         Width           =   1185
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1545
         Width           =   660
      End
      Begin VB.TextBox txtMiddleName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   2850
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   855
         Width           =   2850
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   225
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   510
         Width           =   2850
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "School:"
         Height          =   195
         Left            =   630
         TabIndex        =   32
         Top             =   2205
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         Height          =   195
         Left            =   870
         TabIndex        =   23
         Top             =   1890
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age:"
         Height          =   195
         Left            =   870
         TabIndex        =   21
         Top             =   1545
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   285
         TabIndex        =   17
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   510
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   300
      Left            =   4725
      Picture         =   "Enroll.frx":03A6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1035
      Width           =   360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Height          =   30
      Left            =   -90
      TabIndex        =   9
      Top             =   1695
      Width           =   9945
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Enroll.frx":0730
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Enroll.frx":0ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Enroll.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Enroll.frx":11FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Enroll.frx":1598
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   6120
      TabIndex        =   7
      Top             =   885
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Save transaction"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Search enrolled student"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLevelDummy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5760
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtSchoolYearDummy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2985
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   5385
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtStudentNameDummy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   1035
      Visible         =   0   'False
      Width           =   450
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2835
      Left            =   4935
      TabIndex        =   25
      Top             =   1875
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   5001
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Enroll.frx":1932
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTotal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lsvFees"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSComctlLib.ListView lsvFees 
         Height          =   1920
         Left            =   150
         TabIndex        =   27
         Top             =   450
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   3387
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "AMOUNT"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL FEES WHOLE YEAR:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   405
         TabIndex        =   34
         Top             =   2505
         Width           =   2295
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3660
         TabIndex        =   33
         Top             =   2505
         Width           =   375
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1590
      TabIndex        =   26
      Top             =   1050
      Width           =   3060
   End
   Begin VB.Frame Frame4 
      Caption         =   "Payment Type"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   915
      TabIndex        =   40
      Top             =   6570
      Width           =   2880
      Begin VB.OptionButton optPartialPayment 
         Caption         =   "Partial Payment"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   285
         Width           =   1860
      End
      Begin VB.OptionButton optFullPayment 
         Caption         =   "Full Payment"
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   705
         Width           =   1860
      End
   End
   Begin VB.TextBox txtPaymentDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2760
      TabIndex        =   54
      Top             =   7005
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtDummyTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Level:"
      Height          =   195
      Left            =   960
      TabIndex        =   30
      Top             =   5820
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Date Enrolled:"
      Height          =   195
      Left            =   255
      TabIndex        =   29
      Top             =   5055
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "School Year:"
      Height          =   195
      Left            =   390
      TabIndex        =   28
      Top             =   5430
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Student Name:"
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   1110
      Width           =   1290
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Enroll.frx":194E
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment Info"
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
      TabIndex        =   8
      Top             =   150
      Width           =   1860
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Enroll.frx":2618
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmEnroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummySelect As Integer

Private Sub Form_Load()
    Call PositionForm(frmEnroll)
    Call loadcboLevel
    Call loadcboSchoolYear
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Call textclear
    Call textpaymentclear
End Sub
Private Sub cboLevel_Click()
    On Error Resume Next
    Dim strSQL As String
    Dim lindex
    lindex = cboLevel.ItemData(cboLevel.ListIndex)
    
    strSQL = "SELECT * FROM tblLevels"
    strSQL = strSQL & " WHERE iLevelID=" & lindex
    
    rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
     
    With rs3
        txtLevelDummy = !iLevelID
    End With
    Set rs3 = Nothing
End Sub
Private Sub cboSchoolYear_Click()
    On Error Resume Next
    Dim strSQL As String
    Dim sindex
    sindex = cboSchoolYear.ItemData(cboSchoolYear.ListIndex)
    
    strSQL = "SELECT * FROM tblSchoolYears"
    strSQL = strSQL & " WHERE iSchoolYearID=" & sindex
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
     
    With rs
        txtSchoolYearDummy = !iSchoolYearID
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdAdd_Click()
    Load frmStudentRecord
    frmStudentRecord.Show 1
End Sub

Private Sub Form_Activate()
    Call forminit
    Call loadlsvStudents
End Sub

Sub forminit()
    Toolbar1.Buttons(2).Enabled = False
End Sub
Sub loadlsvStudents()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblStudents"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvStudents.ListItems.Clear
    
    If rs.EOF Then
        lsvStudents.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvStudents.ListItems.Add(, , !istudentID)
                X.SubItems(1) = Trim(!sStudentLastName) + ", " + Trim(!sStudentFirstName) + " " + Trim(!sStudentMiddlename)
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        Call textpaymentclear
        lsvStudents.Visible = False
    End If
End Sub

Private Sub lsvStudents_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim row
    
    dummyButton = "edit"
    row = lsvStudents.SelectedItem.Index
    dummySelect = lsvStudents.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblStudents "
    strSQL = strSQL & "WHERE iStudentID=" & dummySelect
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs1
        txtStudentNameDummy = dummySelect
        txtLastName = !sStudentLastName
        txtFirstName = !sStudentFirstName
        txtMiddleName = !sStudentMiddlename
        txtAge = !sStudentAge
        txtSex = !sStudentSex
        txtSchool = !sSchoolLastAttended
    End With
        lsvStudents.Visible = False
        dtpDateEnrolled.SetFocus
    Set rs1 = Nothing
End Sub

Private Sub lsvStudents_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lsvStudents_Click
    End If
End Sub



Private Sub optFullPayment_Click()
    Toolbar1.Buttons(2).Enabled = True
    Call textpaymentclear
    txtTotal = txtDummyTotal
    txtPaymentDummy.Text = "Full"
    txtUponEnrollment = 0
    Call GenerateNumber
    If optFullPayment.Value = True Then
        Frame5.Enabled = True
        Label16.Visible = False
        txtUponEnrollment.Visible = False
        'txtPTA.SetFocus
    End If
End Sub

Private Sub optPartialPayment_Click()
    Toolbar1.Buttons(2).Enabled = True
    Call textpaymentclear
    Call GenerateNumber
    txtPaymentDummy.Text = "Partial"
    If optPartialPayment.Value = True Then
        Frame5.Enabled = True
        Label16.Visible = True
        txtUponEnrollment.Visible = True
        'txtUponEnrollment.SetFocus
    End If
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim rsEnrollment As New ADODB.Recordset
            Dim strSQLEnrollment As String
            
            strSQLEnrollment = "SELECT * FROM tblEnrollment"
            
            If complete = True Then
            
                If Val(txtCashAmount) >= Val(txtTotal) Then
                    rsEnrollment.Open strSQLEnrollment, cn, adOpenDynamic, adLockOptimistic
                    
                    With rsEnrollment
                        .AddNew
                        !istudentID = txtStudentNameDummy
                        !iSchoolYearID = txtSchoolYearDummy
                        !iLevelID = txtLevelDummy
                        !dtEnrolled = dtpDateEnrolled
                        !sPaymentType = txtPaymentDummy
                        !cUponEnrollment = Val(txtUponEnrollment)
                        !cPTA = Val(txtPTA)
                        !cTotal = Val(txtTotal)
                        !cCashAmount = Val(txtCashAmount)
                        !cChange = Val(txtChange)
                        !cRemaining = Val(txtRemaining)
                        !sORNumber = txtOR
                        !bIsEnroll = True
                        !cAccount = Val(txtDummyTotal)
                        .Update
                    End With
                    Set rsEnrollment = Nothing
                    If MsgBox("Do you want to print the receipt before saving?", vbYesNo + vbQuestion, "Print Official Receipt") = vbYes Then
                        Call printreceipt
                    End If
                    MsgBox "New enrollment was successfully addet to the database", vbInformation + vbOKOnly, "Add new record"
                    Call textclear
                    Call textpaymentclear
                    txtSearch.SetFocus
                Else
                    MsgBox "The amount you type is less than the total amount", vbCritical + vbOKOnly, "Incorrect amount"
                    txtCashAmount.SetFocus
                    Exit Sub
                End If
            Else
                If txtCashAmount = "" Then
                    MsgBox "Payment not accepted", vbCritical + vbOKOnly, "Payment field is empty"
                    txtCashAmount.SetFocus
                    Exit Sub
                ElseIf txtOR = "" Then
                    MsgBox "Please copy the O.R. NUMBER from the receipt", vbCritical, "O.R. field is empty"
                    txtOR.SetFocus
                    Exit Sub
                End If
            End If
            
        Case "update"
            Dim rsEnrollmentE As New ADODB.Recordset
            Dim strSQLEnrollmentE As String
            
            strSQLEnrollmentE = "SELECT * FROM tblEnrollment"
            strSQLEnrollmentE = strSQLEnrollmentE & " WHERE iEnrollmentID=" & txtEnrollmentIDDummy
            If complete = True Then
            
                If Val(txtCashAmount) >= Val(txtTotal) Then
                    rsEnrollmentE.Open strSQLEnrollmentE, cn, adOpenDynamic, adLockOptimistic
                    
                    With rsEnrollmentE
                        !istudentID = txtStudentNameDummy
                        !iSchoolYearID = txtSchoolYearDummy
                        !iLevelID = txtLevelDummy
                        !dtEnrolled = dtpDateEnrolled
                        !sPaymentType = txtPaymentDummy
                        !cUponEnrollment = Val(txtUponEnrollment)
                        !cPTA = Val(txtPTA)
                        !cTotal = Val(txtTotal)
                        !cCashAmount = Val(txtCashAmount)
                        !cChange = Val(txtChange)
                        !cRemaining = Val(txtRemaining)
                        !sORNumber = txtOR
                        !bIsEnroll = True
                        !cAccount = Val(txtDummyTotal)
                        .Update
                    End With
                    Set rsEnrollmentE = Nothing
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update record"
                    Call textclear
                    Call textpaymentclear
                    txtSearch.SetFocus
                Else
                    MsgBox "The amount you type is less than the total amount", vbCritical + vbOKOnly, "Incorrect amount"
                    txtCashAmount.SetFocus
                    Exit Sub
                End If
            Else
                If txtCashAmount = "" Then
                    MsgBox "Payment not accepted", vbCritical + vbOKOnly, "Payment field is empty"
                    txtCashAmount.SetFocus
                    Exit Sub
                ElseIf txtOR = "" Then
                    MsgBox "Please copy the O.R. NUMBER from the receipt", vbCritical, "O.R. field is empty"
                    txtOR.SetFocus
                    Exit Sub
                End If
            End If
        Case "delete"
            Dim rsEnrollmentD As New ADODB.Recordset
            Dim strSQLEnrollmentD As String
            Dim answer As String
            
            strSQLEnrollmentD = "SELECT * FROM tblEnrollment"
            
            rsEnrollmentD.Open strSQLEnrollmentD, cn, adOpenDynamic, adLockOptimistic
            
            answer = MsgBox("Are you sure you want to delete the this record", vbQuestion + vbYesNo, "Remove from database")
            
            If answer = vbYes Then
                With rsEnrollmentD
                    .Delete
                End With
                Call textclear
                Call textpaymentclear
                MsgBox "The record was successfully removed", vbInformation + vbOKOnly
            Else
                Call textclear
                Call textpaymentclear
            End If
            Set rsEnrollmentD = Nothing
        
        Case "search"
            lsvStudents.Visible = False
            Load frmEnrollRecord
            frmEnrollRecord.Show 1
        Case "print"
            Call printreceipt
    End Select
    Exit Sub
End Sub
Function complete()
    If txtCashAmount = "" Or txtOR = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Private Sub txtCashAmount_Change()
    'If txtCashAmount < txtTotal Then
     '   MsgBox "The amount must be greater than the Total", vbCritical + vbOKOnly, "Invalid amount"
    'Else
        txtChange = Val(txtCashAmount) - Val(txtTotal)
    'End If
End Sub

Private Sub txtLevelDummy_Change()
    Call DisplayFees
    Frame4.Enabled = True
End Sub



Private Sub txtPTA_Change()
    If optPartialPayment.Value = True Then
        txtTotal = Val(txtUponEnrollment) + Val(txtPTA)
        txtRemaining = Val(txtDummyTotal) - Val(txtUponEnrollment)
    ElseIf optFullPayment.Value = True Then
        txtTotal = Val(txtDummyTotal) + Val(txtPTA)
    End If
End Sub

Private Sub txtSchoolYearDummy_Change()
    Call DisplayFees
End Sub

Private Sub txtSearch_Change()
    Dim X
    Dim strSQL As String
    
        strSQL = "SELECT * FROM tblStudents "
        strSQL = strSQL & "WHERE sStudentLastName LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudents.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvStudents.ListItems.Add(, , !istudentID)
                X.SubItems(1) = Trim(!sStudentLastName) + ", " + Trim(!sStudentFirstName) + " " + Trim(!sStudentMiddlename)
                .MoveNext
            Loop
        End With
        Set rs = Nothing
End Sub

Private Sub txtSearch_GotFocus()
    lsvStudents.Visible = True
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvStudents.ListItems.Count <> 0 Then
            lsvStudents.SetFocus
        End If
    End If
End Sub
Sub loadcboLevel()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblLevels"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboLevel.AddItem !sLevelName
            cboLevel.ItemData(cboLevel.NewIndex) = CLng(!iLevelID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub
Sub loadcboSchoolYear()
    On Error GoTo err_handler:
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblSchoolYears"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboSchoolYear.AddItem !sSchoolYearName
            cboSchoolYear.ItemData(cboSchoolYear.NewIndex) = CLng(!iSchoolYearID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
Exit Sub
err_handler:
    Set rs = Nothing
End Sub
Sub DisplayFees()
    Dim X
    Dim strSQL As String
    Dim total
    
    
    If cboSchoolYear.ListIndex <> -1 And cboLevel.ListIndex <> -1 Then
    
        SSTab2.Caption = "School Fees for " & cboSchoolYear.Text & " - " & cboLevel.Text
        strSQL = "SELECT * FROM qryFees "
        strSQL = strSQL & "WHERE iSchoolYearID=" & txtSchoolYearDummy
        strSQL = strSQL & " AND iLevelID=" & txtLevelDummy
        
        rs4.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvFees.ListItems.Clear
        
        With rs4
            Do While Not rs4.EOF
                Set X = lsvFees.ListItems.Add(, , !iFeeID)
                X.SubItems(1) = !sFeeName
                X.SubItems(2) = !cAmount
                total = total + !cAmount
                .MoveNext
            Loop
                txtDummyTotal = total
                lblTotal.Caption = Format(total, "###,###.00")
        End With
        Set rs4 = Nothing
    End If
End Sub

Private Sub txtUponEnrollment_Change()
    If optPartialPayment.Value = True Then
        txtTotal = Val(txtUponEnrollment) + Val(txtPTA)
        txtRemaining = Val(txtDummyTotal) - Val(txtUponEnrollment)
    Else
        txtTotal = txtDummyTotal
        txtRemaining = 0
    End If
End Sub

Sub textpaymentclear()
    txtUponEnrollment = ""
    txtPTA = ""
    txtTotal = ""
    txtCashAmount = ""
    txtChange = ""
    txtRemaining = ""
    txtOR = ""
End Sub
Sub textclear()
    txtLastName = ""
    txtFirstName = ""
    txtMiddleName = ""
    txtAge = ""
    txtSex = ""
    txtSchool = ""
    dtpDateEnrolled.Value = Date
    cboSchoolYear.ListIndex = -1
    cboLevel.ListIndex = -1
    lsvFees.ListItems.Clear
    optPartialPayment.Value = False
    optFullPayment.Value = False
    lblTotal.Caption = "0.00"
End Sub
Sub printreceipt()
    row = 10
    col = 10
    With frmEnroll
        Call set_myprinterobject
        set_reportpath = App.Path & "\Reports\OfficialReceipt.xls"
        report_gen.Workbooks.Open (set_reportpath)
        report_gen.Worksheets(1).Cells(7, 10) = .txtOR.Text
        report_gen.Worksheets(1).Cells(10, 3) = .dtpDateEnrolled
        report_gen.Worksheets(1).Cells(12, 3) = Trim(.txtLastName.Text) & ", " & Trim(.txtFirstName.Text) & " " & Trim(.txtMiddleName.Text)
        report_gen.Worksheets(1).Cells(14, 3) = .txtpayment.Text
        report_gen.Worksheets(1).Cells(16, 3) = Format(.lblTotal, "###,###.00")
        report_gen.Worksheets(1).Cells(17, 3) = Format(.txtRemaining, "###,###.00")
        report_gen.Worksheets(1).Cells(18, 3) = Format(.txtCashAmount, "###,###.00")
        report_gen.Worksheets(1).Cells(19, 3) = Format(.txtChange, "###,###.00")
        
        report_gen.Visible = True
        report_gen.ActiveWindow.SelectedSheets.PrintPreview
        report_gen.ActiveWindow.Close (False)
        report_gen.Quit
    End With
        Set report_gen = Nothing
End Sub
Sub GenerateNumber()
    Dim X As String
    Dim Y As Integer
    Dim tdate
    
    'On Error GoTo err_handler:
    
    rs3.Open "SELECT * FROM tblEnrollment order by sORNumber", cn, adOpenDynamic, adLockOptimistic
    rs3.MoveLast
    
    If rs3.EOF = False Then
        X = Mid$(rs3!sORNumber, 1, 3)
        tdate = Format(Date, "yy")
    End If
    X = X
    If tdate = Mid$(rs3!sORNumber, 1, 2) Then
        If CInt(X) > 99999 Then
            X = Mid$(rs3!sORNumber, 1, 3)
            Y = Format(CInt(X + 1), "000")
            txtOR = "-x" & Format(Y, "000") & ""
        ElseIf CInt(X) < 99999 Then
            Y = Format(CInt(X + 1), "000")
            txtOR = Format(Y, "000") & ""
        End If
    Else
        X = 0
        X = Mid$(rs3!sORNumber, 1, 3)
        Y = Format(CInt(X + 1), "000")
        txtOR = Format(Y, "000") & ""
    End If
        Set rs3 = Nothing
    'Exit Sub
'err_handler:
'    Y = Format(CInt(X), "000")
'    txtOR = Format(Y, "000") & ""
'    Set rs3 = Nothing
End Sub

