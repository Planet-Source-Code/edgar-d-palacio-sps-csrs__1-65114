VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4470
      Begin VB.TextBox txtYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   945
         Width           =   1215
      End
      Begin VB.ComboBox cboSchoolYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Enrolled Students.frx":0000
         Left            =   90
         List            =   "Enrolled Students.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   585
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Preview"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3150
         TabIndex        =   1
         Top             =   1020
         Width           =   1215
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1305
         TabIndex        =   4
         Top             =   945
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2002
         BuddyControl    =   "Frame1"
         BuddyDispid     =   196609
         OrigLeft        =   1695
         OrigTop         =   960
         OrigRight       =   1935
         OrigBottom      =   1335
         Max             =   3000
         Min             =   2000
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Select Month and Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Err
'DataEnvironment1.rsCommand1.Source = "SELECT  TBilling.AccountNumber, TAccount.FullName, TAccount.Address, TAccount.Brand, TAccount.MeterNumber, TAccount.TypeOfConnection, TBilling.CubicMeterUsed, TBilling.Date, TBilling.Others, TBilling.PaymentWithDeduction, TBilling.Penalty, TBilling.PresReading, TBilling.PrevReading FROM TBilling, TAccount, Login WHERE TBilling.AccountNumber = TAccount.AccountNumber AND FORMAT(TBilling.Date,'mmmm')='" & cboMonth.Text & "' AND FORMAT(TBilling.Date,'yyyy')='" & txtYear.Text & "'"
DataEnvironment1.rsCommand1.Source = "SELECT qryEnrollment.Name, qryEnrollment.sSchoolYearName,qryEnrollment.sLevelName,qryEnrollment.sStudentSex FROM qryEnrollment WHERE qryEnrollment.sSchoolYearName='" & cboMonth.Text & "'"
DataEnvironment1.rsCommand1.Open
rptReceipt.Show
Unload Me
Exit Sub
Err:
DataEnvironment1.rsCommand1.Close
DataEnvironment1.rsCommand1.Source = "SELECT TBilling.AccountNumber, TAccount.FullName, TAccount.Address, TAccount.Brand, TAccount.MeterNumber, TAccount.TypeOfConnection, TBilling.CubicMeterUsed, TBilling.Date, TBilling.Others, TBilling.PaymentWithDeduction, TBilling.Penalty, TBilling.PresReading, TBilling.PrevReading FROM TBilling, TAccount, Login WHERE TBilling.AccountNumber = TAccount.AccountNumber AND FORMAT(TBilling.Date,'mmmm')='" & cboMonth.Text & "' AND FORMAT(TBilling.Date,'yyyy')='" & txtYear.Text & "'"
End Sub
