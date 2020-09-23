VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Payment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLevelIDDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3420
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txtEnrollmentIDDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2700
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txtStudentIDDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1980
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Frame Frame3 
      Caption         =   "Summary of Payments"
      Height          =   2535
      Left            =   165
      TabIndex        =   9
      Top             =   5625
      Width           =   7710
      Begin MSComctlLib.ListView lsvSummary 
         Height          =   2160
         Left            =   135
         TabIndex        =   10
         Top             =   240
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   3810
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date Paid"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amount Paid"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "O. Balance"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   3210
      TabIndex        =   4
      Top             =   1590
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Payment.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLevel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPayable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame Frame2 
         Caption         =   "Payment Details"
         ForeColor       =   &H00FF0000&
         Height          =   1890
         Left            =   240
         TabIndex        =   13
         Top             =   1650
         Width           =   4200
         Begin MSComCtl2.DTPicker dtpDatePaid 
            Height          =   300
            Left            =   1965
            TabIndex        =   17
            Top             =   330
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60293121
            CurrentDate     =   38825
         End
         Begin VB.TextBox txtAmount 
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
            Left            =   1965
            TabIndex        =   15
            Top             =   690
            Width           =   1590
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Current Oustanding Balance"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   60
            TabIndex        =   23
            Top             =   1200
            Width           =   4065
         End
         Begin VB.Label lblOutstandingBalance 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1035
            TabIndex        =   18
            Top             =   1455
            Width           =   2085
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Date Paid:"
            Height          =   195
            Left            =   975
            TabIndex        =   16
            Top             =   390
            Width           =   900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Amount Paid:"
            Height          =   195
            Left            =   720
            TabIndex        =   14
            Top             =   750
            Width           =   1155
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Amount Payable:"
         Height          =   195
         Left            =   615
         TabIndex        =   12
         Top             =   1185
         Width           =   1470
      End
      Begin VB.Label lblPayable 
         AutoSize        =   -1  'True
         Caption         =   "Amount Payable"
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
         Left            =   2190
         TabIndex        =   11
         Top             =   1185
         Width           =   1590
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         Caption         =   "Level Name"
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
         Left            =   2190
         TabIndex        =   8
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year Level:"
         Height          =   195
         Left            =   1110
         TabIndex        =   7
         Top             =   870
         Width           =   975
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Complete Name"
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
         Left            =   2190
         TabIndex        =   6
         Top             =   585
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1515
         TabIndex        =   5
         Top             =   585
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   180
      TabIndex        =   1
      Top             =   1500
      Width           =   2880
      Begin MSComctlLib.ListView lsvStudentList 
         Height          =   3165
         Left            =   135
         TabIndex        =   3
         Top             =   510
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   5583
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4586
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   2595
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Payment.frx":03A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Payment.frx":0740
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Payment.frx":0ADA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   5730
      TabIndex        =   19
      Top             =   840
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
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
            Key             =   "edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Payment.frx":0E74
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payments Info"
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
      Width           =   1740
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Payment.frx":1B3E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8025
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummySelect As Integer
Private Sub Form_Activate()
    Call forminit
    Call lsvLoadStudentList
End Sub

Private Sub Form_Load()
    Call PositionForm(frmPayment)
End Sub
Sub forminit()

End Sub
Sub lsvLoadStudentList()
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
                Set X = lsvStudentList.ListItems.Add(, , !istudentID)
                X.SubItems(1) = !Name
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub
Sub lsvloadSummary()
    Dim X
    Dim strSQL As String
    Dim rsSummary As New ADODB.Recordset
    
    strSQL = "SELECT * FROM tblPayments"
    strSQL = strSQL & " WHERE iStudentID=" & txtStudentIDDummy
    
    rsSummary.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvSummary.ListItems.Clear
    
    If rsSummary.EOF = True Then
        lsvSummary.ListItems.Clear
    Else
        With rsSummary
            Do While Not rsSummary.EOF
                Set X = lsvSummary.ListItems.Add(, , !iPaymentID)
                X.SubItems(2) = !sStudentName
                X.SubItems(3) = !dtPaid
                X.SubItems(4) = !cAmount
                X.SubItems(5) = !cOutstandingBalance
                X.SubItems(6) = !istudentID
                .MoveNext
            Loop
        End With
    End If
    Set rsSummary = Nothing
End Sub
Private Sub lsvStudentList_Click()
        Dim X As Integer
        Dim strSQL As String
        Dim row
        row = lsvStudentList.SelectedItem.Index
        dummySelect = lsvStudentList.ListItems.Item(row).Text
        
        strSQL = "SELECT * FROM qryEnrollment "
        strSQL = strSQL & "WHERE iStudentID=" & dummySelect
        
        rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        With rs1
            txtStudentIDDummy = !istudentID
            txtEnrollmentIDDummy = !iEnrollmentID
            txtLevelIDDummy = !iLevelID
            lblName = !Name
            lblLevel = !sLevelName
            lblPayable = !cRemaining
        End With
        Call lsvloadSummary
        MsgBox "Double click the last payment in the list", vbInformation + vbOKOnly, "Add new payment"
        lsvSummary.SetFocus
        Set rs1 = Nothing
End Sub

Private Sub lsvStudentList_KeyPress(KeyAscii As Integer)
    lsvStudentList_Click
End Sub

Private Sub lsvSummary_DblClick()
    If lsvSummary.SelectedItem.SubItems(5) = 0 Then
        MsgBox "Outstanding balance closed", vbInformation + vbOKOnly, "Paid"
    Else
        lblPayable = lsvSummary.SelectedItem.SubItems(5)
    End If
    End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo err_handler:
    Select Case Button.Key
        Case "add"
            Dim rsPayments As New ADODB.Recordset
            Dim strSQLPayments As String
            
            strSQLPayments = "SELECT * FROM tblPayments"
            
            If txtAmount <> "" Then
                rsPayments.Open strSQLPayments, cn, adOpenDynamic, adLockOptimistic
                
                With rsPayments
                    .AddNew
                    !iEnrollmentID = txtEnrollmentIDDummy
                    !istudentID = txtStudentIDDummy
                    !iLevelID = txtLevelIDDummy
                    !sStudentName = lblName
                    !sLevelName = lblLevel
                    !dtPaid = dtpDatePaid
                    !cAmount = Val(txtAmount)
                    !cPayable = Val(lblPayable)
                    !cOutstandingBalance = Val(lblOutstandingBalance)
                    .Update
                End With
                    MsgBox "New transaction added to the database", vbInformation + vbOKOnly, "Add new payment"
                    Call lsvloadSummary
            Else
                If txtAmount = "" Then
                    MsgBox "Please type AMOUNT", vbExclamation + vbOKOnly, "Empty amount"
                    txtAmount.SetFocus
                    Exit Sub
                End If
            End If
        Case "edit"
        
        Case "print"
    End Select
    Exit Sub
err_handler:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "Important Fiels are empty"
    Exit Sub
End Sub

Private Sub txtAmount_Change()
    On Error Resume Next
    lblOutstandingBalance = Val(lblPayable) - Val(txtAmount)
    Exit Sub
End Sub


Private Sub txtSearch_Change()
    Dim X
    Dim strSQL As String
    
        strSQL = "SELECT * FROM qryEnrollment "
        strSQL = strSQL & "WHERE name LIKE'" & txtSearch.Text & "%'"
        
        rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudentList.ListItems.Clear
        With rs
            Do While Not rs.EOF
                Set X = lsvStudentList.ListItems.Add(, , !istudentID)
                X.SubItems(1) = !Name
                .MoveNext
            Loop
        End With
        Set rs = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvStudentList.ListItems.Count <> 0 Then
            lsvStudentList.SetFocus
        End If
    End If
End Sub
