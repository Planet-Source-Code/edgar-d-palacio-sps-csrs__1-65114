VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAssign 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Assign.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSection 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5895
      TabIndex        =   27
      Top             =   975
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtSchoolYear 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5910
      TabIndex        =   26
      Top             =   1515
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5910
      TabIndex        =   25
      Top             =   1515
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtLevelIDDummy 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5430
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Section Record"
      Height          =   5595
      Left            =   150
      TabIndex        =   13
      Top             =   855
      Width           =   5205
      Begin VB.Frame Frame3 
         Caption         =   "Search Option"
         Height          =   810
         Left            =   3405
         TabIndex        =   22
         Top             =   150
         Width           =   1650
         Begin VB.OptionButton optLevel 
            Caption         =   "Section"
            Height          =   285
            Left            =   195
            TabIndex        =   24
            Top             =   465
            Width           =   1290
         End
         Begin VB.OptionButton optLastName 
            Caption         =   "Last Name"
            Height          =   285
            Left            =   195
            TabIndex        =   23
            Top             =   225
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin MSComctlLib.ListView lsvStudentSection 
         Height          =   4365
         Left            =   150
         TabIndex        =   15
         Top             =   1080
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   7699
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student #"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Section"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "School Year"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   150
         TabIndex        =   14
         Top             =   435
         Width           =   2745
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   3000
         Picture         =   "Assign.frx":038A
         Stretch         =   -1  'True
         Top             =   435
         Width           =   285
      End
   End
   Begin VB.TextBox txtEnrollmentID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   5430
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3945
      Left            =   5520
      TabIndex        =   0
      Top             =   1965
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   6959
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Section Details"
      TabPicture(0)   =   "Assign.frx":0714
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
      Tab(0).Control(6)=   "lblLastName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblFirstName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblMiddleName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLevel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtStudentNumber"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboSection"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   120
         TabIndex        =   20
         Top             =   1905
         Width           =   4035
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Generate"
         Height          =   435
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1995
         Width           =   1485
      End
      Begin VB.ComboBox cboSection 
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2595
         Width           =   2355
      End
      Begin VB.TextBox txtStudentNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3150
         Width           =   960
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   19
         Top             =   1485
         Width           =   60
      End
      Begin VB.Label lblMiddleName 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   18
         Top             =   1170
         Width           =   60
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   17
         Top             =   855
         Width           =   60
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
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
         Left            =   1800
         TabIndex        =   16
         Top             =   540
         Width           =   60
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Select Section:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   2655
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   1125
         TabIndex        =   7
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Student Number:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   3195
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1170
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   660
         TabIndex        =   3
         Top             =   855
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   675
         TabIndex        =   2
         Top             =   540
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5250
      Top             =   1890
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
            Picture         =   "Assign.frx":0730
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Assign.frx":0ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Assign.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Assign.frx":11FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   7800
      TabIndex        =   10
      Top             =   810
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
            Key             =   "delete"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assign student section"
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
      TabIndex        =   1
      Top             =   150
      Width           =   2760
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "Assign.frx":1598
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Assign.frx":2262
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10035
   End
End
Attribute VB_Name = "frmAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboSection_Click()
    Call GenerateNumber
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qrySections"
    strSQL = strSQL & " WHERE iLevelID=" & txtLevelIDDummy
    
    Command1.Enabled = False
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
        MsgBox "No section(s) assigned to this level", vbExclamation + vbOKOnly, "Empty section"
        Unload Me
    Else
        With rs
            Do While Not rs.EOF
                cboSection.AddItem !Section
                cboSection.ItemData(cboSection.NewIndex) = CLng(!iSectionID)
                .MoveNext
            Loop
        End With
        
    End If
    Set rs = Nothing
End Sub

Sub formint()
    
End Sub
Sub loadcboLevel()
    
End Sub

Sub GenerateNumber()
    Dim X As String
    Dim Y As Integer
    Dim tdate
    
    On Error GoTo err_handler:
    
    rs1.Open "SELECT * FROM tblAssignSection order by sStudentNumber", cn, adOpenDynamic, adLockOptimistic
    rs1.MoveLast
    
    If rs1.EOF = False Then
        'x = Mid$(rs1!s, 10, 5)
        X = Mid$(rs1!sStudentNumber, 4, 4)
        'tdate = Format(Date, "yyyy")
        tdate = Format(Date, "yy")
    End If
    X = X
    'If tdate = Mid$(rs1!sStudentNUmber, 5, 4) Then
    If tdate = Mid$(rs1!sStudentNumber, 1, 2) Then
        If CInt(X) > 99999 Then
            'x = Mid$(rs1!sStudentNumber, 11, 5)
            X = Mid$(rs1!sStudentNumber, 4, 4)
            'Y = Format(CInt(x + 1), "0000")
            Y = Format(CInt(X + 1), "0000")
            'txtStudentNumber = "SPS-" & Format(Date, "yyyy") & "-X" & Format(Y, "0000") & ""
            txtStudentNumber = Format(Date, "yy") & "-x" & Format(Y, "0000") & ""
        ElseIf CInt(X) < 99999 Then
            'Y = Format(CInt(x + 1), "00000")
            Y = Format(CInt(X + 1), "00000")
            'txtStudentNumber = "SPS-" & Format(Date, "yyyy") & "-" & Format(Y, "00000") & ""
            txtStudentNumber = Format(Date, "yy") & "-" & Format(Y, "0000") & ""
        End If
    Else
        X = 0
        'Y = Format(CInt(x), "00000")
        Y = Format(CInt(X), "0000")
        'txtstudid = "SPS-" & Format(Date, "yyyy") & "-" & Format(Y, "00000") & ""
        txtStudentNumber = Format(Date, "yy") & "-" & Format(Y, "0000") & ""
    End If
        Set rs1 = Nothing
    Exit Sub
err_handler:
    X = 0
    'Y = Format(CInt(x), "00000")
    Y = Format(CInt(X), "0000")
    'txtStudentNumber = "SPS-" & Format(Date, "yyyy") & "-" & Format(Y, "00000") & ""
    txtStudentNumber = Format(Date, "yy") & "-" & Format(Y, "0000") & ""
    Set rs1 = Nothing
End Sub

Private Sub Form_Activate()
    Command1.Enabled = True
    Command1.SetFocus
    Call LoadlsvStudentSection
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
    Me.Top = 4000
    Me.Left = 4000
End Sub



Private Sub lsvStudentSection_Click()
    lsvStudentSection.Sorted = True
    lsvStudentSection.SortKey = 2
    lsvStudentSection.SortOrder = lvwAscending
    txtSection = lsvStudentSection.SelectedItem.SubItems(3)
    txtSchoolYear = lsvStudentSection.SelectedItem.SubItems(5)
End Sub

Private Sub lsvStudentSection_DblClick()
    lsvStudentSection.Sorted = True
    lsvStudentSection.SortKey = 4
    lsvStudentSection.SortOrder = lvwDescending
    txtSection = lsvStudentSection.SelectedItem.SubItems(3)
    txtSchoolYear = lsvStudentSection.SelectedItem.SubItems(5)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim rsSection As New ADODB.Recordset
            Dim strSQLSection As String
            
            strSQLSection = "SELECT * FROM tblAssignSection"
            
            If complete = True Then
                rsSection.Open strSQLSection, cn, adOpenDynamic, adLockOptimistic
                
                With rsSection
                    .AddNew
                    !iEnrollmentID = txtEnrollmentID
                    !iSectionID = cboSection.ItemData(cboSection.ListIndex)
                    !sStudentNumber = txtStudentNumber
                    !sStudentName = Trim(lblLastName) & ", " & Trim(lblFirstName) & " " & Trim(lblMiddleName)
                    !sSectionName = cboSection.Text
                    !sStudentSex = txtGender
                    !sSchoolYearName = txtSchoolYear
                    .Update
                End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Student Section"
                    Unload Me
            Else
                If cboSection.ListIndex = -1 Then
                    MsgBox "Please select a section", vbExclamation + vbOKOnly, "Empty section"
                    cboSection.SetFocus
                    Exit Sub
                ElseIf txtStudentNumber = "" Then
                    MsgBox "Generate new Student Number", vbExclamation + vbOKOnly, "Empty student number"
                    txtStudentNumber.SetFocus
                    Exit Sub
                End If
            End If
        
        Case "delete"
            If MsgBox("Do you want to DELETE " & lsvStudentSection.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete record") = vbYes Then
                cn.Execute "DELETE FROM tblAssignSection WHERE iAssignID=" & lsvStudentSection.SelectedItem.Text
                Call LoadlsvStudentSection
            End If
        
        Case "print"
            'Load frmPerSection
            'frmPerSection.Show 1
            'Unload Me
            Call printpersection
    End Select
End Sub

Function complete()
    If cboSection.ListIndex = -1 Or txtStudentNumber = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Sub LoadlsvStudentSection()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblAssignSection"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvStudentSection.ListItems.Clear
    
    If rs.EOF Then
        lsvStudentSection.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvStudentSection.ListItems.Add(, , !iAssignID)
                X.SubItems(1) = !sStudentNumber
                X.SubItems(2) = !sStudentName
                X.SubItems(3) = !sSectionName
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sSchoolYearName
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
        strSQL = "SELECT * FROM tblAssignSection "
        strSQL = strSQL & "WHERE sStudentName LIKE'" & txtSearch.Text & "%'"
        
        rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudentSection.ListItems.Clear
        With rs1
            Do While Not rs1.EOF
                Set X = lsvStudentSection.ListItems.Add(, , !iAssignID)
                X.SubItems(1) = !sStudentNumber
                X.SubItems(2) = !sStudentName
                X.SubItems(3) = !sSectionName
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sSchoolYearName

                .MoveNext
            Loop
        End With
        Set rs1 = Nothing
    ElseIf optLevel.Value = True Then
        strSQL = "SELECT * FROM tblAssignSection "
        strSQL = strSQL & "WHERE sSectionName LIKE'" & txtSearch.Text & "%'"
        
        rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
        
        lsvStudentSection.ListItems.Clear
        With rs2
            Do While Not rs2.EOF
                Set X = lsvStudentSection.ListItems.Add(, , !iAssignID)
                X.SubItems(1) = !sStudentNumber
                X.SubItems(2) = !sStudentName
                X.SubItems(3) = !sSectionName
                X.SubItems(4) = !sStudentSex
                X.SubItems(5) = !sSchoolYearName

                .MoveNext
            Loop
        End With
        Set rs2 = Nothing
    End If
End Sub
Sub printpersection()
    row = 10
    col = 10
    With frmAssign
        Call set_myprinterobject
        set_reportpath = App.Path & "\Reports\PerSection.xls"
        report_gen.Workbooks.Open (set_reportpath)
        
        report_gen.Worksheets(1).Cells(4, 1) = Trim(.txtSection.Text) & " - S.Y. " & Trim(.txtSchoolYear.Text)
        
        i = 8
        j = 1
        
        Do While j <= lsvStudentSection.ListItems.Count
            report_gen.Range(report_gen.Worksheets(1).Cells(i, 2), report_gen.Worksheets(1).Cells(i, 4)).Copy report_gen.Worksheets(1).Cells(i + 1, 2)
            
            With lsvStudentSection
                report_gen.Worksheets(1).Cells(i, 1) = j
                report_gen.Worksheets(1).Cells(i, 2) = .ListItems.Item(j).SubItems(1)
                report_gen.Worksheets(1).Cells(i, 3) = .ListItems.Item(j).SubItems(2)
                report_gen.Worksheets(1).Cells(i, 4) = .ListItems.Item(j).SubItems(4)
            End With
            j = j + 1
            i = i + 1
        Loop
        report_gen.Visible = True
        report_gen.ActiveWindow.SelectedSheets.PrintPreview
        report_gen.ActiveWindow.Close (False)
        report_gen.Quit
    End With
        Set report_gen = Nothing
End Sub
