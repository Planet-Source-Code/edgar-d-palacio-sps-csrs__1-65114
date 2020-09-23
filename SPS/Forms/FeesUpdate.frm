VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeesUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Update School Fees"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FeesUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   855
      TabIndex        =   6
      Top             =   45
      Width           =   4740
      Begin VB.ComboBox cboLevel 
         Height          =   315
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   611
         Width           =   2235
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1395
         TabIndex        =   3
         Top             =   1338
         Width           =   1335
      End
      Begin VB.TextBox txtFeesDummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   180
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox cboSchoolYear 
         Height          =   315
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1395
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1695
         Width           =   3210
      End
      Begin VB.TextBox txtFeeName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1395
         TabIndex        =   2
         Top             =   982
         Width           =   2955
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   765
         TabIndex        =   12
         Top             =   671
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Amount:"
         Height          =   195
         Left            =   555
         TabIndex        =   11
         Top             =   1391
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "School Year:"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrition:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1695
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name of Fee:"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   1035
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   645
      Top             =   1785
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
            Picture         =   "FeesUpdate.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FeesUpdate.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   645
      _ExtentX        =   1138
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
End
Attribute VB_Name = "frmFeesUpdate"
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
        cboSchoolYear.ListIndex = -1
        cboLevel.ListIndex = -1
        txtFeeName = ""
        txtAmount = ""
        txtDescription = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call PositionForm1(frmFeesUpdate)
End Sub
Sub forminit()
    
    Call loadcboSchoolYear
    Call loadcboLevel
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
            
            
            If frmFees.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblFees"
                
                
                If complete = True Then
                rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs
                        .AddNew
                        !iSchoolYearID = cboSchoolYear.ItemData(cboSchoolYear.ListIndex)
                        !iLevelID = cboLevel.ItemData(cboLevel.ListIndex)
                        !sFeeName = txtFeeName
                        !cAmount = txtAmount
                        !sFeeNotes = txtDescription
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new school fee"
                    txtFeeName = ""
                    txtAmount = ""
                    txtDescription = ""
                    txtFeeName.SetFocus
                    'Unload Me
                    Set rs = Nothing
                Else
                    If cboSchoolYear = "" Then
                        MsgBox "Please select SCHOOL YEAR", vbExclamation + vbOKOnly, "Empty School Year"
                        cboSchoolYear.SetFocus
                        Exit Sub
                    ElseIf cboLevel = "" Then
                        MsgBox "Please select LEVEL", vbExclamation + vbOKOnly, "Empty Level"
                        cboLevel.SetFocus
                        Exit Sub
                    ElseIf txtFeeName = "" Then
                        MsgBox "Please type the NAME of fee", vbExclamation + vbOKOnly, "Empty fee name"
                        txtFeeName.SetFocus
                        Exit Sub
                    ElseIf txtAmount = "" Then
                        MsgBox "Please type the FEE AMOUNT", vbExclamation + vbOKOnly, "Empty amount"
                        txtAmount.SetFocus
                        Exit Sub
                    End If
                End If
                
            ElseIf frmFees.dummyButton = "edit" Then
                strSQL = "SELECT * FROM tblFees"
                strSQL = strSQL & " WHERE iFeeID=" & txtFeesDummy
                
                If complete = True Then
                rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs3
                        !iSchoolYearID = cboSchoolYear.ItemData(cboSchoolYear.ListIndex)
                        !iLevelID = cboLevel.ItemData(cboLevel.ListIndex)
                        !sFeeName = txtFeeName
                        !cAmount = txtAmount
                        !sFeeNotes = txtDescription
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update school fee"
                    Unload Me
                    Set rs3 = Nothing
                Else
                    If cboSchoolYear = "" Then
                        MsgBox "Please select SCHOOL YEAR", vbExclamation + vbOKOnly, "Empty School Year"
                        cboSchoolYear.SetFocus
                        Exit Sub
                    ElseIf cboLevel = "" Then
                        MsgBox "Please select LEVEL", vbExclamation + vbOKOnly, "Empty Level"
                        cboLevel.SetFocus
                        Exit Sub
                    ElseIf txtFeeName = "" Then
                        MsgBox "Please type the NAME of fee", vbExclamation + vbOKOnly, "Empty fee name"
                        txtFeeName.SetFocus
                        Exit Sub
                    ElseIf txtAmount = "" Then
                        MsgBox "Please type the FEE AMOUNT", vbExclamation + vbOKOnly, "Empty amount"
                        txtAmount.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case "back"
            Unload Me
    End Select
End Sub

Function complete()
    If cboSchoolYear.ListIndex = -1 Or cboLevel.ListIndex = -1 Or txtFeeName = "" Or txtAmount = "" Then
        complete = False
    Else
        complete = True
    End If
    
End Function
Sub ListCheck()
    Dim litmfound As ListItem
    Set litmfound = frmFees.lsvFees.FindItem(frmFees.txtSearch, 1, , 0)

    If litmfound Is Nothing Then
        
        FindItem = False
    Else
        MsgBox "This FEE is already in the list" & vbCrLf _
               & "Input another FEE", vbCritical + vbOKOnly, "Duplicate Item"
        litmfound.EnsureVisible
        litmfound.Selected = True
        FindItem = True
    End If
End Sub

