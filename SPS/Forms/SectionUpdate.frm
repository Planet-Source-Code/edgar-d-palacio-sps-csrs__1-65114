VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSectionUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Update Section"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SectionUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
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
      Height          =   3150
      Left            =   795
      TabIndex        =   6
      Top             =   0
      Width           =   4965
      Begin VB.TextBox txtMaxAverage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   4
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtMinAverage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   3
         Top             =   1380
         Width           =   870
      End
      Begin VB.TextBox txtAllowed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   2
         Top             =   1005
         Width           =   585
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtSectionDummy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   75
         TabIndex        =   10
         Top             =   2625
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   2310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2130
         Width           =   2490
      End
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2310
         TabIndex        =   1
         Top             =   630
         Width           =   2235
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Max. Average:"
         Height          =   195
         Left            =   960
         TabIndex        =   14
         Top             =   1815
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Min. Average"
         Height          =   195
         Left            =   1095
         TabIndex        =   13
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. of Students Allowed:"
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   1065
         Width           =   2115
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   195
         Left            =   1695
         TabIndex        =   11
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrition:"
         Height          =   195
         Left            =   1290
         TabIndex        =   8
         Top             =   2130
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Section Name:"
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   690
         Width           =   1260
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   615
      Top             =   1740
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
            Picture         =   "SectionUpdate.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SectionUpdate.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   75
      TabIndex        =   9
      Top             =   60
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
Attribute VB_Name = "frmSectionUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cboLevel.ListIndex = -1
        txtSection = ""
        txtAllowed = ""
        txtMinAverage = ""
        txtMaxAverage = ""
        txtDescription = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call forminit
End Sub
Sub forminit()
    Call PositionForm1(frmSectionUpdate)
    Call loadcboLevel
End Sub
Sub loadcboLevel()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblLevels"
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs1
        Do While Not rs1.EOF
            cboLevel.AddItem !sLevelName
            cboLevel.ItemData(cboLevel.NewIndex) = CLng(!iLevelID)
            .MoveNext
        Loop
    End With
    Set rs1 = Nothing
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
                        
            If frmSection.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblSections"
                
                If complete = True Then
                
                    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs2
                        .AddNew
                        !iLevelID = cboLevel.ItemData(cboLevel.ListIndex)
                        !sSectionName = txtSection
                        !iAllowed = Val(txtAllowed)
                        !iMaxAverage = Val(txtMaxAverage)
                        !iMinAverage = Val(txtMinAverage)
                        !sSectionNotes = txtDescription
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new section"
                    Unload Me
                    Set rs2 = Nothing
                    
                Else
                   If cboLevel = "" Then
                        MsgBox "Please select LEVEL", vbExclamation + vbOKOnly, "Empty Level"
                        cboLevel.SetFocus
                        Exit Sub
                    ElseIf txtSection = "" Then
                        MsgBox "Please enter SECTION NAME", vbExclamation + vbOKOnly, "Empty Section Name"
                        txtSection.SetFocus
                        Exit Sub
                    ElseIf txtAllowed = "" Then
                        MsgBox "Please enter ALLOWED", vbExclamation + vbOKOnly, "Empty Allowed Students"
                        txtAllowed.SetFocus
                        Exit Sub
                    ElseIf txtMinAverage = "" Then
                        MsgBox "Pleas enter MINIMUM AVERAGE", vbExclamation + vbOKCancel, "Empty minimum average"
                        txtMinAverage.SetFocus
                        Exit Sub
                    ElseIf txtMaxAverage = "" Then
                        MsgBox "Please enter MAXIMUM AVEARGE", vbExclamation + vbOKCancel, "Empty maximum average"
                        txtMaxAverage.SetFocus
                        Exit Sub
                    End If
                End If
                
            ElseIf frmSection.dummyButton = "edit" Then
            
                strSQL = "SELECT * FROM tblSections "
                strSQL = strSQL & "WHERE iSectionID=" & txtSectionDummy
                
                If complete = True Then
                    rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    
                    With rs3
                        !iLevelID = cboLevel.ItemData(cboLevel.ListIndex)
                        !sSectionName = txtSection
                        !iAllowed = Val(txtAllowed)
                        !iMaxAverage = Val(txtMaxAverage)
                        !iMinAverage = Val(txtMinAverage)
                        !sSectionNotes = txtDescription
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update section"
                    Unload Me
                    Set rs3 = Nothing
                Else
                    If cboLevel = "" Then
                        MsgBox "Please select LEVEL", vbExclamation + vbOKOnly, "Empty Level"
                        cboLevel.SetFocus
                        Exit Sub
                    ElseIf txtSection = "" Then
                        MsgBox "Please enter SECTION NAME", vbExclamation + vbOKOnly, "Empty Section Name"
                        txtSection.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case "back"
            Unload Me
    End Select
End Sub

Function complete()
    If cboLevel.ListIndex = -1 Or txtSection = "" Then
        complete = False
    Else
        complete = True
    End If
    
End Function

