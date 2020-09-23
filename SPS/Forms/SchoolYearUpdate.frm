VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchoolYearUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Update School Year"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SchoolYearUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   855
      TabIndex        =   0
      Top             =   60
      Width           =   4005
      Begin VB.TextBox txtSchoolYearDummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1125
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1335
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   630
         Width           =   2490
      End
      Begin VB.TextBox txtSchoolYear 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1335
         TabIndex        =   2
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrition:"
         Height          =   195
         Left            =   330
         TabIndex        =   3
         Top             =   630
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "School Year:"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   285
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   765
      Top             =   1830
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
            Picture         =   "SchoolYearUpdate.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SchoolYearUpdate.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   765
      _ExtentX        =   1349
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
Attribute VB_Name = "frmSchoolYearUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtSchoolYear = ""
        txtDescription = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call forminit
End Sub

Sub forminit()
    Call PositionForm1(frmSchoolYearUpdate)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
            
            If frmSchoolYear.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblSchoolYears"
                
                If txtSchoolYear <> "" Then
                    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs
                        .AddNew
                        !sSchoolYearName = txtSchoolYear
                        !sSchoolYearNotes = txtDescription
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new school year"
                    Unload Me
                    Set rs = Nothing
                Else
                    MsgBox "Please enter SCHOOL YEAR", vbExclamation + vbOKOnly, "Empty school year"
                    Exit Sub
                End If
                
            ElseIf frmSchoolYear.dummyButton = "edit" Then
            
                strSQL = "SELECT * FROM tblSchoolYears "
                strSQL = strSQL & "WHERE iSchoolYearID=" & txtSchoolYearDummy
                
                If txtSchoolYear <> "" Then
                    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs1
                        !sSchoolYearName = txtSchoolYear
                        !sSchoolYearNotes = txtDescription
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update school year"
                    Unload Me
                    Set rs1 = Nothing
                Else
                    MsgBox "Please enter SCHOOL YEAR", vbExclamation + vbOKOnly, "Empty school year"
                End If
            End If
    
        Case "back"
            Unload Me
    End Select
End Sub
