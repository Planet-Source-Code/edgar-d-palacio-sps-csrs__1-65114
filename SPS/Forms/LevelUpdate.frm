VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLevelUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Remove Level"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LevelUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   855
      TabIndex        =   3
      Top             =   30
      Width           =   4050
      Begin VB.TextBox txtLevelDummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   330
         TabIndex        =   6
         Top             =   1170
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtLevel 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1380
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
         TabIndex        =   1
         Top             =   630
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Gr./Yr. Level:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrition:"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   675
         Width           =   930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   645
      Top             =   1770
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
            Picture         =   "LevelUpdate.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LevelUpdate.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   105
      TabIndex        =   2
      Top             =   90
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
Attribute VB_Name = "frmLevelUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtLevel = ""
        txtDescription = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call forminit
End Sub

Sub forminit()
    Call PositionForm1(frmLevelUpdate)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
            
            If frmLevel.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblLevels"
                
                If txtLevel <> "" Then
                    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs
                        .AddNew
                        !sLevelName = txtLevel
                        !sLevelNotes = txtDescription
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new level"
                    Unload Me
                    Set rs = Nothing
                Else
                    MsgBox "Please enter LEVEL", vbExclamation + vbOKOnly, "Empty level name"
                    Exit Sub
                End If
                
            ElseIf frmLevel.dummyButton = "edit" Then
            
                strSQL = "SELECT * FROM tblLevels "
                strSQL = strSQL & "WHERE iLevelID=" & txtLevelDummy
                
                If txtLevel <> "" Then
                    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs1
                        !sLevelName = txtLevel
                        !sLevelNotes = txtDescription
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update level"
                    Unload Me
                    Set rs1 = Nothing
                Else
                    MsgBox "Please enter SCHOOL YEAR", vbExclamation + vbOKOnly, "Empty level name"
                End If
            End If
    
        Case "back"
            Unload Me
    End Select
End Sub

