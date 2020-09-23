VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuarterUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Update Quarter"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QuarterUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   780
      TabIndex        =   3
      Top             =   30
      Width           =   4170
      Begin VB.TextBox txtQuarterDummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1155
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtQuarter 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1515
         TabIndex        =   0
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1515
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quarter Name:"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   285
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrition:"
         Height          =   195
         Left            =   465
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
            Picture         =   "QuarterUpdate.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "QuarterUpdate.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   60
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
Attribute VB_Name = "frmQuarterUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtQuarter = ""
        txtDescription = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call forminit
End Sub

Sub forminit()
    Call PositionForm1(frmQuarterUpdate)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "update"
            Dim strSQL As String
            
            If frmQuarter.dummyButton = "add" Then
                
                strSQL = "SELECT * FROM tblQuarters"
                
                If txtQuarter <> "" Then
                    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs
                        .AddNew
                        !sQuarterName = txtQuarter
                        !sQuarterNotes = txtDescription
                        .Update
                    End With
                    MsgBox "New record added to the database", vbInformation + vbOKOnly, "Add new quarter"
                    Unload Me
                    Set rs = Nothing
                Else
                    MsgBox "Please enter QUARTER NAME", vbExclamation + vbOKOnly, "Empty quarter name"
                    Exit Sub
                End If
                
            ElseIf frmQuarter.dummyButton = "edit" Then
            
                strSQL = "SELECT * FROM tblQuarters "
                strSQL = strSQL & "WHERE iQuarterID=" & txtQuarterDummy
                
                If txtQuarter <> "" Then
                    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
                    With rs1
                        !sQuarterName = txtQuarter
                        !sQuarterNotes = txtDescription
                        .Update
                    End With
                    MsgBox "The changes you made was successfully updated", vbInformation + vbOKOnly, "Update quarter"
                    Unload Me
                    Set rs1 = Nothing
                Else
                    MsgBox "Please enter QUARTER NAME", vbExclamation + vbOKOnly, "Empty quarter name"
                End If
            End If
    
        Case "back"
            Unload Me
    End Select
End Sub

