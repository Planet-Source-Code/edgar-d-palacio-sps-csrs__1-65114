VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShortcutKeys 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPS - Computerized School Registration Software"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ShortcutKeys.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5295
      Top             =   525
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
            Picture         =   "ShortcutKeys.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4530
      Top             =   405
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
            Picture         =   "ShortcutKeys.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShortcutKeys.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ShortcutKeys.frx":0E58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   765
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5583
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Shortcut Keys"
      TabPicture(0)   =   "ShortcutKeys.frx":11F2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lsvShortcuts"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Add New Shortcut"
      TabPicture(1)   =   "ShortcutKeys.frx":120E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtKey"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtDescription"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Toolbar1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtdummy"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtdummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69930
         TabIndex        =   9
         Top             =   465
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   -74880
         TabIndex        =   8
         Top             =   2235
         Width           =   6435
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   -69840
         TabIndex        =   7
         Top             =   2475
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "edit"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   1275
         Left            =   -73095
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   885
         Width           =   4530
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73095
         TabIndex        =   5
         Top             =   540
         Width           =   2370
      End
      Begin MSComctlLib.ListView lsvShortcuts 
         Height          =   2610
         Left            =   120
         TabIndex        =   1
         Top             =   435
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4604
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Shortcut Key"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   12347
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   -74220
         TabIndex        =   3
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Key Combination:"
         Height          =   195
         Left            =   -74730
         TabIndex        =   2
         Top             =   585
         Width           =   1545
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Shortcuts Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1005
      TabIndex        =   4
      Top             =   195
      Width           =   2385
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "ShortcutKeys.frx":122A
      Top             =   45
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "ShortcutKeys.frx":1EF4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6945
   End
End
Attribute VB_Name = "frmShortcutKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
        Call forminit
        
End Sub

Sub forminit()
    Call loadlsvshortcuts
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(3).Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Toolbar1.Buttons(1).Enabled = True
        Toolbar1.Buttons(3).Enabled = False
        txtKey = ""
        txtDescription = ""
    End If
End Sub

Private Sub lsvShortcuts_DblClick()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    row = lsvShortcuts.SelectedItem.Index
    dummy = lsvShortcuts.ListItems.Item(row).Text
    
    Toolbar1.Buttons(1).Enabled = False
    Toolbar1.Buttons(3).Enabled = True
    SSTab1.Tab = 1
    strSQL = "SELECT * FROM tblKeyboard "
    strSQL = strSQL & "WHERE iKeyboardID=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtdummy = dummy
        txtKey = !sKeyboardName
        txtDescription = !sDescription
        
    End With
    Set rs = Nothing
End Sub

Private Sub lsvShortcuts_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo err_handler:
    Select Case KeyCode
        Case vbKeyDelete
            If MsgBox("Do you want to DELETE " & lsvShortcuts.SelectedItem.SubItems(1), vbYesNo + vbQuestion, "Delete shortcut") = vbYes Then
                cn.Execute "DELETE FROM tblkeyboard WHERE ikeyboardID=" & lsvShortcuts.SelectedItem.Text
                Call loadlsvshortcuts
            End If
    End Select
    Exit Sub
err_handler:
    MsgBox Err.Description, vbExclamation, "SPS"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim rsKeyboard As New ADODB.Recordset
            Dim strSQLKeyboard As String
            
            strSQLKeyboard = "SELECT * FROM tblKeyboard"
            
            rsKeyboard.Open strSQLKeyboard, cn, adOpenDynamic, adLockOptimistic
            If txtKey <> "" Then
                With rsKeyboard
                    .AddNew
                        !sKeyboardName = txtKey
                        !sDescription = txtDescription
                        .Update
                End With
                Set rsKeyboard = Nothing
                MsgBox "New shortcut added to the database", vbInformation, "Add new keyboard shortcut"
                txtKey = ""
                txtDescription = ""
                SSTab1.Tab = 0
                Call loadlsvshortcuts
            Else
                If txtKey = "" Then
                    MsgBox "Please type a key combination", vbExclamation, "Key combination is empty"
                    txtKey.SetFocus
                    Exit Sub
                End If
            End If
        Case "edit"
            Dim rsKeyboarde As New ADODB.Recordset
            Dim strSQLKeyboarde As String
            
            strSQLKeyboarde = "SELECT * FROM tblKeyboard"
            strSQLKeyboarde = strSQLKeyboarde & " WHERE iKeyboardID=" & txtdummy
            
            rsKeyboarde.Open strSQLKeyboarde, cn, adOpenDynamic, adLockOptimistic
            
            If txtKey <> "" Then
            
                With rsKeyboarde
                    !sKeyboardName = txtKey
                    !sDescription = txtDescription
                    .Update
                End With
                    Set rsKeyboarde = Nothing
                    MsgBox "The changes you made was successfully updated", vbInformation, "Edit keyboard shortcut"
                    txtKey = ""
                    txtDescription = ""
                    SSTab1.Tab = 0
                    Call loadlsvshortcuts
                
            Else
                If txtKey = "" Then
                    MsgBox "Please type a key combination", vbExclamation, "Key combination is empty"
                    txtKey.SetFocus
                    Exit Sub
                End If
            End If
        Case "delete"
        
    End Select
End Sub

Sub loadlsvshortcuts()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblKeyboard"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvShortcuts.ListItems.Clear
    
    If rs.EOF Then
        lsvShortcuts.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvShortcuts.ListItems.Add(, , !iKeyboardID, , 1)
                X.SubItems(1) = !sKeyboardName
                X.SubItems(2) = !sDescription
                .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub
