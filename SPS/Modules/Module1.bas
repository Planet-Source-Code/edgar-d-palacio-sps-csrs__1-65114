Attribute VB_Name = "Module1"
Option Explicit
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rs4 As New ADODB.Recordset
Public dummyButton As String
Public lst As ListItem
Public set_reportpath As String, row, col As Integer
Public report_gen As Excel.Application
'Global rpt_header As report_header
'Database connection
Public Sub DBConnect()
    On Error GoTo err_handler:
    'school connection
    cn.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=CSMS"
    'home connection
    'cn.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SPSCSRS"
    Exit Sub
err_handler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Database Connection Error"
    Exit Sub
End Sub
Public Sub DBClose()
    On Error GoTo err_handler:
    cn.Close
    Set cn = Nothing
err_handler:
    Exit Sub
End Sub
'Procedure for text highligh
Public Sub Highlight(ByRef sText As TextBox)
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub
'Procedure to center the form
Public Sub CenterForm(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = (Screen.Height - frm.Height) \ 2
    LeftCorner = (Screen.Width - frm.Width) \ 2
    frm.Move LeftCorner, TopCorner
End Sub
Public Sub PositionForm(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = 2000
    LeftCorner = 900
    frm.Move LeftCorner, TopCorner
End Sub
Public Sub PositionForm1(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = 5270
    LeftCorner = 5000
    frm.Move LeftCorner, TopCorner
End Sub
Public Sub PositionForm2(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = 4000
    LeftCorner = 3000
    frm.Move LeftCorner, TopCorner
End Sub
Public Sub PositionForm3(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = 3180
    LeftCorner = 2300
    frm.Move LeftCorner, TopCorner
End Sub
Function ListFindItem(lstCtrl As Control, lngSearch As Long) As Integer
   'just returns the position, does not set it
   'used to see if item is in list
   Dim intLen As Integer
   Dim intLoop As Integer
   Dim intPos As Integer

   intLen = lstCtrl.ListCount - 1
   intPos = -1
   For intLoop = 0 To intLen
      If lstCtrl.ItemData(intLoop) = lngSearch Then
         intPos = intLoop
         Exit For
      End If
   Next intLoop
   ListFindItem = intPos
End Function
Public Function set_myprinterobject()
    Set report_gen = New Excel.Application
End Function
