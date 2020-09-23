VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmPerSection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Student List Per Section"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   1725
   ClientWidth     =   15270
   Icon            =   "PerSection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15270
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   15255
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmPerSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim Report As New rptSection

Private Sub Form_Load()
    Call loadsectionData
    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = rptSection
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
    CRViewer1.Zoom 90
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub
Private Sub loadsectionData()
    Dim a As Integer
    
        rptSection.txtHeading.SetText frmAssign.txtSection & " - S.Y. " & frmAssign.txtSchoolYear
        MsgBox frmAssign.lsvStudentSection.ListItems.Count
    With rptSection
        For a = 1 To frmAssign.lsvStudentSection.ListItems.Count
            .txtStudentNumber.SetText .txtStudentNumber.Text & vbCrLf & frmAssign.lsvStudentSection.ListItems(a).SubItems(1)
            .txtName.SetText .txtName.Text & vbCrLf & frmAssign.lsvStudentSection.ListItems(a).SubItems(2)
            .txtSex.SetText .txtSex.Text & vbCrLf & frmAssign.lsvStudentSection.ListItems(a).SubItems(4)
        Next a
    End With
'    With SaleReceipt
'         .txtDate.SetText frmSale.txtDate
'         .txtTotal.SetText Format(frmSale.Text2, "P ###,###,###.00")
'         For i = 1 To frmSale.lsvList.ListItems.Count
'             '.txtProductName.SetText .txtProductName.Text & vbCrLf & frmSale.lsvList.ListItems.Item(i).SubItems(1) & " - " & frmSale.lsvList.ListItems(i).SubItems(2)
'             .txtProductName.SetText .txtProductName.Text & vbCrLf & frmSale.lsvList.ListItems(i).SubItems(2) & " - " & frmSale.lsvList.ListItems.Item(i).SubItems(1) & " - " & Format(frmSale.lsvList.ListItems.Item(i).SubItems(3), "P###,###,###.00")
'             '.txtVideo.SetText .txtVideo.Text & vbCrLf & frmRent.lsvList.ListItems.Item(i).SubItems(1) & vbTab & frmRent.lsvList.ListItems.Item(i).SubItems(2)
'         Next i
'         .txtAmtReceived.SetText Format(frmSale.txtAmount, "P ###,###,###.00")
'         .txtChange.SetText Format(frmSale.Text1, "P ###,###,###.00")
'    End With
End Sub
