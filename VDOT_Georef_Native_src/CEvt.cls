VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event sink for one runtime-created MSForms control; forwards to modUI / modView.
Option Explicit

Public Name As String
Public WithEvents eLbl As MSForms.Label
Attribute eLbl.VB_VarHelpID = -1
Public WithEvents eChk As MSForms.CheckBox
Attribute eChk.VB_VarHelpID = -1
Public WithEvents eOpt As MSForms.OptionButton
Attribute eOpt.VB_VarHelpID = -1
Public WithEvents eCbo As MSForms.ComboBox
Attribute eCbo.VB_VarHelpID = -1
Public WithEvents eTxt As MSForms.TextBox
Attribute eTxt.VB_VarHelpID = -1
Public WithEvents eLst As MSForms.ListBox
Attribute eLst.VB_VarHelpID = -1
Public WithEvents eScr As MSForms.ScrollBar
Attribute eScr.VB_VarHelpID = -1
Public WithEvents eImg As MSForms.Image
Attribute eImg.VB_VarHelpID = -1
Public WithEvents eFrm As MSForms.Frame
Attribute eFrm.VB_VarHelpID = -1

Private Sub eLbl_Click()
    modUI.OnEvent Name, "click"
End Sub

Private Sub eLbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modUI.OnHover Name
End Sub

Private Sub eChk_Click()
    modUI.OnEvent Name, "click"
End Sub

Private Sub eChk_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modUI.OnHover ""
End Sub

Private Sub eOpt_Click()
    modUI.OnEvent Name, "click"
End Sub

Private Sub eCbo_Change()
    modUI.OnEvent Name, "change"
End Sub

Private Sub eTxt_Change()
    modUI.OnEvent Name, "change"
End Sub

Private Sub eLst_Click()
    modUI.OnEvent Name, "click"
End Sub

Private Sub eLst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    modUI.OnEvent Name, "dblclick"
End Sub

Private Sub eScr_Change()
    modUI.OnEvent Name, "change"
End Sub

Private Sub eScr_Scroll()
    modUI.OnEvent Name, "change"
End Sub

Private Sub eImg_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modView.ViewMouseDown Button, Shift, X, Y
End Sub

Private Sub eImg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modUI.OnHover ""
    modView.ViewMouseMove Button, Shift, X, Y
End Sub

Private Sub eImg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modView.ViewMouseUp Button, Shift, X, Y
End Sub

Private Sub eFrm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    modUI.OnHover ""
End Sub
