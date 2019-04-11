Attribute VB_Name = "Ribbon"
Option Explicit
Private ThreeDViewerRibbon As IRibbonUI

Public Sub ThreeDviewerRibbonctrl(Ribbon As IRibbonUI)
    '
    ' Code for onLoad callback. Ribbon control customUI
    '
    Set ThreeDViewerRibbon = Ribbon

End Sub

Public Sub BtnTopView_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button

    PredefineViews "frmRotationAxis", 0, 0, 0

End Sub

Public Sub BtnFrontView_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
   
    PredefineViews "frmRotationAxis", 90, -180, 90
  
End Sub

Public Sub BtnLeftView_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '

    PredefineViews "frmRotationAxis", -90, 0, 0
  
End Sub

Public Sub BtnRightView_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '

    PredefineViews "frmRotationAxis", -90, 0, -180
  
End Sub

Public Sub BtnIso30_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
    PredefineViews "frmRotationAxis", -45, 0, 30
  
End Sub

Public Sub BtnIso45_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
    PredefineViews "frmRotationAxis", -45, 0, 45
  
End Sub

Public Sub BtnShowRotationTool_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
    'frmRotationAxis.Show False
    
    PrepareForm
End Sub

Public Sub About_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
    frmMITLicence.Show
End Sub

Public Sub Contact_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    OpenUrl "https://www.linkedin.com/in/jean-marc-flamand-79592422/"

End Sub

Public Sub GitHub_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    OpenUrl "https://github.com/JeanMarcFlamand/VBA--3DMPViewer"

End Sub

Public Sub BtnReActivatePointFinderBtn_onAction(control As IRibbonControl)
    '
    ' Code for onAction callback. Ribbon control button
    '
    SelectChart ThisWorkbook.Sheets("Data")

End Sub

Sub PredefineViews(myform As String, xDeg As Integer, yDeg As Integer, zdeg As Integer)

    If IsUserFormLoaded(myform) Then
        'The form is already open
        PredefineRotation xDeg, yDeg, zdeg
    Else
        'Reopen the form"
        PrepareForm
        PredefineRotation xDeg, yDeg, zdeg
    End If
End Sub

Sub PredefineRotation(xDeg As Integer, yDeg As Integer, zdeg As Integer)

    'Update value in the worksheet
    ThisWorkbook.Sheets("Support").Range("AlphaDeg").Value = xDeg
    ThisWorkbook.Sheets("Support").Range("BetaDeg").Value = yDeg
    ThisWorkbook.Sheets("Support").Range("GammaDeg").Value = zdeg
    
  
    
    'Update value on the form
    frmRotationAxis.txtXRot.Value = xDeg
    frmRotationAxis.ScrollBarX.Value = xDeg + 180
    frmRotationAxis.lblXScrollRot.Caption = xDeg
       
    frmRotationAxis.txtYRot.Value = yDeg
    frmRotationAxis.ScrollBarY.Value = yDeg + 180
    frmRotationAxis.lblYScrollRot.Caption = yDeg
       
    frmRotationAxis.txtZRot.Value = zdeg
    frmRotationAxis.ScrollBarZ.Value = zdeg + 180
    frmRotationAxis.lblZScrollRot.Caption = zdeg
       
 



End Sub

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object

    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function                                     'IsUserFormLoaded

