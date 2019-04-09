VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRotationAxis 
   Caption         =   "Rotation Axis"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3915
   OleObjectBlob   =   "frmRotationAxis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRotationAxis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xRot, yRot, zRot As Integer

'singleton pattern similar to how MSForms work

Private Alpha As Name
Private Beta As Name
Private Gamma As Name

' Research VBA AUTO INIT
Public Sub ShowInit(nrAlpha As Name, nrBeta As Name, nrGamma As Name)
    Set Alpha = nrAlpha
    Set Beta = nrBeta
    Set Gamma = nrGamma
    Me.Show False
    
    Me.txtXRot.Value = Alpha.RefersToRange.Value2
    Me.ScrollBarX.Value = Alpha.RefersToRange.Value2 + 180
    Me.lblXScrollRot.Caption = Alpha.RefersToRange.Value2
    
    
    Me.txtYRot.Value = Beta.RefersToRange.Value2
    Me.ScrollBarY.Value = Beta.RefersToRange.Value2 + 180
    Me.lblYScrollRot.Caption = Beta.RefersToRange.Value2
    
    
    Me.txtZRot.Value = Gamma.RefersToRange.Value2
    Me.ScrollBarZ.Value = Gamma.RefersToRange.Value2 + 180
    Me.lblZScrollRot.Caption = Gamma.RefersToRange.Value2

    
End Sub

Private Sub ScrollBarX_Change()
    Me.txtXRot = Me.ScrollBarX.Value - 180
    Me.lblXScrollRot.Caption = Me.txtXRot
    xRot = Me.txtXRot
    Alpha.RefersToRange.Value2 = xRot
    'ThisWorkbook.Sheets("Support").Range("AlphaDeg").Value = xRot

End Sub

Private Sub ScrollBarY_Change()
    Dim y As Integer

    Me.txtYRot = Me.ScrollBarY.Value - 180
    Me.lblYScrollRot.Caption = Me.txtYRot
    yRot = Me.txtYRot
    Beta.RefersToRange.Value2 = yRot
    'ThisWorkbook.Sheets("Support").Range("BetaDeg").Value = yRot

End Sub

Private Sub ScrollBarZ_Change()
    Dim y As Integer

    Me.txtZRot = Me.ScrollBarZ.Value - 180
    Me.lblZScrollRot.Caption = Me.txtZRot
    zRot = Me.txtZRot
    Gamma.RefersToRange.Value2 = xRot
    'ThisWorkbook.Sheets("Support").Range("GammaDeg").Value = zRot

End Sub

Private Sub txtXRot_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim x As Variant

    Me.lblXScrollRot.Caption = Me.txtXRot
    x = Me.txtXRot
    
    If Trim(x) = "" Then
    
        Me.txtXRot = xRot                        'restore the last value
        Me.lblXScrollRot = xRot
        Exit Sub
        
    Else
        If Not IsNumeric(x) Then                 ' Replace with the previous value
            Me.ScrollBarX.Value = xRot + 180
            MsgBox "Invalid data Entry"
            Me.txtXRot = xRot
            Me.lblXScrollRot.Caption = Me.txtXRot
            
            Exit Sub
        End If
    
    End If
    Me.ScrollBarX.Value = x + 180
    xRot = x
    ThisWorkbook.Sheets("Support").Range("AlphaDeg").Value = x
    Exit Sub

End Sub

Private Sub txtYRot_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim y As Variant

    Me.lblYScrollRot.Caption = Me.txtYRot
    y = Me.txtYRot
    
    If Trim(y) = "" Then
    
        Me.txtYRot = yRot                        'restore the last value
        Me.lblYScrollRot = yRot
        Exit Sub
        
    Else
        If Not IsNumeric(y) Then                 ' Replace with the previous value
            Me.ScrollBarY.Value = yRot + 180
            MsgBox "Invalid data Entry"
            Me.txtYRot = yRot
            Me.lblYScrollRot.Caption = Me.txtYRot
            
            Exit Sub
        End If
    
    End If
    Me.ScrollBarY.Value = y + 180
    yRot = y
    Exit Sub

End Sub

Private Sub txtZRot_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim z As Variant

    Me.lblZScrollRot.Caption = Me.txtZRot
    z = Me.txtZRot
    
    If Trim(z) = "" Then
    
        Me.txtZRot = zRot                        'restore the last value
        Me.lblZScrollRot = zRot
        Exit Sub
        
    Else
        If Not IsNumeric(z) Then                 ' Replace with the previous value
            Me.ScrollBarZ.Value = zRot + 180
            MsgBox "Invalid data Entry"
            Me.txtZRot = zRot
            Me.lblZScrollRot.Caption = Me.txtZRot
            
            Exit Sub
        End If
    
    End If
    Me.ScrollBarZ.Value = z + 180
    zRot = z
    Exit Sub

End Sub

