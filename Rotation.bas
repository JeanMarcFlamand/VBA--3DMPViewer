Attribute VB_Name = "Rotation"
Option Explicit

Dim p1, P2, Q1, Q2, R1, R2 As Double

 


Function AlphaRad(Degree As Double) As Variant
    
    AlphaRad = 2 * (WorksheetFunction.Pi() / 360) * Degree
End Function

Function BetaRad(Degree As Double) As Variant
    
    BetaRad = 2 * (WorksheetFunction.Pi() / 360) * Degree
End Function

Function GammaRad(Degree As Double) As Variant
    
    GammaRad = 2 * (WorksheetFunction.Pi() / 360) * Degree
End Function

Function Pone()

    Pone = Cos(2 * (WorksheetFunction.Pi() / 360) * Application.Range("BetaDeg").Value) * Cos(2 * (WorksheetFunction.Pi() / 360) * Application.Range("GammaDeg").Value)
End Function

