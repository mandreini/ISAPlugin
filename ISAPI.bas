Attribute VB_Name = "Module1"
Option Explicit

Function PISA(h As Double) As Double

'Assumes no temperature deviation
Dim P1 As Double
Dim P2 As Double
Dim T1 As Double
Dim T2 As Double
Dim a As Double
Dim g As Double
Dim R As Double

R = 287.05
g = 9.80665

    Select Case h
        Case Is <= 11000:
            a = -0.0065
            T1 = 288.15
            T2 = T1 + (a * h)
            P1 = 101325
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
        Case Is <= 20000:
            a = 0
            T1 = 216.65
            P1 = 22631.7
            P2 = P1 * Exp((-g / (R * T1)) * (h - 11000))
        Case Is <= 32000:
            a = 0.001
            T1 = 216.65
            T2 = T1 + (a * (h - 20000))
            P1 = 5474.717
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
        Case Is <= 47000:
            a = 0.0028
            T1 = 228.65
            T2 = T1 + (a * (h - 32000))
            P1 = 867.974
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
        Case Is <= 51000:
            a = 0
            T1 = 270.65
            P1 = 110.898
            P2 = P1 * Exp((-g / (R * T1)) * (h - 47000))
        Case Is <= 71000:
            a = -0.0028
            T1 = 270.65
            T2 = T1 + (a * (h - 51000))
            P1 = 66.9335
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
        Case Is <= 84000:
            a = -0.002
            T1 = 214.65
            T2 = T1 + (a * (h - 71000))
            P1 = 3.95598
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
        Case Else
            P2 = 0
    End Select

PISA = P2

End Function

Function RISA(h As Double) As Double

'Assumes no temperature deviation
Dim P1 As Double
Dim P2 As Double
Dim T1 As Double
Dim T2 As Double
Dim a As Double
Dim g As Double
Dim R As Double
Dim rho As Double

R = 287.05
g = 9.80665

    Select Case h
        Case Is <= 11000:
            a = -0.0065
            T1 = 288.15
            T2 = T1 + (a * h)
            P1 = 101325
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
            rho = P2 / (R * T2)
        Case Is <= 20000:
            a = 0
            T1 = 216.65
            P1 = 22631.7
            P2 = P1 * Exp((-g / (R * T1)) * (h - 11000))
            rho = P2 / (R * T1)
        Case Is <= 32000:
            a = 0.001
            T1 = 216.65
            T2 = T1 + (a * (h - 20000))
            P1 = 5474.717
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
            rho = P2 / (R * T2)
        Case Is <= 47000:
            a = 0.0028
            T1 = 228.65
            T2 = T1 + (a * (h - 32000))
            P1 = 867.974
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
            rho = P2 / (R * T2)
        Case Is <= 51000:
            a = 0
            T1 = 270.65
            P1 = 110.898
            P2 = P1 * Exp((-g / (R * T1)) * (h - 47000))
            rho = P2 / (R * T1)
        Case Is <= 71000:
            a = -0.0028
            T1 = 270.65
            T2 = T1 + (a * (h - 51000))
            P1 = 66.9335
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
            rho = P2 / (R * T2)
        Case Is <= 84000:
            a = -0.002
            T1 = 214.65
            T2 = T1 + (a * (h - 71000))
            P1 = 3.95598
            P2 = P1 * ((T2 / T1) ^ (-g / (a * R)))
            rho = P2 / (R * T2)
        Case Else
            rho = 0
    End Select

RISA = rho

End Function


Function TISA(h As Double) As Double

'Assumes no temperature deviation
Dim P1 As Double
Dim P2 As Double
Dim T1 As Double
Dim T2 As Double
Dim a As Double
Dim g As Double
Dim R As Double

R = 287.05
g = 9.80665

    Select Case h
        Case Is <= 11000:
            a = -0.0065
            T1 = 288.15
            T2 = T1 + (a * h)
        Case Is <= 20000:
            a = 0
            T1 = 216.65
            T2 = T1
        Case Is <= 32000:
            a = 0.001
            T1 = 216.65
            T2 = T1 + (a * (h - 20000))
        Case Is <= 47000:
            a = 0.0028
            T1 = 228.65
            T2 = T1 + (a * (h - 32000))
        Case Is <= 51000:
            a = 0
            T1 = 270.65
            T2 = T1
        Case Is <= 71000:
            a = -0.0028
            T1 = 270.65
            T2 = T1 + (a * (h - 51000))
        Case Is <= 84000:
            a = -0.002
            T1 = 214.65
            T2 = T1 + (a * (h - 71000))
        
    End Select

TISA = T2

End Function

