Function moody(ByVal Diam As Double, ByVal EPS As Double, ByVal rho As Single, ByVal my As Single, ByVal Q As Double) As Double
' -----------------------------------------------------------------------------------------
' Function to calculate the pipe flow friction factor by using newton-iteration on the colebrook white function (Moody's friction factor)
' Input: Darcy(Diameter; EPS=rør-roughness [m]; rho=density [kg/m3]; my=viscosity [kg/ms]; Q=flow [m3/s])
' Equation: Colebrook White: Gfunct = 0 = f^-0.5 + 2log( a + b/(Q * f^0.5) )
'                           GfunctDeriv = 0 = -0.5*f^-1.5 - (b / (a * f^1.5 * Q + b*f))
'
' KRISTIAN HOLM JENSEN - mailto: Kristian.Holm.Jensen@gmail.com
' ------------------------------------------------------------------------------------------

'--- dedikerer variable
Dim i As Integer
Dim f0, f1 As Double
Dim diff As Double
Dim NRe As Double
Dim Gfunct, GfunctDeriv As Double
Dim a, b As Double

On Error GoTo errhandler
Q = Abs(Q)

a = EPS / Diam / 3.7
b = 2.51 * WorksheetFunction.pi() * Diam * my / (4 * rho)

NRe = 4 * rho * Q / (WorksheetFunction.pi() * Diam * my)

f0 = 64 / NRe
diff = 1 'setter diff til 1 for å starte whileløkke

If NRe < 2100 Then ' Tester om vi har turbulent eller laminær flowtrømning
    f1 = 64 / NRe ' hvis NRe < 2100 har vi laminær strømning, dvs friksjonsfaktoren = 64 / Reynoldstallet

Else
    i = 0
        Do While diff > 0.00000001 'Konvergerer til differanse på 1e-8
            Gfunct = (f0 ^ -0.5) + 2 * WorksheetFunction.Log10(a + b / (Q * f0 ^ 0.5))
            GfunctDeriv = -0.5 * f0 ^ (-1.5) - (b / (a * f0 ^ 1.5 * Q + b * f0))
            f1 = f0 - Gfunct / GfunctDeriv 'NEWTON ITERASJON
            
            diff = Abs(f0 - f1) 'diff er differanse mellom iterasjon n og n+1.
            f0 = f1 'setter neste startverdi.
            
                If i > 100 Then
                    MsgBox "Moody did not converge. F-factor =  " & f1
                    moody = 64 / NRe 'Returns the basic friction factor for laminar flow (64/NRe)
                    Exit Function 'Prevents infinite loop
                End If
            i = i + 1
        Loop

End If

moody = f1

Exit Function

errhandler:
'MsgBox "Colebrook did not converge: f set to 64/NRe "
moody = 64 / NRe

End Function
