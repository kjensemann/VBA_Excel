Public Function fCubicSplineEng(ByVal x_value As Variant, ByVal x_vect_in As Variant, ByVal y_vect_in As Variant) As Double
'----------------------------------------------------------------
'----- FUNCTION MAKING CUBIC FIT FROM THE THREE CLOSEST POINTS --------
' http://www.korf.co.uk/spline.pdf
' Input: x_value --> f(x) will be calculated based on cubic fit. 
' x_vect_in --> Measured data (in rising order, i.e. 1, 2, 4, 9, 10, 11, 18...
' y_vect_in --> Measured data at "x", e.g. f(2) = 12.3, f(4)=19.2 etc
' NB!!! X_Vector må være i STIGENDE rekkefølge! e.g. x --> 0, 1, 3, 9, 12 etc, not 9,5,3,2,1 !!
'----------------------------------------------------------------

Dim b_int_Low As Integer, b_int_High As Integer, b_int_Act As Integer
Dim bis_Count As Integer
Dim n_test As Integer

Dim x_vect1() ' As Double
Dim y_Vect1() ' As Double
Dim i As Integer, n As Integer, j As Integer
Dim x_vect() As Double
Dim y_vect() As Double
Dim a As Double, b As Double, c As Double, d As Double
Dim dy_dx(0 To 1) As Double
Dim dy_d2x(0 To 1) As Double

Application.Calculation = xlCalculationManual ' TO AVOID DELAYS

x_vect1() = x_vect_in.value
y_Vect1() = y_vect_in.value

If UBound(x_vect1) = 1 Then 'IF HORIZONTAL VECTORS ARE PROVIDED FOR x and y. 
    
    ReDim x_vect(UBound(x_vect1, 2) - 1)
    ReDim y_vect(UBound(y_Vect1, 2) - 1)
    
    If x_vect1(1, 1) - x_vect1(1, 2) < 0 Then 'HVIS X_Vector er oppgitt i Økende rekkefølge, dvs X_vect = [1,2,4,5,5.2,6.1,8] (normalt)
        For i = 0 To UBound(x_vect, 1)
            x_vect(i) = x_vect1(1, i + 1) 'Lager x og y-vektor
            y_vect(i) = y_Vect1(1, i + 1)
        Next i
    Else 'Hvis X_vect1 kommer inn i synkende rekkefølge.
        For i = 0 To UBound(x_vect, 1)
            x_vect(i) = x_vect1(1, UBound(x_vect, 1) - i) 'Lager x og y-vektor
            y_vect(i) = y_Vect1(1, UBound(x_vect, 1) - i)
        Next i
    End If
     
Else 'IF VERTICAL VECTORS ARE PROVIDED FOR x and y. 
    
    ReDim x_vect(UBound(x_vect1, 1) - 1)
    ReDim y_vect(UBound(y_Vect1, 1) - 1)
    
    If x_vect1(1, 1) - x_vect1(2, 1) < 0 Then 'Hvis X_vect1 kommer inn i økende rekkefølge
        For i = 0 To UBound(x_vect, 1)
            x_vect(i) = x_vect1(i + 1, 1)
            y_vect(i) = y_Vect1(i + 1, 1)
        Next i
    Else 'HVIS X_Vect kommer inn i synkende rekkefølge.
        For i = 0 To UBound(x_vect, 1)
            x_vect(i) = x_vect1(UBound(x_vect1, 1) - i, 1)
            y_vect(i) = y_Vect1(UBound(x_vect1, 1) - i, 1)
        Next i
    End If
    
    
End If

i = 0: n = 0

    '------------------- LINEAR EXTRAPOLATION FOR POINTS OUTSIDE THE MEASURED RANGE --------------------------
    If x_value <= x_vect(0) Then 'Sjekker om x er lavere enn laveste (første) x-verdi Y = Bx+A
        n = 0
            If x_vect(n + 1) - x_vect(n) = 0 Then
                b = 1E+30
            Else
                b = (y_vect(n + 1) - y_vect(n)) / (x_vect(n + 1) - x_vect(n))
                a = y_vect(n) - b * x_vect(n)
                khjCubicSplineEng = b * x_value + a
                Exit Function
            End If
    ElseIf x_value >= x_vect(UBound(x_vect)) Then
        n = UBound(x_vect)
            If x_vect(n) - x_vect(n - 1) = 0 Then
                b = 1E+30
            Else
                b = (y_vect(n) - y_vect(n - 1)) / (x_vect(n) - x_vect(n - 1))
                a = y_vect(n) - b * x_vect(n)
                khjCubicSplineEng = b * x_value + a
                Exit Function
            End If
    End If
    
'------------------- ENGINEERING SPLINE INNENFOR MÅLEPUNKTER --------------------------
    'Finner i hvilket intervall vi befinner oss i :
    
i = 0
            
b_int_Low = 0
b_int_High = UBound(x_vect) 'Length of stepwise inclining x_vector
          
    
If b_int_High > 10 Then 'USES BRACKETING IF ARRAY IS LARGE..
    
    Do While i < 200
                                        
        If x_value < x_vect(b_int_High) And x_value > x_vect(b_int_Low) Then
      
        b_int_Act = Application.RoundUp((b_int_High + b_int_Low) / 2, 0) 'Finner midten av array (deler array i to, for så å finne ut om valgt verdi er midt i mellom)
                  
            If b_int_High - b_int_Low = 1 Then 'We have found our point
                n_test = b_int_High
                Exit Do
            ElseIf x_value = x_vect(b_int_Act) Then
                n_test = b_int_Act + 1
                Exit Do
            ElseIf x_value < x_vect(b_int_Act) Then 'And x_value >= x_vect(b_int_Low) Then 'x må ligge i intervallet b_int_act --> b_int_high
                b_int_High = b_int_Act
                       
            ElseIf x_value >= x_vect(b_int_Act) Then 'And x_value < x_vect(b_int_High) Then
               b_int_Low = b_int_Act
            End If
            
        End If
        
        i = i + 1
        bis_Count = bis_Count + 1
        If i > 100 Then
            'MsgBox "Cubic Spline Eng - Bisection did not find interval, ended in 100 iterations"
        End If
        
    Loop
n = n_test
Else 'IF WE DO NOT HAVE MANY ELEMENTS, THEN STEP SEARCH IS USED
  
    For i = 1 To UBound(x_vect)
        If x_value < x_vect(i) Then 'And x_value > x_vect(i - 1)
            n = i 'Setter Intervallnummeret vi befinner oss i
            Exit For
        End If
    Next i
End If

i = 0
'---------- CALCULATES THE 1. DERIVATIVES ----------------------
For j = 0 To 1
    i = n + j - 1 'i er alltid >=1 fra over
        If i = 0 Or i = UBound(x_vect) Then 'Sjekker om vi er på siste intervallet, i så fall er telleren vår ett hakk for langt frem, og vi går over rangen vår..
        dy_dx(j) = 1E+30
        'Utfører noen sjekker:
        ElseIf (y_vect(i + 1) - y_vect(i)) = 0 Or (y_vect(i) - y_vect(i - 1)) = 0 Then
            dy_dx(j) = 0
        ElseIf ((x_vect(i + 1) - x_vect(i)) / (y_vect(i + 1) - y_vect(i)) + (x_vect(i) - x_vect(i - 1)) / (y_vect(i) - y_vect(i - 1))) = 0 Then
    'Pos PLUS neg slope is 0. Prevent div by zero.
            dy_dx(j) = 0
        ElseIf (y_vect(i + 1) - y_vect(i)) * (y_vect(i) - y_vect(i - 1)) < 0 Then
    'Pos AND neg slope, assume slope = 0 to prevent overshoot
            dy_dx(j) = 0
        Else
         dy_dx(j) = 2 / ((((x_vect(i + 1) - x_vect(i)) / (y_vect(i + 1) - y_vect(i))) + ((x_vect(i) - x_vect(i - 1)) / (y_vect(i) - y_vect(i - 1)))))
    End If
Next j

'Sjekker om vi må inn med tilpassede deriverte i første og siste intervall:
If n = 1 Then
    dy_dx(0) = 3 / 2 * (y_vect(n) - y_vect(n - 1)) / (x_vect(n) - x_vect(n - 1)) - dy_dx(1) / 2
ElseIf n = UBound(x_vect) Then
   
       dy_dx(1) = 3 / 2 * (y_vect(n) - y_vect(n - 1)) / (x_vect(n) - x_vect(n - 1)) - dy_dx(0) / 2
End If

'-----------------CALCULATES THE 2nd DERIVATIVES------------------
dy_d2x(0) = -2 * (dy_dx(1) + 2 * dy_dx(0)) / (x_vect(n) - x_vect(n - 1)) + 6 * (y_vect(n) - y_vect(n - 1)) / (x_vect(n) - x_vect(n - 1)) ^ 2
dy_d2x(1) = 2 * (2 * dy_dx(1) + dy_dx(0)) / (x_vect(n) - x_vect(n - 1)) - 6 * (y_vect(n) - y_vect(n - 1)) / (x_vect(n) - x_vect(n - 1)) ^ 2

d = (dy_d2x(1) - dy_d2x(0)) / (6 * (x_vect(n) - x_vect(n - 1)))
c = (x_vect(n) * dy_d2x(0) - x_vect(n - 1) * dy_d2x(1)) / (2 * (x_vect(n) - x_vect(n - 1)))
b = ((y_vect(n) - y_vect(n - 1)) - c * (x_vect(n) ^ 2 - x_vect(n - 1) ^ 2) - d * (x_vect(n) ^ 3 - x_vect(n - 1) ^ 3)) / (x_vect(n) - x_vect(n - 1))
a = y_vect(n - 1) - b * x_vect(n - 1) - c * x_vect(n - 1) ^ 2 - d * x_vect(n - 1) ^ 3

'1st ORDER DERIVATIVES OF INTERMEDIATE POINTS

khjCubicSplineEng = d * x_value ^ 3 + c * x_value ^ 2 + b * x_value + a

Application.Calculation = xlCalculationAutomatic ' TUR ON AUTOMATIC CALCULATION AGAIN

End Function
