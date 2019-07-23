Function OrificeDiam(ByVal OrificeType As Variant, ByVal D_l As Double, ByVal Qv As Double, ByVal rho As Single, ByVal my As Single, ByVal H_in As Double, ByVal H_out As Double, ByVal Kappa As Double, ByVal FluidType As Integer, Optional ByVal epsType As Integer) As Variant
' ---------------------------------------------------------------
' Function calculates required orifice inner diameter d_o with know fluid properties
' Input: OrificeFlow(OrificeType [1=Corner Tappings, 2=D/D2 tappings, 3=Flange Tappings], 
'                    D_l [m], Qv [m3/s], rho [kg/m3], my [kg/ms], H_in [m], H_out [m], 
'                    kappa[- Cp/Cv], fluidtype [1=liquids, 2=gas], epsType [1=ISO_5167, 2=Perry])
' Output: Vector SolVect
' Based on ISO: 5167-2:2003
'
' KRISTIAN HOLM JENSEN - mailto: Kristian.Holm.Jensen@gmail.com
' ---------------------------------------------------------------
Dim g_funct As Double, g_derivfunct As Double, d_o As Double, d_o_old As Double, konvergens As Double, dH As Single, d_Omega As Double, NRe As Double
Dim C_funct As Double, C_derivfunct As Double, C_deriv_ext As Double
Dim EquivKfact As Double, Solvect As Variant, OrificeVel As Double, PipeVel As Double
Dim Beta As Single, PiNr As Single, A_fact As Double, M_2 As Double, L_1 As Single, L_2 As Single, EPS As Double, g As Single, eps_Perry As Double
Dim r As Double, r_c As Double, f_rc As Double, df_rc As Double, Y_exp As Double
Dim i As Integer, MaxMassRate As Double 'Max massrate through orifice under sonic flow.
Dim iterCount As Integer

Dim FlowInfo As String
Dim BetaInfo As String
Dim d_oInfo As String
Dim SolverInfo As String
Dim rInfo As String

ReDim Solvect(1 To 14, 1 To 1)
PiNr = WorksheetFunction.pi()
g = 9.807
EPS = 1
dH = Abs(H_in - H_out) 'Definerer trykkfallet
r = H_out / H_in ' ABSOLUTE meter pressure drop over orifice
konvergens = 1
i = 0

If Not IsNumeric(OrificeType) Then
    MsgBox "OrificeType Not a Number: Must be between 1 and 3. 1 = Corner Tappings, 2 = D & D/2 Tappings, 3 = Flange Tappings "
ElseIf OrificeType = 1 Then 'For Corner Tappings
    L_1 = 0
    L_2 = 0
ElseIf OrificeType = 2 Then 'For D & D/2 Tappings
    L_1 = 1
    L_2 = 0.47
ElseIf OrificeType = 3 Then 'For Flange Tappings
    L_1 = 25.4 / D_l
    L_2 = L_1
Else
    MsgBox "Type flange tappings (measurementpoints): 1 = Corner Tappings, 2 = D & D/2 Tappings, 3 = Flange Tappings "
Exit Function
End If

If FluidType = 1 Then 'For Liquids
    EPS = 1
    
ElseIf FluidType = 2 Then 'for Gases
    Y_exp = Math.Sqr(r ^ (2 / Kappa) * (Kappa / (Kappa - 1)) * ((1 - r ^ ((Kappa - 1) / Kappa)) / (1 - r)) * ((1 - Beta * Beta * Beta * Beta) / (1 - Beta * Beta * Beta * Beta * r ^ (2 / Kappa)))) '
    'eps = Y_exp
    
    'HENTET FRA ISO5167 - 5.3.2.2
    If epsType = 0 Or epsType = 1 Then
       EPS = 1 - (0.351 + 0.256 * Beta * Beta * Beta * Beta + 0.93 * Beta * Beta * Beta * Beta * Beta * Beta * Beta * Beta) * (1 - r ^ (1 / Kappa)) ' Kun gyldig hvis r>0.75 (r=p2/p1) (iht ISO5167)
    ElseIf epsType = 2 Then
    'Section retrieved from Perry - Chap 10 eq 10-29 (Calculation of Y_exp in the case of subsonic gas flow --> r_c < r < 1 (eq 10-29 in perry)
       eps_Perry = 1 - (1 - r) / Kappa * (0.41 + 0.35 * Beta * Beta * Beta * Beta)
       EPS = eps_Perry
    End If
        
    r_c = 0.5
    For i = 0 To 10
        f_rc = r_c ^ ((1 - Kappa) / Kappa) + ((Kappa - 1) / 2) * Beta * Beta * Beta * Beta * r_c ^ (2 / Kappa) - (Kappa + 1) / 2
        df_rc = ((1 - Kappa) / Kappa) * r_c ^ ((1 - 2 * Kappa) / Kappa) + ((Kappa - 1) / Kappa) * Beta * Beta * Beta * Beta * r_c ^ ((2 - Kappa) / Kappa)
        r_c = r_c - f_rc / df_rc
    Next i

    If r < 0.75 Then 'I henhold til ISO 5167 - Eps blir feil ved lavere trykkforhold (trykkrate P2/P1 < 0.75)
            rInfo = "Gas Calculation Does not comply with ISO 5167-2 chap 5.3.2.2 - Expansibility Calculation valid Only For: P2/P1 > 0.75"
    End If

    i = 0
    
ElseIf FluidType = 3 Then 'for twophase

End If

d_o = D_l / 2 'Estimert startvalue, assumption: Sqrt(1-beta^4) = 0.5

Do While konvergens > 10 ^ -13
    Beta = d_o / D_l
    NRe = 4 * rho * Qv / (PiNr * D_l * 0.001 * my)
    A_fact = (19000 * Beta / NRe) ^ 0.8
    M_2 = 2 * L_2 / (1 - Beta)

'----CORRECTING EPS BETWEEN ITERATIONS--------------------------------------
If FluidType = 1 Then 'LIQUID
    EPS = 1
ElseIf FluidType = 2 Then  'GAS
        If epsType = 0 Or epsType = 1 Then 'Coorecting Eps
           EPS = 1 - (0.351 + 0.256 * Beta * Beta * Beta * Beta + 0.93 * Beta * Beta * Beta * Beta * Beta * Beta * Beta * Beta) * (1 - r ^ (1 / Kappa)) ' Kun gyldig hvis r>0.75 (r=p2/p1) (iht ISO5167)
        ElseIf epsType = 2 Then
        'Section retrieved from Perry - Chap 10 eq 10-29 (Calculation of Y_exp in the case of subsonic gas flow --> r_c < r < 1 (eq 10-29 in perry)
           eps_Perry = 1 - (1 - r) / Kappa * (0.41 + 0.35 * Beta * Beta * Beta * Beta)
           EPS = eps_Perry
        End If
End If


'--------------------------------------
    If D_l > 71.12 Then
    C_funct = 0.5961 + 0.0261 * Beta ^ 2 - 0.216 * Beta * Beta * Beta * Beta * Beta * Beta * Beta * Beta + 0.000521 * (Beta * 10 ^ 6 / NRe) ^ 0.7 _
              + (0.0188 + 0.0063 * A_fact) * Beta ^ 3.5 * (10 ^ 6 / NRe) ^ 0.3 _
              + (0.043 + 0.08 * Exp(-10 * L_1) - 0.123 * Exp(-7 * L_1)) * (1 - 0.11 * A_fact) * (Beta * Beta * Beta * Beta / (1 - Beta * Beta * Beta * Beta)) _
              - 0.031 * (M_2 - 0.8 * M_2 ^ 1.1) * Beta ^ 1.3 'if D_l >= 71.12 mm

    ElseIf D_l <= 71.12 Then
    C_funct = 0.5961 + 0.0261 * Beta ^ 2 - 0.216 * Beta * Beta * Beta * Beta * Beta * Beta * Beta * Beta + 0.000521 * (Beta * 10 ^ 6 / NRe) ^ 0.7 _
              + (0.0188 + 0.0063 * A_fact) * Beta ^ 3.5 * (10 ^ 6 / NRe) ^ 0.3 _
              + (0.043 + 0.08 * Exp(-10 * L_1) - 0.123 * Exp(-7 * L_1)) * (1 - 0.11 * A_fact) * (Beta * Beta * Beta * Beta / (1 - Beta * Beta * Beta * Beta)) _
              - 0.031 * (M_2 - 0.8 * M_2 ^ 1.1) * Beta ^ 1.3 _
              + 0.011 * (0.75 - Beta) * (2.8 - D_l / 25.4) 'If D_l < 71.12 mm'if D_l >= 71.12 mm
    End If
                    
    C_deriv_ext = -0.031 * ((2 * D_l * 0.001 * L_2 * 1 / (D_l * 0.001 - d_o * 0.001) ^ 2 - 0.8 * (2 * L_2 * D_l * 0.001) ^ 1.1 * 1.1 / ((D_l * 0.001) - (d_o * 0.001)) ^ 0.1 / (D_l * 0.001 - d_o * 0.001) ^ 2) * Beta ^ 1.3 + _
                (2 * D_l * 0.001 * L_2 / (D_l * 0.001 - d_o * 0.001) - 0.8 * (2 * L_2 * D_l * 0.001 / (D_l * 0.001 - d_o * 0.001)) ^ 1.1) * 1.3 / (D_l * 0.001) ^ 1.3 * (d_o * 0.001) ^ 0.3) _
                - 0.011 / (D_l * 0.001) * (2.8 - D_l / 25.4)
                    
    C_derivfunct = 2 * 0.0261 * d_o * 0.001 / (D_l * 0.001) ^ 2 - 8 * 0.216 * (d_o * 0.001) ^ 7 / (D_l * 0.001) ^ 8 + 0.000521 * 0.7 * (10 ^ 6 / NRe) ^ 0.7 * (d_o * 0.001) ^ -0.3 / (D_l * 0.001) ^ 0.7 + _
                (0.0063 * (19000 / (NRe * D_l * 0.001)) ^ 0.8 * (d_o * 0.001) ^ -0.2 * (d_o / D_l) ^ 3.5 * (10 ^ 6 / NRe) ^ 0.3) + (0.0188 + 0.0063 * (19000 / (NRe * D_l * 0.001)) ^ 0.8 * (d_o * 0.001) ^ 0.8) * (1 / (D_l * 0.001)) ^ 3.5 * (d_o * 0.001) ^ 2.5 * (10 ^ 6 / NRe) ^ 0.3 + _
                (0.043 + 0.08 * Exp(-10 * L_1) - 0.123 * Exp(-7 * L_1)) * ((0.11 * (19000 / (NRe * D_l * 0.001)) ^ 0.8 * 0.8 * (d_o * 0.001) ^ -0.2) * (d_o * 0.001) ^ 4 / ((D_l * 0.001) ^ 4 - (d_o * 0.001) ^ 4) + (1 - 0.11 * (19000 / (NRe * D_l * 0.001)) ^ 0.8 * (d_o * 0.001) ^ 0.8) * 4 * (d_o * 0.001) ^ 3 * (D_l * 0.001) ^ 4 / ((D_l * 0.001) ^ 4 - (d_o * 0.001) ^ 4) ^ 2) _
                 - C_deriv_ext
    
    d_Omega = dH '(((1 - beta*beta*beta*beta * (1 - C_funct ^ 2)) ^ 0.5 + C_funct * beta ^ 2) / ((1 - beta*beta*beta*beta * (1 - C_funct ^ 2)) ^ 0.5 - C_funct * beta ^ 2)) * dH
                       
    g_funct = Qv - C_funct / (1 - Beta * Beta * Beta * Beta) ^ 0.5 * EPS * (PiNr / 4) * (d_o * 0.001) ^ 2 * (2 * g * Abs(d_Omega)) ^ 0.5 'Funksjon som vi utfører NewtonRhapson på
    g_derivfunct = -EPS * PiNr / 4 * (2 * g * Abs(d_Omega)) ^ 0.5 * ((d_o * 0.001) ^ 2 / (1 - Beta * Beta * Beta * Beta) * C_derivfunct + C_funct / (1 - Beta * Beta * Beta * Beta) * 2 * d_o * 0.001 + 2 * (d_o * 0.001) ^ 5 / (D_l * 0.001) ^ 4 * C_funct * (1 - Beta * Beta * Beta * Beta) ^ -1.5)
    
    d_o_old = d_o
    d_o = d_o - 1000 * g_funct / g_derivfunct
    'd_o = d_o - g_funct / g_derivfunct
    konvergens = Abs(d_o_old - d_o)
    i = i + 1
    iterCount = i
    If i > 250 Then
        MsgBox "# iterations is > 200.   Iteration ended! Convergense: " & konvergens
        Exit Do
    End If
Loop
    
Beta = d_o / D_l
If Beta < 0.1 Or Beta > 0.75 Then
    
    BetaInfo = "beta = D(line) / d(orifice) = " & Beta & "   VALUE OUT OF RANGE: Calculation valid for: 0.1 < beta < 0.75  - ref ISO-5167 2 chap 5.1.8.1 "
ElseIf Beta > 1 Then
    FlowInfo = "OBS beta > 1, impossible solution!! beta = " & Beta
End If

If d_o < 12.5 Then
    d_oInfo = "d_o < 12.5 mm. VALUE OUT OF EQUATION VALIDITY RANGE, ref ISO-5167 2 chap 5.1.8.1 "
End If

'Determining Flow Type - Sonic or Subsonic
Dim CritFlowInfo As String
    If r > r_c Then
        CritFlowInfo = "SubSonic Gas Flow"
                
    ElseIf r < r_c Then
        CritFlowInfo = "Critical/Sonic Gas Flow"
        If Beta < 0.2 Then
            Dim A_o As Double, P_inlet As Double
            P_inlet = rho * g * H_in / 100000#
            A_o = PiNr * (d_o * 0.001) ^ 2 / 4
            If Not Kappa = 1 Then
                MaxMassRate = C_funct * A_o * Sqr(Kappa * rho * P_inlet * 10 ^ 5 * (2 / (Kappa + 1)) ^ ((Kappa + 1) / (Kappa - 1)))
            End If
        End If
            
    End If

EquivKfact = (((1 - Beta * Beta * Beta * Beta * (1 - C_funct ^ 2)) ^ 0.5 / C_funct * Beta ^ 2) - 1) ^ 2
'EquivKfact = 2 * dH / (rho * (Qv / (PiNr * (D_l * 0.001 / 2))) ^ 2) 'Ekvivalent K-factor for strupeskive
OrificeVel = 4 * Qv / (PiNr * (d_o * 0.001) ^ 2)
PipeVel = 4 * Qv / (PiNr * (D_l * 0.001) ^ 2)
d_Omega = (((1 - Beta * Beta * Beta * Beta * (1 - C_funct * C_funct)) ^ 0.5 - C_funct * Beta * Beta) / ((1 - Beta * Beta * Beta * Beta * (1 - C_funct * C_funct)) ^ 0.5 + C_funct * Beta * Beta)) * dH 'PRESSURE LOSS [m]

Dim AllFlowInfo As String
AllFlowInfo = FlowInfo & vbCrLf & CritFlowInfo & vbCrLf & SolverInfo & vbCrLf & BetaInfo & vbCrLf & d_oInfo & vbCrLf & rInfo

'MsgBox "ITER: =     " & i & vbCrLf & "beta =     " & beta & vbCrLf & "C(NRe,beta):     " & C_funct & vbCrLf & "H_omega=     " & d_omega & vbCrLf & "Konvergens:     " & konvergens & vbCrLf & "Equivalent K-fact:    " & EquivKfact
Solvect(1, 1) = d_o 'Calculated orifice diameter 
Solvect(2, 1) = iterCount
Solvect(3, 1) = konvergens
Solvect(4, 1) = EquivKfact
Solvect(5, 1) = C_funct
Solvect(6, 1) = NRe
Solvect(7, 1) = Beta
Solvect(8, 1) = OrificeVel
Solvect(9, 1) = PipeVel
Solvect(10, 1) = r_c
Solvect(11, 1) = EPS
Solvect(12, 1) = AllFlowInfo 'Text Info about the calculation - E.g. if the result and input is according to ISO5167 etc.
Solvect(13, 1) = r
Solvect(14, 1) = d_Omega 'm

OrificeDiam = Solvect

End Function
