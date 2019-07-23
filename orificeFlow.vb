Function OrificeFlow(OrificeType As Integer, D_l As Double, d_o As Double, rho As Single, my As Single, H_in As Double, H_out As Double, Kappa As Double, FluidType As Integer, Optional ByVal epsType As Integer) As Variant
' ---------------------------------------------------------------
' Funksjon Calculates flow through an orifice with known pipe diameter D_l and orifice opening d_o - Ref ISO 5167. 
' Input : OrificeFlow(OrificeType, D_l, d_o, rho, my, H_in, H_out, kappa, FluidType)
' Output: Vector with several values - ref bottom of function 
' NB: H_in og H_out = Head at inlet and outlet pressure tapping. Dstr Pipe Pressure is calculated from PressureLoss

' Based On ISO: 5167-2:2003 *REMEMBER - P_in = rho_in * g * H_in <-AND-> P_out = rho_out * g * H_out
' H_In and H_Out in absolute head meter. Should be specified as absolute head meter!
' COPYRIGHT: KRISTIAN HOLM JENSEN - mob: +47 92428196, mailto: Kristian.Holm.Jensen@gmail.com
' FluidType: 1=liquid, 2=Gas, 3 = twophase
' kappa = Cp/Cv
' ---------------------------------------------------------------
Dim g_funct As Double, g_derivfunct As Double, Qv As Double, Qv_old As Double, konvergens As Double, dH As Single, d_Omega As Double, NRe As Double
Dim C_funct As Double, C_derivfunct As Double, MaxMassRate As Double
Dim Solvect As Variant, EquivKfact As Double, OrificeVel As Double, PipeVel As Double
Dim Beta As Single, PiNr As Single, A_fact As Double, dAfact_dQv As Double, M_2 As Double, L_1 As Single, L_2 As Single, EPS As Double, g As Single, eps_Perry As Double
Dim r As Double, r_c As Double, f_rc As Double, df_rc As Double, r_c_Conv As Double, Y_exp As Double
Dim i As Integer, iterCount As Integer
ReDim Solvect(1 To 14, 1 To 1) 'Solution Vector (Qv,antall iterasjoner,Abs error (konvergens))
Dim FlowInfo As String
Dim BetaInfo As String
Dim rInfo As String
Dim d_oInfo As String
Dim SolverInfo As String

Beta = d_o / D_l
PiNr = WorksheetFunction.pi()
g = 9.807
EPS = 1 'Her bør det legges inn kompressibilitetsfaktor...
dH = Abs(H_in - H_out) 'Definerer trykkfallet
r = H_out / H_in 'ABSOLUTE pressure drop over orifice - H_in and H_Out Must be specified as absolute head!
r_c = 0.5 ' Initial guess
r_c_Conv = 1

If FluidType = 1 Then 'For Liquids
    Y_exp = 1

ElseIf FluidType = 2 Then 'for Pure Ideal Gases
    Y_exp = Math.Sqr(r ^ (2 / Kappa) * (Kappa / (Kappa - 1)) * ((1 - r ^ ((Kappa - 1) / Kappa)) / (1 - r)) * ((1 - Beta * Beta * Beta * Beta) / (1 - Beta * Beta * Beta * Beta * r ^ (2 / Kappa)))) '
    EPS = Y_exp
    
    'HENTET FRA ISO5167 - 5.3.2.2
    If epsType = 0 Or epsType = 1 Then
       EPS = 1 - (0.351 + 0.256 * Beta * Beta * Beta * Beta + 0.93 * Beta * Beta * Beta * Beta * Beta * Beta * Beta * Beta) * (1 - r ^ (1 / Kappa)) ' Kun gyldig hvis r>0.75 (r=p2/p1) (iht ISO5167)
    ElseIf epsType = 2 Then
    'Section retrieved from Perry - Chap 10 eq 10-29 (Calculation of Y_exp in the case of subsonic gas flow --> r_c < r < 1 (eq 10-29 in perry)
       eps_Perry = 1 - (1 - r) / Kappa * (0.41 + 0.35 * Beta * Beta * Beta * Beta)
       EPS = eps_Perry
    End If
    
   i = 0
     Do While r_c_Conv > 0.00000001 And i < 20 'BEREGNING AV r_c --> Perry Chap 10 eq 10-24 (assumption: Perfect Gas, frictionless Nozzle)
        f_rc = r_c ^ ((1 - Kappa) / Kappa) + ((Kappa - 1) / 2) * Beta * Beta * Beta * Beta * r_c ^ (2 / Kappa) - (Kappa + 1) / 2
        df_rc = ((1 - Kappa) / Kappa) * r_c ^ ((1 - 2 * Kappa) / Kappa) + ((Kappa - 1) / Kappa) * Beta * Beta * Beta * Beta * r_c ^ ((2 - Kappa) / Kappa)
        r_c = r_c - f_rc / df_rc
        r_c_Conv = Abs(f_rc / df_rc)
        i = i + 1
     Loop
       
       'Section retrieved from Perry - Chap 10 eq 10-29 (Calculation of Y_exp in the case of subsonic gas flow --> r_c < r < 1
       Dim Y_exp_Perry As Double
       Y_exp_Perry = 1 - (1 - r) / Kappa * (0.41 + 0.35 * Beta * Beta * Beta * Beta)
       
       If r < 0.75 Then 'I henhold til ISO 5167 - Eps blir feil ved lavere trykkforhold (trykkrate P2/P1 < 0.75)
            rInfo = "Gas Calculation Does not comply with ISO 5167-2 chap 5.3.2.2 - Expansibility Calculation valid Only For: P2/P1 > 0.75"
       End If
       
       
ElseIf FluidType = 3 Then 'for twophase
    Exit Function 'Does not work
End If

Qv = 0.61 / ((1 - Beta * Beta * Beta * Beta) ^ 0.5) * (PiNr / 4) * (d_o / 1000) ^ 2 * (2 * g * dH * (1 - Beta ^ 1.9)) ^ 0.5 'Initialverdi: C=0.61
Qv = Qv * EPS
konvergens = 1
i = 0

If Not IsNumeric(OrificeType) Then
    FlowInfo = "OrificeType Not a Number: Must be between 1 and 3. 1 = Corner Tappings, 2 = D & D/2 Tappings, 3 = Flange Tappings "
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

If Beta < 0.1 Or Beta > 0.75 Then
    BetaInfo = "beta = D(line) / d(orifice) = " & Beta & "   VALUE OUT OF RANGE: Calculation valid for: 0.1 < beta < 0.75  "
End If
    
If d_o < 12.5 Then
    d_oInfo = "d_o < 12.5 mm. VALUE OUT OF RANGE - ISO 5167-2 Not Appliccable, ref chap 5.1.8.1 "
End If
Do While konvergens > 10 ^ -13
    NRe = 4 * rho * Qv / (PiNr * D_l * 0.001 * my)
    A_fact = (19000 * Beta / NRe) ^ 0.8
    dAfact_dQv = -0.8 * (19000 * Beta * PiNr * D_l * 0.001 * my / (4 * rho)) ^ 0.8 * Qv ^ -1.8
    M_2 = 2 * L_2 / (1 - Beta)
    
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
                    
    C_derivfunct = -0.7 * 0.000521 * (Beta * 10 ^ 6) ^ 0.7 * (PiNr * D_l * 0.001 * my / (4 * rho)) ^ 0.7 * Qv ^ -1.7 _
        - 0.3 * 0.0188 * Beta ^ 3.5 * 10 ^ 1.8 * ((PiNr * D_l * 0.001 * my / (4 * rho)) ^ 0.3) * Qv ^ -1.3 _
        + 0.0063 * Beta ^ 3.5 * 10 ^ 1.8 * ((1 / NRe) ^ 0.3 * dAfact_dQv - A_fact * 0.3 * (PiNr * D_l * 0.001 * my / (4 * rho)) ^ 0.3 * Qv ^ -1.3) _
        - 0.11 * (0.043 + 0.08 * Exp(-10 * L_1) - 0.123 * Exp(-7 * L_1)) * (Beta * Beta * Beta * Beta / (1 - Beta * Beta * Beta * Beta)) * dAfact_dQv
                 
    d_Omega = (((1 - Beta * Beta * Beta * Beta * (1 - C_funct * C_funct)) ^ 0.5 - C_funct * Beta * Beta) / ((1 - Beta * Beta * Beta * Beta * (1 - C_funct * C_funct)) ^ 0.5 + C_funct * Beta * Beta)) * dH
    'd_Omega = (((1 - Beta * Beta * Beta * Beta * (1 - C_funct ^ 2)) ^ 0.5 + C_funct * Beta ^ 2) / ((1 - Beta * Beta * Beta * Beta * (1 - C_funct ^ 2)) ^ 0.5 - C_funct * Beta * Beta)) * dH 'DENNE ER FEIL!!
                       
    'g_funct = Qv - C_funct / ((1 - Beta * Beta * Beta * Beta) ^ 0.5) * eps * (PiNr / 4) * (d_o * 0.001) ^ 2 * (2 * g * d_Omega) ^ 0.5 'Funksjon som vi utfører NewtonRhapson på
    'g_derivfunct = 1 - 1 / (1 - Beta * Beta * Beta * Beta) ^ 0.5 * eps * PiNr / 4 * (d_o * 0.001) ^ 2 * (2 * g * d_Omega) ^ 0.5 * C_derivfunct 'Den deriverte til g
    g_funct = Qv - C_funct / ((1 - Beta * Beta * Beta * Beta) ^ 0.5) * EPS * (PiNr / 4) * (d_o * 0.001) ^ 2 * (2 * g * dH) ^ 0.5 'Funksjon som vi utfører NewtonRhapson på
    g_derivfunct = 1 - 1 / (1 - Beta * Beta * Beta * Beta) ^ 0.5 * EPS * PiNr / 4 * (d_o * 0.001) ^ 2 * (2 * g * dH) ^ 0.5 * C_derivfunct 'Den deriverte til g
    
    Qv_old = Qv
    Qv = Qv - g_funct / g_derivfunct
    konvergens = Abs(Qv_old - Qv)
    i = i + 1
    iterCount = i
    If i > 200 Then
        MsgBox "Solution did not converge properly: Convergence:  " & konvergens
        SolverInfo = "Solution did not converge properly: Convergence:  " & konvergens
        Exit Do
    End If
Loop
    
'Determining Flow Type - Sonic or Subsonic
 Dim CritFlowInfo As String
    If r > r_c Then
        CritFlowInfo = "SubSonic Flow"
                
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
    
    i = 0
    
EquivKfact = (((1 - Beta * Beta * Beta * Beta * (1 - C_funct ^ 2)) ^ 0.5 / C_funct / Beta / Beta) - 1) ^ 2
'EquivKfact = 2 * dH / (rho * (Qv / (PiNr * (D_l * 0.001 / 2))) ^ 2) 'Ekvivalent K-faktor for strupeskive
OrificeVel = 4 * Qv / (PiNr * (d_o * 0.001) ^ 2)
PipeVel = 4 * Qv / (PiNr * (D_l * 0.001) ^ 2)

Dim AllFlowInfo As String
AllFlowInfo = FlowInfo & vbCrLf & CritFlowInfo & vbCrLf & SolverInfo & vbCrLf & BetaInfo & vbCrLf & d_oInfo & vbCrLf & rInfo
    
'MsgBox "ITER: =     " & i & vbCrLf & "beta =     " & beta & vbCrLf & "C(NRe,beta):     " & C_funct & vbCrLf & "H_omega=     " & d_omega & vbCrLf & "Konvergens:     " & konvergens
Solvect(1, 1) = Qv '(m3/s)
Solvect(2, 1) = iterCount
Solvect(3, 1) = konvergens
Solvect(4, 1) = EquivKfact
Solvect(5, 1) = C_funct
Solvect(6, 1) = NRe
Solvect(7, 1) = Beta
Solvect(8, 1) = OrificeVel 'm/s
Solvect(9, 1) = PipeVel 'm/s
Solvect(10, 1) = r_c
Solvect(11, 1) = EPS
Solvect(12, 1) = AllFlowInfo
Solvect(13, 1) = r
Solvect(14, 1) = d_Omega 'Pressure Loss [m]



OrificeFlow = Solvect

End Function
