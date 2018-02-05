Option Explicit
Option Base 0

'Devoir 3
'Programmation de modèles quantitatifs en finance
'Louis Gendreau, 11205384
'HEC Montréal, Automne 2017

Function Trinomial_Am(S, X, T, rf, sigma, n)

    Dim dt As Double, u As Double, d As Double, m As Double, p_u As Double, p_d As Double, p_m As Double

    'Pas de temps
    dt = T / n
    
    'Rendement si up, down et middle
    u = Exp(sigma * (3 * dt) ^ (1 / 2))
    d = Exp(-sigma * (3 * dt) ^ (1 / 2))
    m = 1
    
    'Probabilité de chaque mouvement (up, down et middle)
    p_u = ((dt / (12 * sigma ^ 2)) ^ (1 / 2)) * (rf - (sigma ^ 2 / 2)) + 1 / 6
    p_d = -(dt / (12 * sigma ^ 2)) ^ (1 / 2) * (rf - (sigma ^ 2 / 2)) + 1 / 6
    p_m = 2 / 3
    
    'Arbre du sous-jacent -------------------------------------
    Dim SJ() As Double
    ReDim SJ(0 To n * 2, 0 To n)
    SJ(0, 0) = S
    
    'On construit une matrice contenant le prix de l'action à chaque node
    Dim cc As Integer 'Column
    Dim rr As Integer 'Row
    For cc = 1 To n
        For rr = 0 To cc * 2
            'Straight
            If rr = cc * 2 - 1 Then
                SJ(rr, cc) = SJ(rr - 1, cc - 1)
            'Down
            ElseIf rr = cc * 2 Then
                SJ(rr, cc) = SJ(rr - 2, cc - 1) * d
            'Up
            Else
                SJ(rr, cc) = SJ(rr, cc - 1) * u
            End If
        Next rr
    Next cc
    
    'Arbre du put ----------------------------------------------
    'On construit une matrice avec le prix de l'option à chaque
    'node car on peut exercer une option américaine à n'importe quel moment
    Dim p_euro As Double
    Dim P() As Double
    ReDim P(0 To n * 2, 0 To n)
    
    'Prix a t=T
    For rr = 0 To n * 2
        P(rr, n) = Application.Max(X - SJ(rr, n), 0)
    Next rr
    
    'Induction à rebours
    For cc = n - 1 To 0 Step -1
        For rr = 0 To cc * 2
            p_euro = Exp(-rf * dt) * (P(rr, cc + 1) * p_u + P(rr + 1, cc + 1) * p_m + P(rr + 2, cc + 1) * p_d)
            P(rr, cc) = Application.Max(p_euro, X - SJ(rr, cc))
        Next rr
    Next cc
    
Trinomial_Am = P(0, 0)

End Function

Function Trinomial_Euro(S, X, T, rf, sigma, n)

    Dim dt As Double, u As Double, d As Double, m As Double, p_u As Double, p_d As Double, p_m As Double

    'Pas de temps
    dt = T / n
    
    'Rendement si up, down et middle
    u = Exp(sigma * (3 * dt) ^ (1 / 2))
    d = Exp(-sigma * (3 * dt) ^ (1 / 2))
    m = 1
    
    'Probabilité de chaque mouvement (up, down et middle)
    p_u = ((dt / (12 * sigma ^ 2)) ^ (1 / 2)) * (rf - (sigma ^ 2 / 2)) + 1 / 6
    p_d = -(dt / (12 * sigma ^ 2)) ^ (1 / 2) * (rf - (sigma ^ 2 / 2)) + 1 / 6
    p_m = 2 / 3
    
    'Arbre du sous-jacent -------------------------------------
    Dim SJ() As Double
    ReDim SJ(0 To n * 2, 0 To n)
    SJ(0, 0) = S
    
    Dim cc As Integer 'Column
    Dim rr As Integer 'Row
    For cc = 1 To n
        For rr = 0 To cc * 2
            'Straight
            If rr = cc * 2 - 1 Then
                SJ(rr, cc) = SJ(rr - 1, cc - 1)
            'Down
            ElseIf rr = cc * 2 Then
                SJ(rr, cc) = SJ(rr - 2, cc - 1) * d
            'Up
            Else
                SJ(rr, cc) = SJ(rr, cc - 1) * u
            End If
        Next rr
    Next cc

    'Arbre du put ----------------------------------------------
    Dim P() As Double
    ReDim P(0 To n * 2, 0 To n)
    
    'Prix a t=T
    For rr = 0 To n * 2
        P(rr, n) = Application.Max(X - SJ(rr, n), 0)
    Next rr
    
    'Induction à rebours
    For cc = n - 1 To 0 Step -1
        For rr = 0 To cc * 2
            P(rr, cc) = Exp(-rf * dt) * (P(rr, cc + 1) * p_u + P(rr + 1, cc + 1) * p_m + P(rr + 2, cc + 1) * p_d)
        Next rr
    Next cc
    
Trinomial_Euro = P(0, 0)

End Function

Function Mon_AmericanPut_Trinomial(S, X, T, rf, sigma, n, VC)

    'B représente une option similaire à l'option A
    
    Dim B_Black_Scholes As Double
    Dim B_Analytique As Double
    Dim A_Analytique As Double
    Dim prix_variate_control

    If VC = 0 Then
        prix_variate_control = Trinomial_Am(S, X, T, rf, sigma, n)
    Else
        
        'On a besoin de cest données pour l'approche de variable de controle
        B_Black_Scholes = BSPut(S, X, T, rf, sigma)
        B_Analytique = Trinomial_Euro(S, X, T, rf, sigma, n)
        A_Analytique = Trinomial_Am(S, X, T, rf, sigma, n)
        
        'Algorithme de l'approche de variable de controle
        prix_variate_control = B_Black_Scholes + (A_Analytique - B_Analytique)
    End If
        
    Mon_AmericanPut_Trinomial = prix_variate_control

End Function








