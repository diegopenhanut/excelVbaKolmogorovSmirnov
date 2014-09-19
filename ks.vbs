Option Explicit
Option Base 1
Sub normalidade()

 If Selection.Cells.Count = 1 Then
 MsgBox "você precisa selecionar mais que uma célula"
 End
 Else
  Dim n As Integer
  n = Selection.Cells.Count
  Dim contador As Integer
  contador = n
  Dim amostra() As Double
  ReDim amostra(n)
  
  'obtém média e desvio padrão da amostra
Dim média As Double
Dim DP As Double
média = Application.WorksheetFunction.Average(Selection.Cells)
DP = Application.WorksheetFunction.StDev(Selection.Cells)

  'seleciona a amostra,organiza do menor para o maior, transforma em normal reduzida, armazena em amostra().
  Dim x As Variant
  contador = n
  For Each x In Selection.Cells
  amostra(contador) = Application.WorksheetFunction.Small(Selection.Cells, contador)
  amostra(contador) = (amostra(contador) - média) / DP
    contador = contador - 1
    Next

'obtém a distribuição cumulativa empírica
Dim contador_2 As Integer
contador_2 = n + 1
Dim Dist_emp() As Double
ReDim Dist_emp(n + 1)
Dist_emp(1) = 0
Do While contador_2 > 0
Dist_emp(contador_2) = ((contador_2 - 1) / n)
contador_2 = contador_2 - 1
Loop

'obtém a distribuição cumulativa teórica
Dim Dis_teo() As Double
ReDim Dis_teo(n)
Dim contador_3 As Integer
contador_3 = n
Do While contador_3 > 0
Dis_teo(contador_3) = Application.WorksheetFunction.NormSDist(amostra(contador_3))
contador_3 = contador_3 - 1
Loop

'obtém a maior diferença absoluta
Dim Maior_dif() As Double
ReDim Maior_dif(2 * n)
contador_3 = n
Do While contador_3 > 0
Maior_dif(contador_3) = Dist_emp(contador_3) - Dis_teo(contador_3)
contador_3 = contador_3 - 1
Loop
contador_3 = 2 * n
contador = n
Do While contador_3 > n
Maior_dif(contador_3) = Dist_emp(contador + 1) - Dis_teo(contador)
contador_3 = contador_3 - 1
contador = contador - 1
Loop

'torna todos os valores positivos
contador = 2 * n
Do While contador > 0
If Maior_dif(contador) < 0 Then
Maior_dif(contador) = Maior_dif(contador) * (-1)
End If
contador = contador - 1
Loop

'seleciona a maior diferença
Dim MDA As Double
Dim p As Double
Dim Q As Double
Dim Z As Double
MDA = Application.WorksheetFunction.Large(Maior_dif(), 1)
Z = (Sqr(n) * MDA)
Select Case Sqr(n) * MDA
    Case Is < 0.27
        p = 1
    Case 0.27, Is < 1
        Q = Exp(-1.233701 * Z ^ -2)
        p = 1 - (2.506628 / Z) * (Q + (Q ^ 9) + (Q ^ 25))
    Case 1, 3.1
        Q = Exp(-2 * Z ^ 2)
        p = 2 * (Q - Q ^ 4 + Q ^ 9 - Q ^ 16)
    Case Is >= 3.1
        p = 0
    End Select
    If p > 0.05 Then
    MsgBox ("-> Média = " & média & Chr(13) _
    & "-> Desvio padrão = " & DP & Chr(13) _
    & "-> Maior diferença = " & (MDA) & Chr(13) _
    & "-> Escore z = " & Sqr(n) * MDA & Chr(13) _
    & "-> Valor de p = " & p & Chr(13) _
    & "A distribuição é normal com alfa < 0,05")
    Else
    MsgBox ("-> Média = " & média & Chr(13) _
    & "-> Desvio padrão = " & DP & Chr(13) _
    & "-> Maior diferença = " & (MDA) & Chr(13) _
    & "-> Escore z = " & Sqr(n) * MDA & Chr(13) _
    & "-> Valor de p = " & p & Chr(13) _
    & "A distribuição Nâo é normal com alfa < 0,05")
End If
End If
End Sub

