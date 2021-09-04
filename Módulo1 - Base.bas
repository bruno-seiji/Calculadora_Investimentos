Attribute VB_Name = "Módulo1"
Sub lucro()
Dim JurosM As Double
Dim Valor As Double
Dim Valorm As Double
Dim Total As Double
Dim Tempo As Integer
Dim JurosA As Double
Dim i As Integer

'-------------------------------------------
'Modulo de Limpeza de Planilha
Call limpar
'-------------------------------------------

Range("F1").Value = "Meses"
Range("G1").Value = "Valores Acumulados"
Range("A9").Select
JurosA = Cells(2, 2).Value
Valor = Cells(3, 2).Value
Tempo = Cells(4, 2).Value + 1
JurosM = ((1 + (JurosA / 100)) ^ (1 / 12)) - 1
Cells(2, 7).Value = Cells(5, 2).Value

For i = 3 To Tempo + 1
Cells(i, 7).Value = (Cells(i - 1, 7).Value * JurosM) + Valor + Cells(i - 1, 7).Value
Cells(i - 1, 6).Value = "Mês " & i - 3

Next
Cells(i - 1, 6).Value = "Mês " & i - 3

'-------------------------------------------
'Modulo de Criação de Tabela
Call CriaTabela(i)
'-------------------------------------------

Total = Cells(i - 1, 7)
Cells(6, 2).Value = Total
Cells(7, 2).Value = (Cells(3, 2).Value * Cells(4, 2).Value) + Cells(5, 2).Value
Cells(8, 2).Value = Cells(6, 2).Value - Cells(7, 2).Value
Range("A9").Select

End Sub
