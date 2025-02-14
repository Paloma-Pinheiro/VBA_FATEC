Attribute VB_Name = "Exercicio_12_VBA"
Option Explicit
Sub consumo()

'Receba a quantidade de alimento em quilos. Calcule e mostre quantos dias durará esse alimento sabendo que a pessoa consome 50g ao dia.

Dim k As Double
Dim d As Double

k = InputBox("insira o a quantidade de alimento: ")
k = k * 100
d = k / 50

MsgBox ("a comida irá durar por " & d & " dias.")

End Sub
