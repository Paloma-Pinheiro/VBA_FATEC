Attribute VB_Name = "Exercicio_10_VBA"
Option Explicit
Sub raio()

'Receba o raio de uma circunfer�ncia. Calcule e mostre o comprimento da circunfer�ncia.

Dim r As Double
Dim c As Double

r = InputBox("Insira um valor:")
c = 6.28 * r
MsgBox ("Comprimento da circufer�ncia: " & c)

End Sub
