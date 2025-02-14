Attribute VB_Name = "Exercicio_09_VBA"
Option Explicit
Sub zsub()

'Receba 2 números reais. Calcule e mostre a diferença desses valores.

Dim x As Double
Dim y As Double
Dim z As Double

x = InputBox("Insira o primeiro valor:")
y = InputBox("Insira o segundo valor:")

z = x - y
MsgBox ("Diferença dos valores: " & z)
End Sub
