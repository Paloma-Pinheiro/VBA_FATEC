Attribute VB_Name = "Exercicio_09_VBA"
Option Explicit
Sub zsub()

'Receba 2 n�meros reais. Calcule e mostre a diferen�a desses valores.

Dim x As Double
Dim y As Double
Dim z As Double

x = InputBox("Insira o primeiro valor:")
y = InputBox("Insira o segundo valor:")

z = x - y
MsgBox ("Diferen�a dos valores: " & z)
End Sub
