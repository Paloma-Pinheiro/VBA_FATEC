Attribute VB_Name = "Exercicio_08_VBA"
Option Explicit
Sub soma()

'Receba os 2 números inteiros. Calcule e mostre a soma dos quadrados.

Dim x As Integer
Dim y As Integer
Dim z As Integer

x = InputBox("Insira o primeiro valor:")
y = InputBox("Insira o segundo valor:")

z = (x * x) + (y * y)

MsgBox ("Soma dos quadrados: " & z)
End Sub
