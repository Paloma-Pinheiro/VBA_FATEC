Attribute VB_Name = "Exercicio_13_VBA"
Option Explicit
Sub triangulo()

'Receba 2 �ngulos de um tri�ngulo. Calcule e mostre o valor do 3� �ngulo.

Dim p As Integer
Dim s As Integer
Dim t As Integer

p = InputBox("Primeiro �ngulo: ")
s = InputBox("Segundo �ngulo: ")

t = 180 - (p + s)
MsgBox ("Terceiro �ngulo: " & t)
End Sub
