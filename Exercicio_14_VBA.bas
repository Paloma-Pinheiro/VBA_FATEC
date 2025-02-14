Attribute VB_Name = "Exercicio_14_VBA"
Option Explicit
Sub hipotenusa()

'Receba os valores de 2 catetos de um triângulo retângulo. Calcule e mostre a hipotenusa.

Dim p As Integer
Dim s As Integer
Dim h As Integer

p = InputBox("Primeiro cateto: ")
s = InputBox("Segundo cateto: ")
h = (p * p) + (s * s)
h = Sqr(h)
MsgBox ("Hipotenusa: " & h)

End Sub
