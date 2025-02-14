Attribute VB_Name = "Exercicio_13_VBA"
Option Explicit
Sub triangulo()

'Receba 2 ângulos de um triângulo. Calcule e mostre o valor do 3º ângulo.

Dim p As Integer
Dim s As Integer
Dim t As Integer

p = InputBox("Primeiro ângulo: ")
s = InputBox("Segundo ângulo: ")

t = 180 - (p + s)
MsgBox ("Terceiro ângulo: " & t)
End Sub
