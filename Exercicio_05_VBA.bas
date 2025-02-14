Attribute VB_Name = "Exercicio_05_VBA"
Option Explicit
Sub trocar()

'Receba os valores em x e y. Efetua a troca de seus valores e mostre seus conteúdos.

Dim x As Integer
Dim y As Integer
Dim z As Integer

x = InputBox("Insira o valor de x:")
y = InputBox("Insira o valor de y:")

z = x
x = y
y = z

MsgBox ("Valor de X: " & x & "Valor de Y: " & y)

End Sub
