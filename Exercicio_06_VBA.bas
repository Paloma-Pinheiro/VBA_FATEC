Attribute VB_Name = "Exercicio_06_VBA"
Option Explicit
Sub volume()

'Receba os valores do comprimento, largura e altura de um paralelep�pedo. Calcule e mostre seu volume.

Dim comprimento As Integer
Dim largura As Integer
Dim altura As Integer
Dim volume As Integer

comprimento = InputBox("Insira o comprimento:")
largura = InputBox("Insira a largura:")
altura = InputBox("Insira a altura")

volume = comprimento * largura * altura

MsgBox ("Volume do paralelep�pedo: " & volume)

End Sub
