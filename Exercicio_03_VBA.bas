Attribute VB_Name = "Exercicio_03_VBA"
Option Explicit
Sub triangulo()

'Receba a base e a altura de um triângulo. Calcule e mostre a sua área

Dim base As Double
Dim altura As Double
Dim area As Double

base = InputBox("Insira a base dos triângulo")
altura = InputBox("Insira a altura do triângulo")

area = (base * altura) / 2

MsgBox ("Área do triângulo: " & area)
End Sub
