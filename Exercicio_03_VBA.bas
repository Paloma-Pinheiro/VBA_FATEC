Attribute VB_Name = "Exercicio_03_VBA"
Option Explicit
Sub triangulo()

'Receba a base e a altura de um tri�ngulo. Calcule e mostre a sua �rea

Dim base As Double
Dim altura As Double
Dim area As Double

base = InputBox("Insira a base dos tri�ngulo")
altura = InputBox("Insira a altura do tri�ngulo")

area = (base * altura) / 2

MsgBox ("�rea do tri�ngulo: " & area)
End Sub
