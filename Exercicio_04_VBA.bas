Attribute VB_Name = "Exercicio_04_VBA"
Option Explicit
Sub temperatura()

'Receba a temperatura em graus Celsius. Calcule e mostre a sua temperatura convertida em fahrenheit F = 1.8 * C + 32.

Dim celsius As Integer
Dim fahrenheit

celsius = InputBox("Insira a temperatura em Celsius")
fahrenheit = 1.8 * celsius + 32

MsgBox ("A temperatura " & celsius & "ºC convertida para Fahrenheit fica: " & fahrenheit & "ºF")


End Sub
