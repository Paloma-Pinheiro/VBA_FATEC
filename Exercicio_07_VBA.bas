Attribute VB_Name = "Exercicio_07_VBA"
Option Explicit
Sub deposito()

'Receba o valor de um dep�sito em poupan�a. Calcule e mostre o valor ap�s 1 m�s de aplica��o sabendo que rende 1,3% a. m.

Dim deposito As Double

deposito = InputBox("Insira o valor de deposito:")
deposito = deposito + (deposito * 0.013)
MsgBox ("Valor ap�s 1 m�s de aplica��o: " & deposito)
End Sub
