Attribute VB_Name = "Exercicio_07_VBA"
Option Explicit
Sub deposito()

'Receba o valor de um depósito em poupança. Calcule e mostre o valor após 1 mês de aplicação sabendo que rende 1,3% a. m.

Dim deposito As Double

deposito = InputBox("Insira o valor de deposito:")
deposito = deposito + (deposito * 0.013)
MsgBox ("Valor após 1 mês de aplicação: " & deposito)
End Sub
