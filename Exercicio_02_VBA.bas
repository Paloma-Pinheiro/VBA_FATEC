Attribute VB_Name = "Exercicio_02_VBA"
Option Explicit
Sub SalarioReal()

'Receba o sal�rio de um funcion�rio e mostre o novo sal�rio com reajuste de 15%.

Dim salario As Double
Dim reajuste As Double

salario = InputBox("Insira o sal�rio do funcion�rio:")
reajuste = (salario * 1.15)

MsgBox ("Nov sal�rio: " & reajuste)
End Sub
