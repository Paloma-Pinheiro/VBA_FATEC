Attribute VB_Name = "Exercicio_02_VBA"
Option Explicit
Sub SalarioReal()

'Receba o salário de um funcionário e mostre o novo salário com reajuste de 15%.

Dim salario As Double
Dim reajuste As Double

salario = InputBox("Insira o salário do funcionário:")
reajuste = (salario * 1.15)

MsgBox ("Nov salário: " & reajuste)
End Sub
