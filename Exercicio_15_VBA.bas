Attribute VB_Name = "Exercicio_15_VBA"
Option Explicit
Sub pagamento()

'Receba a quantidade de horas trabalhadas, o valor por hora, o percentual de desconto e o n�mero de descendentes.
'Calcule o sal�rio que ser�o as horas trabalhadas x o valor por hora. Calcule o sal�rio l�quido (= Sal�rio Bruto - desconto).
'A cada dependente ser� acrescido R$ 100 no Sal�rio L�quido. Exiba o sal�rio a receber.

Dim t As Double
Dim v As Double
Dim f As Double
Dim p As Double
Dim d As Double

t = InputBox("Insira a quantidade de horas trabalhadas: ")
v = InputBox("Insira o valor por hora: ")
p = InputBox("Insira o percentual de desconto (em decimal): ")
d = InputBox("Insira a quantidade de dependentes: ")


f = salario(t, v, p, d)

End Sub

Function salario(t#, v#, p#, d As Double)

Dim l As Double
Dim s As Double
Dim b As Double

b = t * v
l = b - (b * p)
d = d * 100
s = l + b
MsgBox ("Sal�rio a receber: " & s)
End Function



