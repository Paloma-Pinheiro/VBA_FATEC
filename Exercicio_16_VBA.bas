Attribute VB_Name = "Exercicio_16_VBA"
Option Explicit
Sub automovel()

'Calcule a quantidade de litros gastos em uma viagem, sabendo que o automóvel faz 12 km/l. Receber o tempo de percurso e a velocidade média.

Dim t As Double
Dim v As Double
Dim f As Double


t = InputBox("Insira tempo de percurso: ")
v = InputBox("Insira a velocidade média")

f = litros(t, v)
End Sub

Function litros(t#, v As Double)
Dim d As Double
Dim c As Double

d = t * v
c = d / 12
MsgBox ("Quantidade de litros gastos em uma viagem: " & c)

End Function

