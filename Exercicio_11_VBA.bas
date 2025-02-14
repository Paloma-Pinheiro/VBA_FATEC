Attribute VB_Name = "Exercicio_11_VBA"
Option Explicit
Sub idade()

'Receba o ano de nascimento e o ano atual. Calcule e mostre a sua idade e quantos anos terá daqui a 17 anos.

Dim nascimento As Integer
Dim atual As Integer
Dim idade As Integer
Dim dezessete As Integer

nascimento = InputBox("Insira o ano de nascimento:")
atual = InputBox("Insira o ano atual:")
idade = atual - nascimento
dezessete = idade + 17
MsgBox ("A idade atual: " & idade & " Daqui 17 anos a pessoa terá: " & dezessete)

End Sub
