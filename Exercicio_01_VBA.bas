Attribute VB_Name = "Exercicio_01_VBA"
Option Explicit
Sub quadrado()

'Coletar o valor do lado de um quadrado, calcular sua �rea e apresentar o resultado.

Dim lado As Double
Dim area As Double

lado = InputBox("Insira o tamanho do lado do quadrado:")

area = lado * lado

MsgBox ("�rea do quadrado: " & area)

End Sub
