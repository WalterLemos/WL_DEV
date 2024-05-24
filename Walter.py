
import pyautogui

#Abertura
print("Bem-vindo à Calculadpra")
#Entrada de dados
Num1 = float(input("Insira o Primeiro Numero: "))
Num2 = float(input("Insira o Segundo Numero : ")) 
Operacao = (input("Digite a Operação: (+,-,*,/)")) 
if Operacao == "+": 
  resultado =Num1+Num2
  print("O Resultado é:  ", resultado)
if Operacao == "-": 
  resultado =Num1-Num2
  print("O Resultado é:  ", resultado)
if Operacao == "*": 
  resultado =Num1*Num2
  print("O Resultado é:  ", resultado) 
if Operacao == "/": 
   resultado =Num1/Num2
   print("O Resultado é:  ", resultado)
 