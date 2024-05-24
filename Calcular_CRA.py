def calcular_cra(notas, creditos):
    soma_creditos = sum(creditos)
    soma_ponderada = sum(nota * credito for nota, credito in zip(notas, creditos))
    cra = soma_ponderada / soma_creditos
    return cra

# Exemplo de notas e créditos das disciplinas
notas = [6.5, 8.25, 9.5, 6, 8, 8.75, 7]
creditos = [4, 4, 4, 4, 4, 4, 4]

# Chamar a função para calcular o CRA
cra = calcular_cra(notas, creditos)
print("Seu CRA é:", cra)