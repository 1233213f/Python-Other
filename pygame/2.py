import random

base = {'naipe':['PAUS', 'COPAS', 'ESPADA', 'OUROS'],
    'numero':['A', '2', '3', '4', '5', '6', '7', 'J', 'Q', 'K']} #Colocar o 8, 9, 10 se for usar o baralho inteiro
baralho = list()
jogador1 = list()
computador = list()
carta = {}

def embaralhar(baral):
    n = len(baral)
    while len(baral) <= 39: #Se for usar o baralho inteiro <= 51
        n = len(baral)
        carta['naipe'] = base['naipe'][random.randint(0,3)]
        carta['numero'] = base['numero'][random.randint(0,9)] #Se for usar o baralho inteiro, até 12
        baral.append(carta.copy())
        for i, v in enumerate(baral):
                if i == len(baral)-1:
                    break
                elif v == carta:
                    baral.pop()
                    break

def distribuir(baralho, jogador):
    for i in range(3):
        n = len(baralho)
        jogador.append(baralho[n-1])
        baralho.pop()

def coringa(baralho):
    coringa = baralho[len(baralho)-1]
    baralho.pop()
    return coringa

embaralhar(baralho)
distribuir(baralho, jogador1)
distribuir(baralho, computador)
coringa = coringa(baralho)
print(baralho)
print(jogador1)
print(computador)
print(coringa)