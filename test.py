from multiprocessing.connection import wait
import os, sys
print("\n")
print(os.getcwd())
path = os.getcwd()+"\Attestati Pxx Edxx"

if not(os.path.exists(path)):
    print("Creando")
    os.mkdir(path)

tupla = [("prova", 102),("hs"),("kdfj",120)]
print(tupla)