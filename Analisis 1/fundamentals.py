import os


with open('archivo.txt', 'w') as arch:
    for i in range(100):
        arch.write(str(i))
        arch.write('\n')

with open('archivo.txt', 'r') as arch:
    for linea in arch:
        print(linea)


