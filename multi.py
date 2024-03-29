from os import system as sys
from time import sleep
for y in range(2014, 2024):
    for s in range(1, 5):
        sys(f'start "" python main.py {y} {s}')
        # sys("exit") # idk if this is correct
    sleep(1800)
