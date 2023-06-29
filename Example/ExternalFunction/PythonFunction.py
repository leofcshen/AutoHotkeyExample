from sys import argv

import pyperclip as pc
a1 = "Hey, nice to see you"
#pc.copy(a1)

def Func():
    total = int(argv[1]) + int(argv[2]) + int(argv[3])
    pc.copy(total)   
    return total
    
if __name__ == "__main__":
    print(Func())