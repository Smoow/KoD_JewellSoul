#####################################
#                                   #
#   Author: Smoow                   #
#   Contact: smoowpdr@gmail.com     #
#                                   #
#####################################

# Bibliotecas e módulos necessários
import win32com 
import win32gui
import win32con
import win32api

from win32com import client
from time import sleep

print('\n\
       #####################################\n\
       #                                   #\n\
       #   Author:  Smoow                  #\n\
       #   Contact: smoow@protonmail.com   #\n\
       #                                   #\n\
       #####################################\n\n')

time = int(input('Digite o tempo (em minutos) que você deseja renovar a joia: '));


# ------ EXEMPLO ------
# [hwnd] Não importa o que você encontrar de gambiarra por aí. Isso daqui gerencia o Unique ID.
# ["Notepad"] Esse é o nome da aplicação/processo pai. Uma forma rápida de encontrar o nome que você precisa é checar no gerenciador de tarefas.
# ["Untitled - Notepad"] Esse é o nome da sub-aplicação/processo filho. Cheque o nome que você precisa através de gerenciador de tarefas (clicando no droparrow do processo).
# hwndMain = win32gui.FindWindow("Notepad", "Untitled - Notepad") isso retorna o Unique ID do main/pai.
# ------ EXEMPLO ------


# Essa é uma forma de avisarmos ao sistema que queremos a aplicação do KoD. Há outras maneiras, mas essa é uma boa forma.
# Ele reconhece como "processo do KoD" o último KoD que você deixou em foco. Ou seja, vai utilizar a renovação somente nesse último.
shell = win32com.client.Dispatch("WScript.Shell")
shell.AppActivate("King of Destiny")
hwndMain = win32gui.FindWindow(None, "King of Destiny")

# ["hwndMain"] Como informado, é o main/parent Unique ID - é usado para obter o Unique ID do processo filho.

# [win32con.GW_CHILD] Não testei por completo, mas isso aponta para o processo filho. 
#   Se há múltiplas intâncias, então provavelmente precisaremos iterar até achar a que queremos (ou olhar na documentação certinho);
#   Não afeta o modo que queremos no KoD.

# hwndChild = win32gui.GetWindow(hwndMain, win32con.GW_CHILD) retorna o Unique ID do processo filho indicado.
hwndChild = win32gui.GetWindow(hwndMain, win32con.GW_CHILD)

# Loop infinito para o macro
while(True):
    # [win32con.WM_CHAR] Isso é o que a SendMessage espere (suporta - e indica - tanto KeyDown como KeyUp)
    # [0x65] hexcode para 'E' (se precisar)
    # [101] int number para 'e'
    # [0] Algo sobre retorno?
    print(f'\n*** Enviando ao processo: &{hwndMain}   |   Renovando a cada {time} minutos! ***')
    temp = win32api.SendMessage(hwndMain, win32con.WM_CHAR, 101, 0)
    print(f"*** Item no slot '{chr(101)}' usado! ***\n")

    # print(temp)
    sleep(60*time)
