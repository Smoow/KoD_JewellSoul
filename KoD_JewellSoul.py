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
import math

from win32com import client
from time import sleep


def main():
    show_credits()
    mode = choose_mode()

    init_macro(mode)


def show_credits():
    print('\n\
        #####################################\n\
        #                                   #\n\
        #   Author:  Smoow                  #\n\
        #   Contact: smoow@protonmail.com   #\n\
        #                                   #\n\
        #####################################\n\n')


def choose_mode():
    print('Modos: 1) Joia       2) Soul       3) Joia + Soul       4) Fada Lan (800x600)')
    mode = int(input('Digite o modo de operação do programa: '))

    return mode


def attach_process():
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

    return hwndMain


def init_macro(mode):
    # Loop infinito para o macro
    if (mode == 1):
       print('\nModo escolhido: Renovação de Joia.')
       print('Uso:\n\t- Deixe a joia de renovação no slot R (quarto slot) da barra de atalho')
       time = int(
           input('\nDigite o tempo (em minutos) que você deseja renovar a joia: '))

       process = attach_process()
       while(True):
            # [win32con.WM_CHAR] Isso é o que a SendMessage espere (suporta - e indica - tanto KeyDown como KeyUp)
            # [0x65] hexcode para 'E' (se precisar)
            # [114] ascii code para 'R'
            # [0] Algo sobre retorno?
            # temp = win32api.SendMessage(hwndMain, win32con.WM_CHAR, 101, 0)
              print(
                f'\n*** Enviando ao processo: &{process}   |   Renovando a cada {time} minutos! ***')
              temp = win32api.SendMessage(process, win32con.WM_CHAR, 114, 0)
              print(f"*** Item no slot '{chr(114)}' usado! ***\n")

              # print(temp)
              sleep(60*time)

    if (mode == 2):
       print('\nModo escolhido: Renovação de Soul.')
       print('Uso:\n\t- Deixe a skill Soul primeiro slot barra de skills')

       process = attach_process()
       while(True):
            # [win32con.WM_CHAR] Isso é o que a SendMessage espere (suporta - e indica - tanto KeyDown como KeyUp)
            # [0x31] hexacode para '1'
            # [0] Algo sobre retorno?
            # temp = win32api.SendMessage(hwndMain, win32con.WM_CHAR, 101, 0)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.AppActivate("King of Destiny")
            print(
                    f'\n*** Enviando ao processo: &{process}   |   Renovando a cada 17 minutos! ***')
            win32api.SendMessage(process, win32con.WM_CHAR, 0x32, 0)

            # Para obter altura e largura da janela
            shell.AppActivate("King of Destiny")
            rect = win32gui.GetWindowRect(process)
            x1 = rect[0] 
            y1 = rect[1] 
            x2 = rect[2]
            y2 = rect[3]

            x_click = x1 + math.floor((x2-x1)/2) 
            y_click = y1+ math.floor((y2-y1)/5) 

            print(x1, y1, x2, y2)

            # Colocar o cursor na posição da janela
            win32api.SetCursorPos((x_click, y_click))

            # Click direito
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x1, y1, 0, 0)
            sleep(0.1)
            win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x1, y1, 0, 0)

            # Para previnir a barra ficar travada, após a tentativa de usar a soul
            # é enviado um comando para setar a última skill da barra.
            sleep(1)
            win32api.SendMessage(process, win32con.WM_CHAR, 0x30, 0)

            sleep(60*17)

    if (mode == 3):
       print('Em construção.')
       return

    if (mode == 4):
        print('\nModo escolhido: Macro de Fada para Lan.')

        process = attach_process()

        shell = win32com.client.Dispatch("WScript.Shell")
        shell.AppActivate("King of Destiny")
        print(
                f'\n*** Enviando ao processo: &{process}   |   Macro de Fada para Lan! ***')

        while (True):
            # Para obter altura e largura da janela
            rect = win32gui.GetWindowRect(process)
            x = rect[0]+730
            y = rect[1]+310

            # Colocar o cursor na posição da janela
            win32api.SetCursorPos((x, y))

            # Click Esquerdo no ultimo slot da primeira linha do inv
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

            # Click Esquerdo no slot da fada
            y -= 160
            win32api.SetCursorPos((x, y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            sleep(2)

            x = rect[0]+730
            y = rect[1]+310

            # Colocar o cursor na posição da janela
            win32api.SetCursorPos((x, y))

            # Click Esquerdo no ultimo slot da primeira linha do inv
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

            # Click Esquerdo no slot da fada
            y -= 160
            win32api.SetCursorPos((x, y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            sleep(58)

main()
