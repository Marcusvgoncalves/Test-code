# Bibliotecas
import os
import platform
from win32api import GetFileVersionInfo, LOWORD, HIWORD
from win32com.client import Dispatch

# Arquivos sendo buscados em SO
def get_localFile(File):
    folder = 'C:\\'
    print ("Os arquivos .exe localizados %s:" % folder)
    file_list = []
    for (paths, dirs, files) in os.walk(folder):
        for file in files:
            if file.endswith(File):
                file_list.append(os.path.join(paths, file))

    return file_list

def get_version_number(SeuApp):
   try:
        info = GetFileVersionInfo(SeuApp, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        return Landesk, HIWORD(ms), LOWORD(ms), HIWORD(ls), LOWORD(ls)
   except:
        return Landesk, "Versão desconhecida"

if __name__ == "__main__":
    localFile = get_localFile('SeuApp')
    lenlocalFiles = len(localFile)
    for x in range(lenlocalFiles):
        version = ".".join([str(i) for i in get_version_number(localFile[x])])
        print(version)

    localFile = get_localFile('SeuApp')
    lenlocalFiles = len(localFile)
    for x in range(lenlocalFiles):
        version = ".".join([str(i) for i in get_version_number(localFile[x])])
        print(version)

# Procurando plataforma SO
so = platform.system()
so_version = platform.version()
nome_maquina = platform.node()

if so == 'Windows':
    print("Sistema Operacional é Windows")
elif so == 'Linux':
    print("Sistema Operacional é Linux")
else:
    print("Sistema Operacional é MacOS")

print(nome_maquina)
#print('Release do SO: ', so_version)