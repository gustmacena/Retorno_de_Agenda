import PyInstaller.__main__
import os
import shutil

# Limpar diretório dist se existir
if os.path.exists('dist'):
    shutil.rmtree('dist')

# Configurações do PyInstaller
PyInstaller.__main__.run([
    'retornos_atacadao.py',
    '--name=RetornoAtacadao',
    '--onefile',
    '--windowed',
    '--add-data=imgs/*;imgs',
    '--hidden-import=PIL._tkinter_finder',
    '--hidden-import=selenium',
    '--hidden-import=openpyxl',
    '--hidden-import=webdriver_manager',
    '--clean',
    '--noconfirm'
])

# Copiar arquivos necessários para a pasta dist
dist_folder = os.path.join('dist')
if not os.path.exists(dist_folder):
    os.makedirs(dist_folder)

# Copiar pasta de imagens
imgs_source = 'imgs'
imgs_dest = os.path.join(dist_folder, 'imgs')
if os.path.exists(imgs_source):
    if os.path.exists(imgs_dest):
        shutil.rmtree(imgs_dest)
    shutil.copytree(imgs_source, imgs_dest)

print("Build concluído com sucesso!")
print(f"O executável está disponível em: {dist_folder}") 