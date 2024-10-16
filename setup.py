import sys
from cx_Freeze import setup, Executable

# Arquivos adicionais (inclua todos os arquivos necessários)
arquivos = ['excel.ico', 'HorariosEscritorio.xlsx','HorariosEstagiario3h30.xlsx','HorariosEstagiario6hr.xlsx','HorariosPatio.xlsx']

# Opções para incluir pacotes necessários manualmente
build_exe_options = {
    'packages': ['openpyxl', 'PySimpleGUI', 'locale', 'datetime'],  # Bibliotecas necessárias
    'include_files': arquivos,  # Arquivos adicionais
    'include_msvcr': True  # Incluir as bibliotecas do Visual C++ se necessário
}

# Configuração do executável
configuracao = Executable(
    script='app.py',  # O arquivo principal do seu projeto
    icon='excel.ico',  # Ícone do executável
    base='Win32GUI' if sys.platform == 'win32' else None  # Para interface gráfica sem console
)

# Configuração geral
setup(
    name='Horario dos Funcionarios',
    version='1.0',
    description='Programa que automatiza Excel',
    author='gabriel',
    options={'build_exe': build_exe_options},
    executables=[configuracao]
)