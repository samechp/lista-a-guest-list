# -*- coding: utf-8 -*-
"""
Empaqueta app.py como un unico .exe que no necesita Python instalado.

    python construir_exe.py

Deja el programa en dist/Lista a Guest List.exe. Requiere pyinstaller
(pip install pyinstaller).
"""

import os
import shutil
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
NOMBRE = 'Lista a Guest List'


def main():
    # La plantilla va incrustada en plantilla_incrustada.py, asi que el .exe no
    # necesita ningun archivo suelto al lado.
    subprocess.check_call([sys.executable, os.path.join(AQUI, 'generar_plantilla.py')])

    # openpyxl importa estas si estan disponibles, pero solo las necesita para cosas
    # que no usamos (imagenes, parseo acelerado). Sacarlas baja el .exe a la mitad.
    # Verificado: la conversion da el mismo resultado sin ellas.
    SOBRA = ['numpy', 'PIL', 'lxml', 'defusedxml', 'pandas', 'matplotlib', 'scipy',
             'setuptools', 'pydoc_data', 'pytest', 'IPython', 'unittest']

    orden = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--windowed',
        '--noconfirm',
        '--clean',
        '--name', NOMBRE,
        '--icon', os.path.join(AQUI, 'icono.ico'),
        '--hidden-import', 'openpyxl',
        '--distpath', os.path.join(AQUI, 'dist'),
        '--workpath', os.path.join(AQUI, 'build'),
        '--specpath', os.path.join(AQUI, 'build'),
    ]
    for m in SOBRA:
        orden += ['--exclude-module', m]
    orden.append(os.path.join(AQUI, 'app.py'))
    print(' '.join(orden))
    subprocess.check_call(orden, cwd=AQUI)

    destino = os.path.join(AQUI, 'dist', NOMBRE + '.exe')
    print('')
    print('Listo: %s  (%.1f MB)' % (destino, os.path.getsize(destino) / 1e6))

    # build/ solo son restos del empaquetado
    shutil.rmtree(os.path.join(AQUI, 'build'), ignore_errors=True)


if __name__ == '__main__':
    main()
