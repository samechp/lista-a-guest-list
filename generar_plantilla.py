# -*- coding: utf-8 -*-
"""Regenera plantilla_incrustada.py a partir de plantilla.glf.

Solo hace falta si la app Guest List cambia el formato: exporta un evento nuevo,
guardalo como plantilla.glf y corre este archivo.
"""
import os

AQUI = os.path.dirname(os.path.abspath(__file__))

def main():
    with open(os.path.join(AQUI, 'plantilla.glf'), encoding='utf-8', newline='') as f:
        base = f.read()
    cabecera, resto = base.split('\t\t<Guests>\n', 1)
    _, pie = resto.split('\t\t</Guests>\n', 1)
    src = ['# -*- coding: utf-8 -*-',
           '"""Cabecera y pie de un .glf exportado por la propia app Guest List.',
           '',
           'Generado desde plantilla.glf por generar_plantilla.py. No editar a mano.',
           '"""',
           '',
           'CABECERA = %r' % cabecera,
           '',
           'PIE = %r' % pie,
           '']
    destino = os.path.join(AQUI, 'plantilla_incrustada.py')
    with open(destino, 'w', encoding='utf-8') as f:
        f.write('\n'.join(src))
    print('Escrito %s (%d + %d caracteres).' % (destino, len(cabecera), len(pie)))

if __name__ == '__main__':
    main()
