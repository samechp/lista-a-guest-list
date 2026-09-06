# -*- coding: utf-8 -*-
"""
Dibuja el icono del programa: una mesa redonda vista desde arriba con sus puestos.

    python hacer_icono.py

Deja `icono.ico` (para el .exe y la barra de titulo) e `icono.py` (el mismo dibujo
en base64, para que la ventana lo pueda poner sin depender de archivos sueltos).
Solo hace falta correrlo si se cambia el diseno; necesita Pillow.
"""

import io
import base64
import math
import os

from PIL import Image, ImageDraw

AQUI = os.path.dirname(os.path.abspath(__file__))
LADO = 1024

AZUL_ARRIBA = (78, 140, 246)
AZUL_ABAJO = (37, 89, 196)
BLANCO = (255, 255, 255)


def rectangulo_redondo(dibujo, caja, radio, relleno):
    dibujo.rounded_rectangle(caja, radius=radio, fill=relleno)


def dibujar(lado=LADO):
    """La mesa y sus puestos, sobre un cuadrado azul de esquinas redondeadas."""
    img = Image.new('RGBA', (lado, lado), (0, 0, 0, 0))

    # Fondo con un degradado vertical suave, recortado a esquinas redondeadas.
    degradado = Image.new('RGB', (1, lado))
    for y in range(lado):
        t = y / (lado - 1)
        degradado.putpixel((0, y), tuple(
            int(AZUL_ARRIBA[i] + (AZUL_ABAJO[i] - AZUL_ARRIBA[i]) * t) for i in range(3)))
    degradado = degradado.resize((lado, lado))

    mascara = Image.new('L', (lado, lado), 0)
    ImageDraw.Draw(mascara).rounded_rectangle(
        (0, 0, lado - 1, lado - 1), radius=int(lado * 0.22), fill=255)
    img.paste(degradado, (0, 0), mascara)

    d = ImageDraw.Draw(img)
    centro = lado / 2

    # Los puestos alrededor de la mesa.
    radio_puesto = lado * 0.083
    distancia = lado * 0.325
    for i in range(6):
        ang = math.radians(-90 + i * 60)
        x = centro + distancia * math.cos(ang)
        y = centro + distancia * math.sin(ang)
        d.ellipse((x - radio_puesto, y - radio_puesto,
                   x + radio_puesto, y + radio_puesto), fill=BLANCO)

    # La mesa.
    radio_mesa = lado * 0.225
    d.ellipse((centro - radio_mesa, centro - radio_mesa,
               centro + radio_mesa, centro + radio_mesa), fill=BLANCO)
    return img


def main():
    grande = dibujar()

    # El .ico lleva varios tamanos; Windows escoge el que necesite.
    medidas = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    ruta_ico = os.path.join(AQUI, 'icono.ico')
    grande.save(ruta_ico, format='ICO', sizes=medidas)

    # El mismo dibujo en base64, para ponerlo en la ventana sin archivos externos.
    buffer = io.BytesIO()
    grande.resize((256, 256), Image.LANCZOS).save(buffer, format='PNG')
    datos = base64.b64encode(buffer.getvalue()).decode('ascii')
    trozos = '\n'.join("    '%s'" % datos[i:i + 92] for i in range(0, len(datos), 92))

    with open(os.path.join(AQUI, 'icono.py'), 'w', encoding='utf-8') as f:
        f.write('# -*- coding: utf-8 -*-\n'
                '"""Icono de la ventana en base64. Generado por hacer_icono.py, no editar."""\n\n'
                'PNG = (\n%s\n)\n' % trozos)

    print('Escritos icono.ico (%d bytes) e icono.py' % os.path.getsize(ruta_ico))


if __name__ == '__main__':
    main()
