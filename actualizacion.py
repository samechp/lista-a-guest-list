# -*- coding: utf-8 -*-
"""
Avisa si hay una version nueva publicada en GitHub.

Le pregunta a la API de GitHub cual es el ultimo Release y compara su etiqueta
con la version de este programa. No descarga ni instala nada: solo avisa y abre
la pagina de descarga si el usuario quiere.

Todo esto es opcional por definicion. El programa se usa muchas veces sin
internet, asi que cualquier fallo aqui se traga en silencio y la ventana sigue
funcionando igual.
"""

import json
import re
import urllib.error
import urllib.request

VERSION = '1.2'

REPOSITORIO = 'samechp/lista-a-guest-list'
API_ULTIMO = 'https://api.github.com/repos/%s/releases/latest' % REPOSITORIO
PAGINA_DESCARGA = 'https://github.com/%s/releases/latest' % REPOSITORIO

ESPERA = 6          # segundos; si la red esta lenta no vale la pena insistir


def numeros(etiqueta):
    """'v1.10.2' -> (1, 10, 2). Lo que no sea numero se ignora."""
    return tuple(int(n) for n in re.findall(r'\d+', etiqueta or ''))


def es_mas_nueva(remota, propia=None):
    """Compara por numeros, no por texto: v1.10 es mayor que v1.9."""
    a, b = numeros(remota), numeros(propia or VERSION)
    if not a:
        return False
    largo = max(len(a), len(b))
    return a + (0,) * (largo - len(a)) > b + (0,) * (largo - len(b))


def buscar(version_propia=None):
    """Devuelve {'version', 'url', 'notas'} si hay una mas nueva, o None.

    None significa tanto "esta al dia" como "no se pudo averiguar": para el
    usuario es lo mismo, y no hay nada util que mostrarle en ninguno de los dos
    casos.
    """
    version_propia = version_propia or VERSION
    try:
        peticion = urllib.request.Request(
            API_ULTIMO,
            headers={'Accept': 'application/vnd.github+json',
                     'User-Agent': 'lista-a-guest-list/%s' % version_propia})
        with urllib.request.urlopen(peticion, timeout=ESPERA) as respuesta:
            datos = json.load(respuesta)
    except (urllib.error.URLError, OSError, ValueError, TimeoutError):
        return None

    etiqueta = datos.get('tag_name') or ''
    if not es_mas_nueva(etiqueta, version_propia):
        return None
    return {'version': etiqueta.lstrip('vV'),
            'url': datos.get('html_url') or PAGINA_DESCARGA,
            'notas': (datos.get('body') or '').strip()}
