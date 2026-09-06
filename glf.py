# -*- coding: utf-8 -*-
"""
Lectura y escritura de archivos .glf (Guest List, iPad/iPhone).

Un .glf es XML plano UTF-8 sin BOM, saltos de linea LF e indentacion con
tabuladores. La estructura de cabecera (ReportFields / DefaultFields) se
reutiliza tal cual de una plantilla exportada por la propia app, para
garantizar que la app lo reconozca.
"""

import os
import re
import uuid
import datetime

import plantilla_incrustada

RUTA_PLANTILLA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plantilla.glf')

# Cada seccion del bloque <Sections> de un invitado: (id_seccion, [(id_campo, titulo, clave)])
SECCIONES = [
    ('0', [('0', 'First Name', 'nombre'), ('1', 'Last Name', 'apellido')]),
    ('3', [('0', 'Phone', 'telefono'), ('6', 'Email', 'email')]),
    ('4', [('0', 'Table', 'mesa')]),
    ('6', [('0', 'Notes', 'notas')]),
]

NULO = '(null)'


def escapar(valor):
    """Valor de campo listo para XML. Vacio se escribe como (null), igual que la app."""
    if valor is None:
        return NULO
    texto = str(valor).strip()
    if not texto or texto == NULO:
        return NULO
    return texto.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def epoch(fecha, hora='18:00'):
    """'2026-08-29' + '18:00' -> segundos epoch como los guarda la app."""
    if isinstance(fecha, datetime.datetime):
        dt = fecha
    else:
        h, m = (hora or '18:00').split(':')[:2]
        y, mes, d = [int(x) for x in str(fecha).split('-')]
        dt = datetime.datetime(y, mes, d, int(h), int(m))
    return dt.timestamp()


def bloque_invitado(g):
    """Genera el XML de un <Guest>. `g` es un dict con las claves de SECCIONES."""
    t = '\t\t\t'
    L = [t + '<Guest>',
         t + '\t<ID>%s</ID>' % (g.get('id') or str(uuid.uuid4()).upper()),
         t + '\t<Arrived>%d</Arrived>' % int(g.get('llego', 0)),
         t + '\t<Invited>%d</Invited>' % int(g.get('invitado', 1)),
         t + '\t<Child>%d</Child>' % int(g.get('nino', 0)),
         t + '\t<Man>%d</Man>' % int(g.get('hombre', 1)),
         t + '\t<ArrivalTime>%s</ArrivalTime>' % escapar(g.get('hora_llegada')),
         t + '\t<Confirmation>%d</Confirmation>' % int(g.get('confirmado', 1)),
         t + '\t<InvitedBy>%s</InvitedBy>' % escapar(g.get('invitado_por')),
         t + '\t<OtherGuestsCount>%d</OtherGuestsCount>' % int(g.get('acompanantes', 0)),
         t + '\t<OtherAttendedGuestsCount>%d</OtherAttendedGuestsCount>' % int(g.get('acompanantes_presentes', 0)),
         t + '\t<Sections>']
    for id_sec, campos in SECCIONES:
        L.append(t + '\t\t<Section>')
        L.append(t + '\t\t\t<ID>%s</ID>' % id_sec)
        L.append(t + '\t\t\t<CustomFields>')
        for id_campo, titulo, clave in campos:
            L += [t + '\t\t\t\t<CustomField>',
                  t + '\t\t\t\t\t<ID>%s</ID>' % id_campo,
                  t + '\t\t\t\t\t<Title>%s</Title>' % titulo,
                  t + '\t\t\t\t\t<Value>%s</Value>' % escapar(g.get(clave)),
                  t + '\t\t\t\t\t<Multiple>0</Multiple>',
                  t + '\t\t\t\t\t<CustomFields />',
                  t + '\t\t\t\t</CustomField>']
        L += [t + '\t\t\t</CustomFields>', t + '\t\t</Section>']
    L += [t + '\t</Sections>', t + '</Guest>']
    return '\n'.join(L) + '\n'


def escribir_glf(destino, invitados, fecha=None, hora='18:00', plantilla=None):
    """Escribe el .glf. `invitados` es una lista de dicts.

    Por defecto usa la cabecera y el pie incrustados en plantilla_incrustada.py, para
    no depender de archivos externos. `plantilla` permite pasar la ruta de otro .glf
    exportado por la app, si algun dia cambia el formato.
    """
    if plantilla:
        with open(plantilla, encoding='utf-8', newline='') as f:
            base = f.read()
        cabecera, resto = base.split('\t\t<Guests>\n', 1)
        _, pie = resto.split('\t\t</Guests>\n', 1)
    else:
        cabecera, pie = plantilla_incrustada.CABECERA, plantilla_incrustada.PIE

    # Si no dan fecha se pone la de hoy: dejar la de la plantilla seria heredar
    # la fecha de otro evento.
    marca = epoch(fecha, hora) if fecha else datetime.datetime.now().timestamp()
    cabecera = re.sub(r'<Date>[^<]*</Date>', '<Date>%f</Date>' % marca, cabecera, count=1)
    pie = re.sub(r'<ModificationDate>[^<]*</ModificationDate>',
                 '<ModificationDate>%f</ModificationDate>' % datetime.datetime.now().timestamp(), pie, count=1)

    cuerpo = ''.join(bloque_invitado(g) for g in invitados)
    contenido = cabecera + '\t\t<Guests>\n' + cuerpo + '\t\t</Guests>\n' + pie

    with open(destino, 'w', encoding='utf-8', newline='') as f:   # UTF-8 sin BOM, LF
        f.write(contenido)
    return destino


def leer_glf(ruta):
    """Devuelve la lista de invitados de un .glf existente (util para verificar)."""
    import xml.etree.ElementTree as ET
    raiz = ET.parse(ruta).getroot()
    fuera = []
    for g in raiz.findall('.//Guests/Guest'):
        d = {'id': g.findtext('ID'),
             'llego': int(g.findtext('Arrived') or 0),
             'invitado': int(g.findtext('Invited') or 1),
             'nino': int(g.findtext('Child') or 0),
             'hombre': int(g.findtext('Man') or 1),
             'confirmado': int(g.findtext('Confirmation') or 0),
             'acompanantes': int(g.findtext('OtherGuestsCount') or 0)}
        titulos = {c: k for _, campos in SECCIONES for _, c, k in campos}
        for s in g.findall('Sections/Section'):
            for c in s.findall('CustomFields/CustomField'):
                clave = titulos.get(c.findtext('Title'))
                if clave:
                    v = c.findtext('Value')
                    d[clave] = '' if v == NULO else v
        fuera.append(d)
    return fuera
