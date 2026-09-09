# Lista de invitados → Guest List

Convierte el Excel de invitados de un evento en el archivo `.glf` que abre la app
**Guest List** del iPad. `README.md` está escrito para el usuario final y explica
muy bien qué formatos de plano entiende; aquí va lo de dentro.

Repositorio: `samechp/lista-a-guest-list`.

## Cómo está montado

```
app.py                  La ventana (tkinter). Punto de entrada.
interfaz.py             Las piezas redondeadas de la ventana.
tema.py                 Los colores. Modo claro y oscuro.
glf.py                  Escribe el archivo .glf.
excel_a_glf.py          Lee el Excel y lo convierte.
plantilla_incrustada.py La plantilla, metida en el código (no se lee de disco).
generar_plantilla.py    Regenera plantilla.glf.
icono.py / hacer_icono.py
construir_exe.py        Empaqueta con pyinstaller.
```

Sin internet, sin instalación, sin Excel instalado: `openpyxl` lee el `.xlsx`
directamente.

## Compilar

```bash
python construir_exe.py
```

Deja `dist/Lista a Guest List.exe`. Requiere `pyinstaller`. `Convertir.bat` es el
atajo desde el Explorador.

## Lo que hay que entender antes de tocar el conversor

**Cada evento llega con un Excel distinto**, porque cada quien lo arma a su
manera. Lo único constante es la hoja del **plano de mesas**, y esa es la fuente:
**cada renglón del plano es una entrada de Guest List**, no una persona.
`Ana Gómez y Pedro Ruiz` es *una* entrada con *+1 acompañante*.

El conteo separa por `y`, `e`, `&`, comas y barras; ignora lo que no es una
persona ("Sra.", "pareja", "hijo", "familia"); descarta paréntesis. Si el plano
trae columnas con números al lado del nombre, **esos mandan** sobre el conteo.

La hoja de contactos es opcional: si está, aporta teléfono, email y
observaciones cruzando por nombre, aguantando tildes, títulos, apellidos
incompletos y errores de tipeo.

Hay comprobación automática: suma las personas de cada mesa y las compara con los
puestos que el propio Excel declara (`(9 PUESTOS)`, `Total: 11`). Si no cuadra,
lo dice con el número de mesa.

**Antes de "arreglar" el parseo, probar con `ejemplo/Boda de ejemplo.xlsx`.** Las
reglas raras están ahí porque un evento real las necesitó.

## Idioma

**Todo en español** y sin tildes en los nombres de archivo y de función. Los
textos de la ventana se escriben para alguien que no programa: "Qué revisar", no
"Validaciones".
