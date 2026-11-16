---
version: 1
name: BiciFriendly-Ciclovia
description: "Señalética de calle para un mapa de lugares bikefriendly en Santiago. Día: fondo blanco, letreros negros, naranjo #E85D00 solo como carril y llamado a la acción. Noche: el mismo sistema invertido. Categorías por forma, no por color."

colors:
  # Día (html[data-theme="light"])
  canvas: "#FFFFFF"
  map: "#EEF1F3"
  surface: "#F2F4F5"
  surface-2: "#E6E9EC"
  hairline: "#D5DADE"
  rule: "#000000"
  ink: "#000000"
  ink-muted: "#4F565C"
  ink-dim: "#5A6168"
  orange: "#E85D00"        # superficies: carril, descuento, pin elegido, indicador de pestaña
  orange-text: "#A84300"   # todo texto naranjo sobre blanco (6:1)
  sign: "#000000"
  sign-ink: "#FFFFFF"
  sign-accent: "#FF7A1A"
  cta: "#000000"
  cta-ink: "#FFFFFF"
  # Noche (:root)
  night-canvas: "#000000"
  night-surface: "#1A1A1A"
  night-ink: "#F5F5F5"
  night-orange: "#FF7A1A"
  night-sign: "#F2F4F5"
  night-sign-accent: "#A84300"
  night-cta: "#FF7A1A"
  night-cta-ink: "#000000"
  location: "#1C6FD1"

typography:
  display: "Unbounded 800 (títulos, nombres, cifras, botones)"
  body: "Overpass 400–800 (texto, etiquetas, chips)"
  scale: [12, 14, 16, 20, 24, 30]
  display-tracking: "-0.02em a -0.03em"

spacing: "múltiplos de 4: 4, 8, 12, 16, 20, 24, 28"
radius: 0
shadow: "una sola: 0 12px 32px var(--shadow) para ficha flotante y hoja móvil"
breakpoints: { mobile: "≤820px", tablet: "821–1100px", desktop: ">1100px" }
---

## Idea

El sitio se lee como señalética de ciclovía. Se usa en la calle, a menudo a pleno sol, así que el modo día manda: blanco, negro y un solo naranjo. La identidad está en tres gestos, nada más:

1. **El carril**: banda naranja de 10px con línea segmentada blanca en el borde izquierdo de la lista.
2. **El letrero**: la cabecera de cada ficha es un bloque negro (claro de noche) con tipo de lugar, nombre y distancia en naranjo.
3. **Las formas**: cada categoría es una forma geométrica, igual en filtros, lista, pines y ficha.

## Categorías

| Tipo | Forma | Clase |
|---|---|---|
| Café | círculo | `.t-cafe .gl` |
| Bar | cuadrado | `.t-bar .gl` |
| Comida rápida | triángulo | `.t-rapida .gl` |
| Restaurant | rombo | `.t-restaurant .gl` |
| Punto ciclista | cruz | `.t-punto .gl` |

No volver a colores por categoría: el naranjo es de la marca y no puede significar "café".

## Componentes

- **Chips de filtro**: rectángulo con borde 2px tinta; activo = fondo tinta, texto fondo, forma naranja.
- **Botones**: rectángulos. Principal = `--cta` (negro de día, naranjo de noche). Secundario = borde 2px tinta.
- **Fila de lista**: cuadrado 44px con la forma, nombre en Unbounded 16, meta en Overpass 14, distancia a la derecha en Unbounded 20. Elegida = franja naranja interior de 6px y cuadrado naranjo.
- **Pin**: cuadrado 30px, fondo tinta, forma en blanco, borde 3px del color de fondo. Elegido = naranjo y 1,35×.
- **Ficha**: letrero arriba, dirección, bloque naranja de descuento (texto negro), valoración, "Para ciclistas" en mosaicos con borde izquierdo de 4px, fuente, botones.
- **Barra inferior (celular)**: activa = texto tinta + raya naranja superior de 4px; "Sugerir lugar" en `--cta`.

## Sí

- Texto naranjo siempre con `--or-text` (contraste AA en ambos modos).
- Líneas de estructura de 2px en `--rule` (negro de día) para encabezado, barra y bordes de panel.
- Esquinas rectas en todo.
- Separar datos con coma, no con punto medio.

## No

- Fondo crema, títulos en cursiva, botones con forma de píldora.
- Naranjo #E85D00 como color de texto sobre blanco (3,5:1).
- Nuevos tamaños de letra fuera de la escala, ni medios píxeles.
- Sombras adicionales o degradados decorativos.
- Mayúsculas sostenidas en etiquetas.

## Dónde vive

Fuente: `template.html` (bloque `<style>` después del CSS de Leaflet). Se compila con `build.py` a `index.html`.
