# GH_Calcpad

![Grasshopper](https://img.shields.io/badge/Grasshopper-Rhino%208-green)
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue)
![Version](https://img.shields.io/badge/Version-2.0.0-orange)
![CPD](https://img.shields.io/badge/.cpd-Supported-brightgreen)

**GH_Calcpad v2.0.0 integra el motor de cálculo Calcpad dentro de Grasshopper para ejecutar hojas `.cpd`, modificar variables selectivamente, optimizar (Galapagos / Wallacei / Octopus) y exportar informes técnicos (HTML / PDF / Word) idénticos a los de la app oficial de Calcpad.**

---

## 📥 Descargas

<p align="center">
  <!-- Plugin: GitHub Releases -->
  <a href="https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/releases/latest">
    <img src="https://img.shields.io/badge/Plugin-.zip%20(v2.0.0)-blue?style=for-the-badge" alt="Download ZIP">
  </a>
  <!-- Manual: VER (vista previa en GitHub) -->
  <a href="https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/GH_Calcpad/Documents/Instructivo_GH-Calcpad_V2.pdf">
    <img src="https://img.shields.io/badge/Manual-VER-orange?style=for-the-badge" alt="Manual VER">
  </a>
  <!-- Food4Rhino -->
  <a href="https://www.food4rhino.com/en/app/calcpad">
    <img src="https://img.shields.io/badge/Food4Rhino-Page-green?style=for-the-badge" alt="Food4Rhino">
  </a>
  <!-- Examples -->
  <a href="https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/tree/master/GH_Calcpad/Examples">
    <img src="https://img.shields.io/badge/Examples-4%20sheets-brightgreen?style=for-the-badge" alt="Examples">
  </a>
</p>

| Recurso | Descripción | Enlace |
|---------|-------------|--------|
| Plugin (.zip) | `.gha` + worker autocontenido + manual | [Ver](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/releases/latest) / [Descargar](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/releases/latest/download/GH_Calcpad.zip) |
| Manual PDF | Instructivo GH_Calcpad V2.pdf | [Ver](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/GH_Calcpad/Documents/Instructivo_GH-Calcpad_V2.pdf) / [Descargar](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/raw/master/GH_Calcpad/Documents/Instructivo_GH-Calcpad_V2.pdf) |
| Food4Rhino | Página oficial | [Abrir](https://www.food4rhino.com/en/app/calcpad) |
| Examples | 4 hojas `.cpd` + `.gh` de ejemplo | [Ver carpeta](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/tree/master/GH_Calcpad/Examples) |

---

## 🔍 Descripción

Calcpad proporciona un motor de cálculo declarativo con manejo de unidades y generación de resultados. GH_Calcpad lo integra directamente en Grasshopper para habilitar:
- Ejecución nativa de hojas `.cpd`, con el motor real de Calcpad (vendored, sin dependencia de una instalación externa)
- Cambios selectivos de variables por nombre, sin reordenar listas
- Optimización multi-objetivo vía Galapagos / Wallacei / Octopus (convención `gh_` / `ghc_`, sin componente Optimizer dedicado)
- Extracción filtrada de resultados, leída directamente del motor de cálculo
- Exportación profesional (HTML / PDF / Word) — el mismo pipeline y plantillas que usa la app de escritorio de Calcpad

---

## ✨ Características Principales

| Área | Funcionalidad |
|------|---------------|
| Carga | Lectura de `.cpd`/`.txt`, extracción de variables, valores y unidades; auto-refresco si el archivo cambia en disco |
| Modificación | `Search Variables` sobrescribe variables por nombre (auto-detecta el prefijo `gh_` si no se especifica) |
| Ejecución | `Play CPD` calcula la hoja a través del motor real de Calcpad |
| Optimización | Sin componente dedicado: `Play CPD` + `Search Variables`/`Search Results` se conectan directo al genoma/fitness de Galapagos, Wallacei u Octopus |
| Filtrado | `Search Results` devuelve solo las variables de interés (auto-detecta el prefijo `ghc_`) |
| Guardado | `Save CPD/TXT` (`.cpd` / `.txt`), escritura segura (archivo temporal + reemplazo) |
| Exportación | HTML / PDF / Word — mismo motor, plantilla y pipeline que la app oficial de Calcpad |
| Arquitectura | Motor de cálculo en un proceso worker separado (.NET 10, autocontenido), el plugin nunca referencia Calcpad en proceso |

---

## 🧩 Componentes (v2.0.0)

1. Calcpad Info
2. Load CPD
3. Search Variables
4. Play CPD
5. Search Results
6. Save CPD/TXT
7. Export HTML
8. Export PDF
9. Export Word
10. Calcpad Help

---

## 🛠 Requisitos

- Rhino 8 + Grasshopper
- .NET Framework 4.8 (normalmente ya presente en Windows)
- No requiere ninguna instalación separada de Calcpad — el motor de cálculo viaja incluido (vendored) dentro del plugin

---

## 📦 Instalación

1. Descarga el ZIP (ver sección "Descargas") y desbloquéalo si Windows lo marca como bloqueado.
2. Extrae y copia la carpeta `GH_Calcpad` completa a:
   `C:\Users\<USUARIO>\AppData\Roaming\Grasshopper\Libraries`
3. (Solo si no aparece la pestaña) Propiedades → "Desbloquear" sobre `GH_Calcpad.gha` y sus DLLs.
4. Reinicia Rhino y abre Grasshopper.
5. Verifica la pestaña **Calcpad**.

---

## ⚡ Quick Start

1. Coloca **Load CPD** y asigna ruta a un `.cpd` (por ejemplo, [Example_01.cpd](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/GH_Calcpad/Examples/Example_01.cpd)).
2. (Opcional) **Search Variables** para modificar parámetros por nombre.
3. Conecta a **Play CPD** → calcula la hoja con el motor real de Calcpad.
4. (Opcional) **Search Results** para filtrar resultados específicos.
5. Exporta con **Export PDF / Word / HTML**.

---

## 🔄 Workflows

| Tipo | Secuencia |
|------|-----------|
| Básico | Load CPD → Play CPD → Export/Save |
| Selectivo | Load CPD → Search Variables → Play CPD → Export/Save |
| Completo | Load CPD → Search Variables → Play CPD → Search Results → Export/Save |
| Optimización | Igual al completo, dejando ambos "Filter Names" vacíos (auto-detección `gh_`/`ghc_`), conectado directo al genoma/fitness de Galapagos, Wallacei u Octopus |

---

## 🧪 Optimización

Load CPD → Search Variables (Filter Names vacío) → Play CPD → Search Results (Filter Names vacío) → genoma/fitness del optimizador externo.

- No existe un componente "Optimizer" propio: Galapagos/Wallacei/Octopus son el optimizador; GH_Calcpad solo convierte cada genoma en un valor de fitness.
- Prefija tus variables de diseño con `gh_` y tus resultados/objetivos con `ghc_` para que ambos Search se auto-detecten sin escribir nombres a mano.
- Verifica siempre **Success** en Play CPD antes de confiar en los resultados — penaliza las iteraciones fallidas en vez de dejar que un NaN se propague silenciosamente.
- Ver [Example_04-Optimization.cpd](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/GH_Calcpad/Examples/Example_04-Optimization.cpd) para un caso completo (optimización de sección de viga).

---

## 🔧 Buenas Prácticas

- Prefija variables de diseño con `gh_` y resultados/objetivos con `ghc_` para habilitar la auto-detección en Search Variables/Search Results.
- Usa el carácter de prima real (′), nunca un apóstrofo recto ('): en Calcpad, `'` siempre inicia un comentario.
- Verifica siempre la salida **Success** en Play CPD y en cada componente de Export/Save antes de confiar en los resultados posteriores.
- **Not Found** en Search Variables/Search Results señala un nombre de variable mal escrito o inexistente.
- El `Sheet` calculado se reutiliza tal cual para todas las exportaciones — no hace falta recalcular por cada formato.

---

## 📝 Licencia / Atribución

Calcpad Core se distribuye bajo su propia licencia MIT (ver `THIRD-PARTY-NOTICES.txt` para el detalle completo de las dependencias vendored, incluyendo Calcpad.Core, Calcpad.OpenXml y wkhtmltopdf).
GH_Calcpad: ver [LICENSE](LICENSE) del repositorio para los términos del wrapper y los componentes de Grasshopper.

---

## 🆘 Soporte

- Componente **Calcpad Help** dentro de la pestaña Calcpad
- Issues / sugerencias: [Abrir Issue](../../issues)

---

Gracias por usar **GH_Calcpad**.
