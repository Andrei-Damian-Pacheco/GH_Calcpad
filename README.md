# GH_Calcpad

![Grasshopper](https://img.shields.io/badge/Grasshopper-Rhino%208-green)
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue)
![Version](https://img.shields.io/badge/Version-1.2.0-orange)
![CPD](https://img.shields.io/badge/.cpd-Supported-brightgreen)

**GH_Calcpad v1.2.0 integra el motor de cálculo Calcpad dentro de Grasshopper para ejecutar hojas `.cpd`, modificar variables selectivamente, optimizar y exportar informes técnicos (HTML / PDF / Word).**

---

## 📥 Descargas

<p align="center">
  <!-- Plugin: descarga directa -->
  <a href="https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Plugin/GH_Calcpad.zip">
    <img src="https://img.shields.io/badge/Plugin-.zip%20(v1.2.0)-blue?style=for-the-badge" alt="Download ZIP">
  </a>

  <!-- Manual: VER (vista previa GitHub) -->
  <a href="https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/Instructivo/GH_Calcpad.pdf">
    <img src="https://img.shields.io/badge/Manual-VER-orange?style=for-the-badge" alt="Manual VER">
  </a>

  <!-- Manual: DESCARGA DIRECTA -->
  <a href="https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Instructivo/GH_Calcpad.pdf">
    <img src="https://img.shields.io/badge/Manual-DESCARGAR-orange?style=for-the-badge" alt="Manual DESCARGAR">
  </a>

  <!-- Food4Rhino -->
  <a href="https://www.food4rhino.com/en/app/calcpad">
    <img src="https://img.shields.io/badge/Food4Rhino-Page-green?style=for-the-badge" alt="Food4Rhino">
  </a>

  <!-- Example: descarga directa -->
  <a href="https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Examples/Example_01.cpd">
    <img src="https://img.shields.io/badge/Examples-Example%2001-brightgreen?style=for-the-badge" alt="Example 01">
  </a>
</p>

| Recurso | Descripción | Enlace |
|---------|-------------|--------|
| Plugin (.zip) | `.gha` + DLL necesarias + manual | [Descargar](https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Plugin/GH_Calcpad.zip) |
| Manual PDF | Instructivo GH_Calcpad.pdf | [Ver](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/Instructivo/GH_Calcpad.pdf) / [Descargar](https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Instructivo/GH_Calcpad.pdf) |
| Food4Rhino | Página oficial | [Abrir](https://www.food4rhino.com/en/app/calcpad) |
| Examples   | Hoja de ejemplo `.cpd` | [Descargar](https://raw.githubusercontent.com/Andrei-Damian-Pacheco/GH_Calcpad/master/Examples/Example_01.cpd) |



---

## 🔍 Descripción

Calcpad proporciona un motor de cálculo declarativo con manejo de unidades y generación de resultados. GH_Calcpad lo integra directamente en Grasshopper para habilitar:
- Ejecución nativa de hojas `.cpd`
- Cambios selectivos de variables sin reordenar listas
- Optimización multi-objetivo (Optimizer + Galapagos / Octopus)
- Extracción filtrada de resultados
- Exportación profesional (HTML / PDF / Word)

---

## ✨ Características Principales

| Área | Funcionalidad |
|------|---------------|
| Carga | Lectura de `.cpd`, extracción de variables, valores y unidades |
| Modificación | `Search Variables` para sobrescribir subconjuntos (NaN = no cambiar) |
| Ejecución | `Play CPD` aplica valores y calcula resultados finales |
| Optimización | `Optimizer` genera fitness único + valores objetivos (auto-detección) |
| Filtrado | `Search Results` devuelve solo variables de interés |
| Guardado | `Save CPD` ( `.cpd` / `.txt` ) |
| Exportación | HTML / PDF / Word reutilizando el mismo `UpdatedSheet` |
| Rendimiento | Cache interno para iteraciones GA más rápidas |
| Unidad / Consistencia | Delegado al core de Calcpad |

---

## 🧩 Componentes (v1.2.0)

1. Info  
2. Load CPD  
3. Search Variables  
4. Play CPD  
5. Optimizer  
6. Search Results  
7. Save CPD  
8. Export HTML  
9. Export PDF  
10. Export Word  
11. Help  

---

## 🛠 Requisitos

- Rhino 8 + Grasshopper  
- Calcpad 7+ instalado (recomendado)  

---

## 📦 Instalación

1. Descarga el ZIP (ver sección “Descargas”) y desbloquear.  
2. Extrae y copia la carpeta `GH_Calcpad` a:  
   `C:\Users\<USUARIO>\AppData\Roaming\Grasshopper\Libraries`
3. (Solo si no aparece la pestaña) Propiedades → “Desbloquear” sobre `GH_Calcpad.gha`.
4. Reinicia Rhino y abre Grasshopper.  
5. Verifica la pestaña **Calcpad**.

---

## ⚡ Quick Start

1. Coloca **Load CPD** y asigna ruta a un `.cpd` (por ejemplo, [Example_01.cpd](https://github.com/Andrei-Damian-Pacheco/GH_Calcpad/blob/master/Examples/Example_01.cpd)).  
2. (Opcional) **Search Variables** para modificar algunos parámetros.  
3. Conecta a **Play CPD** → obtienes ecuaciones, valores y unidades.  
4. (Opcional) **Search Results** para filtrar específicos.  
5. Exporta con **Export PDF / Word / HTML**.  

---

## 🔄 Workflows

| Tipo | Secuencia |
|------|----------|
| Básico | Load CPD → Play → Export |
| Modificación selectiva | Load CPD → Search Variables → Play → Export |
| Filtrado de resultados | Load CPD → Play → Search Results → Export |
| Optimización | Load CPD → Optimizer → Galapagos → Save / Export |
| Completo | Load → Search Variables → Play → Search Results → Save → Export |

---

## 🧪 Optimización
Load CPD → Optimizer → Galapagos (Genome = Variable Values) → Mejor fitness → Save / Export
- Deja listas de diseño/objetivos vacías para auto-detección inicial.
- Fitness menor = solución mejor.
- Usa `Convergence Info` y `Status` para detener o ajustar.

---

## 🔧 Buenas Prácticas

- Mantener 1:1: Variables / Values / Units.  
- No reordenar listas originales; `Search Variables` aplica cambios preservando orden.  
- `UpdatedSheet` se reutiliza para todas las exportaciones.  
- Filtra antes de graficar (menos coste en canvas).  
- Nombres de variables finales sin espacios para un filtrado fiable.  
- Verificar siempre salida **Success** en Play / Export / Save.  

---

## 📝 Licencia / Atribución

Calcpad Core se distribuye bajo su propia licencia MIT (ver archivo LICENSE del proveedor).  
GH_Calcpad: ver licencia del repositorio para términos del wrapper y componentes Grasshopper.

---

## 🆘 Soporte

- Componente **Help** dentro de la pestaña Calcpad  
- Issues / sugerencias: [Abrir Issue](../../issues)  

---

Gracias por usar **GH_Calcpad**.
