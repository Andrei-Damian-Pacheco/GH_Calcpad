# GH_Calcpad

![Grasshopper](https://img.shields.io/badge/Grasshopper-Rhino%208-green)
![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue)
![Version](https://img.shields.io/badge/Version-1.2.0-orange)
![Status](https://img.shields.io/badge/.cpd-Supported-brightgreen)
![Status](https://img.shields.io/badge/.cpdz-Experimental-lightgrey)

**GH_Calcpad v1.2.0 embeds the Calcpad calculation engine inside Grasshopper for engineering-grade parametric computation, selective variable editing, multi‑objective optimization and professional report export.**

---

## 🔍 What It Does

Calcpad sheets (`.cpd`) can be executed directly inside Grasshopper:
- Override selected input variables
- Recompute on demand or under optimization
- Extract specific result variables
- Export structured reports (HTML / PDF / Word)

> Current full support: **.cpd**. The loader for **.cpdz** is present but treated as *experimental* (only works if textual source is available). In this build the `Load CPDz` component is excluded from compilation.

---

## ✨ Key Features

| Area | Feature |
|------|---------|
| File Loading | Native `.cpd` sheet loading with variable & unit extraction |
| Variable Editing | `Search Variables` component: targeted overrides without breaking ordering |
| Execution | `Play CPD`: applies overrides (NaN = skip) and evaluates results |
| Optimization | `Optimizer`: auto‑detect design variables & objectives, single fitness for Galapagos / Octopus |
| Result Filtering | `Search Results`: extract only the result names you care about |
| Reporting | Export to **HTML**, **PDF**, **Word (.docx)** |
| Saving | Save modified sheet as `.cpd` or `.txt` |
| Caching | Smart reuse of internal structures for faster iterative / GA runs |
| Units & Equations | Delegated to Calcpad core (consistent unit and expression handling) |

---

## 🧩 Components (v1.2.0)

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

*(Load CPDz component exists in source but is excluded in this build.)*

---

## 🛠 Requisitos

- Rhino 8 (Grasshopper)
- .NET Framework 4.8
- Calcpad 7+ instalado (recomendado; aporta runtime y validación de hojas)
- Windows (x64)

---

## 📦 Instalación

1. Descarga el ZIP desde [Releases](../../releases).  
2. Extrae el ZIP.  
3. Copia la carpeta `GH_Calcpad` (y su contenido) a tu carpeta de librerías Grasshopper:  
   `C:\Users\<TU_USUARIO>\AppData\Roaming\Grasshopper\Libraries`
4. (Solo si no aparece la pestaña) Clic derecho en `GH_Calcpad.gha` → Propiedades → Desbloquear.
5. Reinicia Rhino → abre Grasshopper → pestaña **Calcpad**.

---

## ⚡ Quick Start

1. Coloca el componente **Load CPD** y conecta la ruta a un archivo `.cpd`.
2. (Opcional) Usa **Search Variables** para sobrescribir algunos valores.
3. Conecta a **Play CPD** para ejecutar.
4. Usa **Search Results** si quieres filtrar resultados por nombre.
5. Exporta con **Export PDF** / **Export Word** / **Export HTML**.

---

## 🧪 Optimization Workflow
Load CPD → Optimizer → (Galapagos / Octopus) → Play CPD (implicit) → Best solution → Save / Export
- Deja vacías las listas de diseño/objetivos para auto‑detección inicial.
- Usa la salida Fitness como objetivo primario (menor = mejor).

---

## 🔧 Best Practices

- Mantén correspondencia 1:1: Variables / Values / Units.
- Usa **NaN** en la lista “Values” de Play CPD para no modificar una posición.
- Reutiliza la salida `UpdatedSheet` para todas las exportaciones en cadena.
- Aplica **Search Results** antes de graficar en Grasshopper para reducir coste.
- Nombres de variables finales: sin espacios (facilita filtrado).

---

## 🗂 Example Workflows

| Workflow | Secuencia |
|----------|-----------|
| Básico | Load CPD → Play → Export |
| Edición selectiva | Load CPD → Search Variables → Play → Export |
| Filtrado | Load CPD → Play → Search Results → Export |
| Optimización | Load CPD → Optimizer → Galapagos → Save / Export |
| Completo | Load → Search Variables → Play → Search Results → Save → PDF / Word |

---

## 🧭 Roadmap (Resumen)

| Estado | Elemento |
|--------|----------|
| En curso | Mejora de compatibilidad `.cpdz` |
| Planificado | Métricas avanzadas de convergencia en Optimizer |
| Planificado | Carpeta oficial de ejemplos (parametric + optimization) |
| Evaluación | Exportación incremental / diff |

---

## 📝 License / Atribución

Calcpad Core (Calcpad.Core.dll) se distribuye bajo su propia licencia (MIT según archivo LICENSE del proveedor).  
Este plugin integra la API expuesta; revisa la licencia del repositorio para condiciones específicas del wrapper GH_Calcpad.

---

## 🆘 Support

- Revisa el componente **Help** dentro de la pestaña
- Issues / sugerencias: abre un [Issue](../../issues)

---

## ⚠ Disclaimer

`.cpdz` compiled packages: soporte parcial / experimental en esta versión. Si necesitas plena compatibilidad, usa `.cpd` por ahora.

---

## 🔗 GUID (Assembly)

`1A9A80A9-0512-4001-8A73-1FECB82117B7`

---

Gracias por usar **GH_Calcpad**. Mejora tus flujos de cálculo e ingeniería dentro de Grasshopper de forma directa y reproducible.
