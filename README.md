# GH_Calcpad

![GH_Calcpad Logo](https://img.shields.io/badge/Grasshopper-Plugin-green) ![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue) ![Calcpad](https://img.shields.io/badge/Calcpad-Integration-orange)

**Complete Calcpad integration in Grasshopper with revolutionary optimization capabilities for parametric design and engineering calculations.**

## 🚀 Overview

GH_Calcpad is a powerful Grasshopper plugin that seamlessly integrates Calcpad's calculation engine into the Grasshopper environment. This plugin enables engineers and designers to perform complex mathematical calculations, structural analysis, and parametric optimization directly within their Grasshopper workflows.

### Key Features

- **📁 File Loading**: Support for both `.cpd` and `.cpdz` file formats
- **⚡ High-Performance Execution**: Optimized calculation engine for real-time parametric design
- **🔧 Variable Management**: Dynamic variable assignment and unit handling
- **📊 Results Extraction**: Automatic extraction of equations and calculated values
- **📄 Export Capabilities**: Export calculations to Word documents and HTML reports
- **🔍 File Monitoring**: Automatic recomputation when source files change
- **🎯 Optimization Ready**: Designed for integration with optimization algorithms

## 🛠️ Installation

### Prerequisites

- **Rhino 7/8** with Grasshopper
- **.NET Framework 4.8**
- **Calcpad** installed on your system
  - Required files: `Calcpad.Core.dll` and `PyCalcpad.dll`

### Installation Steps

1. Download the latest release from the [Releases](../../releases) page
2. Close Rhino/Grasshopper if running
3. Copy `GH_Calcpad.gha` to your Grasshopper Libraries folder:
4. Restart Rhino
5. The components will appear in the **Calcpad** tab in Grasshopper

## 📋 Components

### 🔧 Core Components

| Component | Description | Inputs | Outputs |
|-----------|-------------|---------|---------|
| **Load CPD** | Loads .cpd files and extracts variables | FilePath, CaptureExplicit | Variables, Values, Units, SheetObj |
| **Load CPDz** | Loads compiled .cpdz files | FilePath, Password | Variables, Values, Units, SheetObj |
| **Play CPD** | Executes calculations with new values | Values, SheetObj | ResultEq, ResultVal, Elapsed, Success, UpdatedSheet |

### 📤 Export Components

| Component | Description | Function |
|-----------|-------------|----------|
| **Export Word** | Exports calculations to Word document | Generate professional reports |
| **Export HTML** | Exports calculations to HTML format | Web-compatible documentation |

### ℹ️ Utility Components

| Component | Description | Purpose |
|-----------|-------------|---------|
| **Info** | Plugin and Calcpad version information | System diagnostics |

## 💡 Usage Examples
