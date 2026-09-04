<p align="left">
  <a href="https://github.com/Patrick2788/BigSpill-Excel-Lambda-Library">
    <img src="assets/BigSpillLogo.png" alt="BigSpill Logo" width="320">
  </a>
</p>

![Version](https://img.shields.io/badge/BigSpill-v1.1.0-blue)

<p align="left"><em>It began with a big grid…</em></p>

## Introduction

**BigSpill** is a modular Excel LAMBDA library offering elegant, fast, and efficient solutions for grid operations, data analysis, and text processing. It is designed for users who want powerful, pickup-and-go functions without the need to write complex formulas. BigSpill includes primitives, mid‑level operators, and developer‑level tools, all built for modern Excel 365 and fully leveraging dynamic arrays, `LET`, and `LAMBDA`.

## Modules

BigSpill consists of 92 functions across 10 modules:
| Module	| Description |
|---------|-------------|
|  **Gridwork**	|  Foundational coordinate systems for 2D arrays  |
|  **Grid Geometry**	|  Geometric extraction operators for 2D arrays  |
|  **2D Array Shaping**	|  Reshape, resize, wrap, pad, traverse, or reflow grids  |
|  **Grid Algebra**	|  Deletion, selection, filtering, rolling, expansion, and structural mapping  |
|  **Repetition**	|  Replicate arrays along rows, columns, or both  |
|  **Grid Analytics**	|  Sequencing, categorical analysis, neighborhood aggregation, and statistics  |
|  **Text**	|  Structural alignment, extraction, and transformation of text arrays  |
|  **Engineering**	|  Binary, Gray‑code, and bit‑level utilities  |
|  **Combinatorics**	|  High‑performance generators for permutations, combinations, subsets, and related structures  |
|  **Diagnostics**	|  Explain or inspect the structure of a grid  |

## **Compatability**

**BigSpill** requires the modern Excel calculation engine with full support for `LAMBDA`, dynamic arrays, and Element‑to‑Apply (ETA) evaluation used by functions such as `BYROW`, `BYCOL`, and `GROUPBY`.

| Platform | Support |
|---------|-------------|
| Excel 365 (Desktop)  |  Full (Recommended)  |
| Excel 2024 (Desktop)  |  Full*  |
| Excel for the Web  |  Not Supported|

\* _As of July 2026. Excel 2024 will not receive future updates and may not be supported in the future._

## Getting Started

### Quick Start
If you're using Excel 365, all [BigSpill Onboarding Workbooks](/onboarding) include the full library of functions along with demonstrations of each module and every function. These workbooks are the fastest way to explore BigSpill and begin using the tools immediately.

Start here:  [BigSpill Onboarding Workbooks](/onboarding) 

### Installation & Setup
To install BigSpill, import the modules, and verify your environment, please see [Installation & Setup](docs/User%20Manual.md#2-installation--setup) in the user manual.

## License

BigSpill is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.

