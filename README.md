<p align="left">
  <img src="assets/BigSpill Logo.png" alt="BigSpill Logo" width="320">
</p>

A suite of Excel Lambda functions for grid shaping, analytics, algebra, geometry, repetition, text processing, engineering, combinatorics, and diagnostics.

Introduction
BigSpill is a modular Excel LAMBDA library providing fast, efficient, and reliable functions for grid operations, data analysis, engineering, and text processing.  
It is designed for users who want powerful, pickup‑and‑go functions without the need to write complex formulas.
BigSpill’s modular organization spans from primitives to mid‑level operators to developer‑level functions, giving users a consistent toolkit at every layer of complexity.
The library is built for modern Excel 365 and takes full advantage of the innovations introduced with dynamic arrays, LET, and LAMBDA.


BigSpill Compatibility
BigSpill is designed for the modern Excel calculation engine and requires full support for LAMBDA, dynamic arrays, and advanced function evaluation (AFE). Compatibility varies by platform.

Excel 365 (Desktop)
Fully supported.
All BigSpill functions run correctly on:

Current Channel

Monthly Enterprise Channel

Semi‑Annual Enterprise Channel

Excel 365 receives ongoing feature updates, including new dynamic array functions and Lambda helpers. Future BigSpill versions may rely on these updates.

Excel 2024 (Desktop)
Supported with caveats.
Excel 2024 includes all Lambda and dynamic array features available at the time of its release, and BigSpill is fully compatible with those features.

However:

Excel 2024 does not receive new functions after release.

Future BigSpill versions may require features available only in Excel 365.

Excel for the Web
Not supported.
While Excel for the web includes partial LAMBDA functionality, it has significant limitations:

ETA is not supported

Implicit aggregation often fails

Certain nested or recursive Lambdas do not evaluate correctly

AFE modules may not load or resolve reliably

BigSpill uses advanced Lambda patterns that Excel for the web cannot consistently execute.

Excel 2019 / Excel 2021
Not supported.
These versions do not include LAMBDA or modern dynamic array functions. BigSpill functions will return #NAME?.

Google Sheets and other spreadsheet apps
Not supported.
Google Sheets does not implement LAMBDA, dynamic arrays, or AFE. BigSpill is an Excel‑native library.

Summary
Fully supported: Excel 365 (desktop)

Supported today, but static: Excel 2024

Not supported: Excel for the web, Excel 2019/2021, Google Sheets

BigSpill relies on modern Excel features and may adopt new Excel 365 functions in future releases.
