# SOLIDWORKS VBA and Process Automation Portfolio

This repository brings together practical SOLIDWORKS VBA macros and notes from their development. It is intended as a useful reference for people working with similar tasks and as a record of the technical work behind each macro.

## What you'll find

- Automate repetitive engineering and manufacturing handoffs.
- Integrate SOLIDWORKS VBA/COM APIs with document properties, assemblies, Cut Lists, flat patterns, and file export.
- Add defensive validation and Immediate Window diagnostics.
- Preserve native SOLIDWORKS documents and clearly communicate save/close behavior.

## Projects

| Project | Domain | Status | Evidence |
|---|---|---|---|
| [Assembly Sheet Metal DXF Exporter](AssemblySheetMetalDxfExporter/) | Assembly export automation | Bilingual package prepared | Static executable-code comparison passed; SOLIDWORKS test is user-reported |

Each macro has its own folder, named after the macro, with its source and supporting documentation.

## Repository structure

```text
.
├── README.md
├── .gitignore
├── CONTRIBUTING.md
└── <MacroName>/
    ├── <MacroName>.bas
    ├── <MacroName>.pt-BR.bas
    ├── README.md
    └── CONTINUACAO.md
```

The `.bas` files are source copies. Each package README identifies its source marker, validation status, limitations, and operating steps.

## Validation policy

This repository distinguishes:

- **Static review:** source structure and executable-code comparison were reviewed.
- **Compilation:** the VBA project was compiled successfully in the reported environment.
- **Observed execution:** the macro was run successfully in SOLIDWORKS and the result was inspected.
- **User-reported execution:** execution was reported by the project owner but not independently observed by the maintainer.

No package will present compilation or execution as proven unless that evidence is recorded.

## Before using a macro

1. Work on a copy of the SOLIDWORKS assembly or folder.
2. Review save, close, overwrite, export, and restore behavior.
3. Compile the module in the intended VBA project.
4. Verify the documented source marker.
5. Run in a controlled test set before production use.
6. Inspect generated files and logs.

## Diagnostics

Where supported, macros write processing steps and failures to the SOLIDWORKS VBA Immediate Window. Record the source marker, SOLIDWORKS version, and relevant log lines when documenting an issue. Remove project-sensitive names and paths before sharing logs outside the project.

## License and confidentiality

No license is currently included. This intentionally avoids implying redistribution terms before the repository owner selects one.

Remove confidential models, drawing numbers, customer names, network paths, internal properties, and proprietary manufacturing data before publishing. The repository contains sample-oriented documentation and source copies, not production customer datasets.

---

# Portfolio de automacao VBA/SOLIDWORKS

Este repositorio reune macros VBA para SOLIDWORKS e notas sobre seu desenvolvimento. A ideia e compartilhar solucoes uteis para tarefas recorrentes e registrar as decisoes tecnicas, limites e validacao de cada macro.

Cada macro fica em uma pasta com seu proprio nome e os arquivos relacionados.

Remova dados confidenciais antes de distribuir os arquivos.
