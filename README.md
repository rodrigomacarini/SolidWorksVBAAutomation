# SOLIDWORKS VBA and Process Automation Portfolio

This repository presents practical process-automation solutions built for SOLIDWORKS and supported by VBA. It is designed to show programming ability, software integration with real engineering workflows, debugging practices, and clear technical communication for teams outside mechanical design.

## Portfolio focus

- Automate repetitive engineering and manufacturing handoffs.
- Integrate SOLIDWORKS VBA/COM APIs with document properties, assemblies, Cut Lists, flat patterns, and file export.
- Add defensive validation and Immediate Window diagnostics.
- Preserve native SOLIDWORKS documents and clearly communicate save/close behavior.
- Document process value without inventing cost or time-savings metrics.

## Published projects

| Project | Domain | Status | Evidence |
|---|---|---|---|
| [Assembly Sheet Metal DXF Exporter](AssemblySheetMetalDxfExporter/) | Assembly export automation | Bilingual package prepared | Static executable-code comparison passed; SOLIDWORKS test is user-reported |

Additional macro projects can be added here through the same package structure.

## Repository structure

```text
.
├── README.md
├── .gitignore
├── CONTRIBUTING.md
└── <ProfessionalMacroName>/
    ├── <ProfessionalMacroName>.bas
    ├── <ProfessionalMacroName>.pt-BR.bas
    ├── README.md
    └── CONTINUACAO.md
```

The `.bas` files are presentation copies. Each copy retains compatibility-critical SOLIDWORKS property names and API logic. The package README identifies source markers, validation status, limitations, and debugging steps.

## Validation policy

This repository distinguishes:

- **Static review:** source structure and executable-code comparison were reviewed.
- **Compilation:** the VBA project was compiled successfully in the reported environment.
- **Observed execution:** the macro was run successfully in SOLIDWORKS and the result was inspected.
- **User-reported execution:** execution was reported by the project owner but not independently observed by the maintainer.

No package will present compilation or execution as proven unless that evidence is recorded.

## Process-impact measurement

Performance claims require observed measurements. For each comparable workflow, record:

```text
time saved per cycle = observed manual time - observed automated time
time saved in period = time saved per cycle × confirmed cycle count
reliability change = documented errors/rework before - after
```

The sample size, period, environment, and measurement method must accompany each claim.

## Before using a macro

1. Work on a copy of the SOLIDWORKS assembly or folder.
2. Review save, close, overwrite, export, and restore behavior.
3. Compile the module in the intended VBA project.
4. Verify the documented source marker.
5. Run in a controlled test set before production use.
6. Inspect generated files and logs.

## Debugging

Where supported, macros write steps and failures to the SOLIDWORKS VBA Immediate Window. README files include prompts that can be reused with an AI assistant. Prompts should include the macro marker, anonymized model/configuration context, relevant log lines, SOLIDWORKS version, and the earliest failing step. Do not share customer names, proprietary drawing data, or network paths.

## License and confidentiality

No license is currently included. This intentionally avoids implying redistribution terms before the repository owner selects one.

Remove confidential models, drawing numbers, customer names, network paths, internal properties, and proprietary manufacturing data before publishing. The repository contains sample-oriented documentation and source copies, not production customer datasets.

---

# Portfolio de Automacao VBA/SolidWorks

Este repositorio apresenta solucoes praticas de automacao de processos desenvolvidas para SOLIDWORKS com VBA. O objetivo e demonstrar capacidade de programacao, integracao com fluxos reais de engenharia, depuracao, confiabilidade e comunicacao tecnica para equipes fora da area de projetos mecanicos.

Cada nova macro devera ter sua propria pasta com nome profissional, copia em ingles, copia em portugues quando aplicavel, README, status de validacao e arquivo curto de continuidade.

Nao publicar licenca, dados confidenciais ou alegacoes de economia ate revisar os termos e validar as evidencias.
