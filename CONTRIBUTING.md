# Contributing and adding a macro

## Package requirements

For every macro project, create one directory with a professional English name:

```text
<ProfessionalMacroName>/
```

Include:

- `<ProfessionalMacroName>.bas`
- `<ProfessionalMacroName>.pt-BR.bas` when bilingual documentation is useful
- `README.md`
- `CONTINUACAO.md`
- `<ProfessionalMacroName>.swp` only if it was supplied or explicitly updated

Do not fabricate or claim synchronization with a `.swp` file.

## Source rules

- Preserve functional VBA logic and compatibility-critical SOLIDWORKS property names.
- Add comments that explain purpose, API assumptions, state changes, validation, and failure behavior.
- Use ASCII in newly created VBA modules when practical.
- Preserve `ChrW(...)` for SOLIDWORKS property names that require special characters.
- Keep the source marker/version in the macro and documentation.
- Do not use VBA `Dir$` for recursive or nested folder traversal; prefer FSO, collections, or explicit traversal state.
- Do not alter executable logic solely to translate comments or diagnostics.

## Validation labels

Use only the status that is supported by evidence:

- Static review
- VBA compilation observed
- SOLIDWORKS execution observed
- User-reported; not independently observed

## Publishing checklist

- Remove confidential drawing and project data.
- Remove customer names and internal network paths.
- Verify save/close behavior and disclose unsaved-data risks.
- Verify generated-file validation and known limits.
- Remove temporary source generators and archive files.
- Record a source marker and test scope.
- Do not publish an unmeasured time or reliability gain.
- Decide and add a license before redistribution.
