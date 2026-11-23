# Assembly Sheet Metal DXF Exporter

**Language / Idioma:** English | [Português](#versao-em-portugues)

## Purpose

This SOLIDWORKS VBA macro automates a repetitive manufacturing handoff: it scans the active assembly, finds unique physical parts with sheet-metal Cut List items, exports each eligible flat pattern as a DXF, and names the output using information used by downstream laser-cutting programming.

This package contains copies of the supplied source. The original macro in the workspace was not modified.

## Process and workflow

1. Validate that SOLIDWORKS has a saved assembly open.
2. Enumerate the active assembly's component tree once.
3. Skip suppressed components and descendants, virtual components, components without a file path, and non-part documents. Deduplicate physical parts by full path.
4. Open a referenced part silently when it is not already open, activate it, and switch to the referenced configuration. If one file is referenced in multiple configurations, only the first encountered configuration is exported and the difference is reported.
5. Read existing Cut List properties and quantities. The macro does not call `UpdateCutList`; update the Cut List manually before running if model changes require it.
6. Match each sheet-metal body to its `FlatPattern` using the project's established `FlatPattern -> GetDefinition -> FixedFace2 -> GetBody` association, validate the selected feature, and call `IPartDoc.ExportToDWG2`.
7. Count an export as successful only when the API returns success and the DXF exists with a non-zero file size. Restore the main assembly and report totals in the SOLIDWORKS Immediate Window and a completion dialog.

## Inputs and outputs

- **Input:** a saved `.SLDASM` assembly with referenced `.SLDPRT` files available and Cut Lists prepared.
- **Output:** DXF files in the main assembly folder.
- **Filename pattern:** `Material #Thickness DrawingNumber-CutListItemIndex TotalQuantitypçs.dxf` (thickness is included when available). The legacy `pçs` suffix and filename structure are intentionally preserved because downstream laser-programming steps may depend on them.
- Cut List property keys and detection tokens are intentionally unchanged for compatibility with existing SOLIDWORKS documents. They are data/API inputs, not user-facing documentation.

## Reliability, limits, and operational warning

- The code checks selection count and object identity before exporting, then validates the API return and output file size.
- Repeated instances of the same physical file are consolidated. For different referenced configurations of the same file, only the first encountered configuration is processed.
- A missing or invalid configuration quantity falls back to `1`; a missing or invalid Cut List quantity falls back to the body's count and then to `1`. Verify these source values if totals are unexpected.
- The macro does not save native SOLIDWORKS documents and does not update Cut Lists automatically.
- **Important:** after each processing attempt, the macro closes every processed part document without saving, including parts that were already open before the run. Unsaved changes in those parts can be discarded. The main assembly remains open and is reactivated. This is existing behavior and was not changed in this package.
- The code does not show an explicit preview of same-name DXFs before export. Inspect the destination folder and protect existing outputs before running.
- Diagnostics are written to the VBA Immediate Window; the macro does not create a persistent log file. The runtime-provided `Err.Description` text depends on the local SOLIDWORKS/VBA environment.

## Run and debug

1. Work on a copy of the assembly and review the destination folder.
2. Open the correct SOLIDWORKS VBA project. To update a module, copy the complete source text and paste it into the intended standard module; do not use **Import File** in this workspace because it has caused character corruption.
3. Compile in the VBA editor and confirm the log marker `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL`.
4. Open the Immediate Window (`Ctrl+G` in the VBA editor), run `main`, and review `[VERSION]`, `[SUMMARY]`, `[DXF ERROR]`, `[PART RESULT]`, and close/restore diagnostics.
5. Verify the generated DXFs and their names before using them for manufacturing.

## Validation

- **Reported validation:** marker `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL` was tested in SOLIDWORKS 2025 and reported to work.

## Files in this package

- [AssemblySheetMetalDxfExporter.bas](AssemblySheetMetalDxfExporter.bas) — English comments, diagnostics, and internal helper names.
- [AssemblySheetMetalDxfExporter.pt-BR.bas](AssemblySheetMetalDxfExporter.pt-BR.bas) — Portuguese comments and diagnostics.
- [CONTINUACAO.md](CONTINUACAO.md) — short handoff for future work in this package.
- No `.swp` is included: none was provided or updated. The `.bas` copies and any SOLIDWORKS `.swp` project must be treated as separate artifacts.

---

# Versao em portugues

## Objetivo

Esta macro VBA para SOLIDWORKS percorre a montagem ativa, encontra pecas fisicas unicas com itens de chapa na Cut List, exporta cada planificado elegivel em DXF e nomeia a saida para uso na programacao de corte a laser.

Este pacote contem copias da fonte fornecida. A macro original no workspace nao foi alterada.

## Processo e fluxo

1. Valida se ha uma montagem salva aberta no SOLIDWORKS.
2. Enumera uma vez a arvore de componentes da montagem ativa.
3. Ignora componentes suprimidos e seus descendentes, componentes virtuais, componentes sem caminho e documentos que nao sejam pecas. Consolida pecas fisicas pelo caminho completo.
4. Abre silenciosamente a peca referenciada quando necessario, ativa o documento e seleciona a configuracao referenciada. Se o mesmo arquivo aparecer em configuracoes diferentes, exporta somente a primeira configuracao encontrada e registra a diferenca.
5. Le propriedades e quantidades da Cut List existente. A macro nao chama `UpdateCutList`; atualize a Cut List manualmente antes da execucao se houve alteracoes no modelo.
6. Associa cada corpo de chapa ao `FlatPattern` pela relacao usada neste projeto: `FlatPattern -> GetDefinition -> FixedFace2 -> GetBody`. Valida a feature selecionada e chama `IPartDoc.ExportToDWG2`.
7. So contabiliza sucesso quando a API retorna sucesso e o DXF existe com tamanho maior que zero. Reativa a montagem principal e informa os totais na Immediate Window e em uma mensagem final.

## Entradas e saidas

- **Entrada:** montagem `.SLDASM` salva, pecas `.SLDPRT` referenciadas disponiveis e Cut Lists preparadas.
- **Saida:** arquivos DXF na pasta da montagem principal.
- **Formato do nome:** `Material #Espessura NumeroDesenho-IndiceCutList QuantidadeTotalpçs.dxf` (a espessura aparece quando disponivel). O sufixo legado `pçs` e a estrutura do nome sao preservados intencionalmente porque etapas posteriores da programacao do laser podem depender deles.
- Os nomes das propriedades da Cut List e os termos de deteccao permanecem inalterados por compatibilidade com os documentos SOLIDWORKS existentes. Sao dados/entradas da API, nao texto de interface.

## Confiabilidade, limites e aviso operacional

- O codigo valida a quantidade e a identidade do objeto selecionado antes da exportacao, depois verifica o retorno da API e o tamanho do arquivo.
- Instancias repetidas do mesmo arquivo fisico sao consolidadas. Para configuracoes diferentes do mesmo arquivo, processa somente a primeira configuracao encontrada.
- Se a quantidade da configuracao estiver ausente ou invalida, usa `1`; se a quantidade da Cut List estiver ausente ou invalida, usa a quantidade de corpos e depois `1`. Confira esses valores se os totais parecerem incorretos.
- A macro nao salva documentos nativos do SOLIDWORKS e nao atualiza Cut Lists automaticamente.
- **Importante:** depois de cada tentativa de processamento, a macro fecha cada documento de peca processado sem salvar, inclusive pecas que ja estavam abertas antes da execucao. Alteracoes nao salvas nessas pecas podem ser descartadas. A montagem principal permanece aberta e e reativada. Esse comportamento existente nao foi alterado neste pacote.
- O codigo nao apresenta uma previa explicita de DXFs com o mesmo nome antes da exportacao. Verifique a pasta de destino e proteja arquivos existentes antes de executar.
- Os diagnosticos sao enviados para a Immediate Window; a macro nao cria um arquivo de log persistente. O texto de `Err.Description` depende do ambiente local do SOLIDWORKS/VBA.

## Execucao e depuracao

1. Trabalhe em uma copia da montagem e confira a pasta de destino.
2. Abra o projeto VBA correto do SOLIDWORKS. Para atualizar um modulo, copie o texto completo da fonte e cole no modulo padrao correto; nao use **Import File** neste workspace, pois ja causou corrupcao de caracteres.
3. Compile no editor VBA e confirme o marcador `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL` no log.
4. Abra a Immediate Window (`Ctrl+G` no editor VBA), execute `main` e confira `[VERSION]`, `[SUMMARY]`, `[DXF ERROR]`, `[PART RESULT]` e os diagnosticos de fechamento/restauracao.
5. Verifique os DXFs gerados e seus nomes antes de usa-los na fabricacao.

## Validacao

- **Validacao relatada:** o marcador `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL` foi testado no SOLIDWORKS 2025 e funcionou.

## Arquivos deste pacote

- [AssemblySheetMetalDxfExporter.bas](AssemblySheetMetalDxfExporter.bas) — comentarios, diagnosticos e nomes internos em ingles.
- [AssemblySheetMetalDxfExporter.pt-BR.bas](AssemblySheetMetalDxfExporter.pt-BR.bas) — comentarios e diagnosticos em portugues.
- [CONTINUACAO.md](CONTINUACAO.md) — resumo curto para continuar o trabalho deste pacote.
- Nao ha `.swp`: nenhum foi fornecido ou atualizado. As copias `.bas` e qualquer projeto `.swp` do SOLIDWORKS sao artefatos distintos.
