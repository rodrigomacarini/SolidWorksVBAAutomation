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

# Versão em português

## Objetivo

Esta macro VBA para SOLIDWORKS percorre a montagem ativa, encontra peças físicas únicas com itens de chapa na Cut List, exporta cada planificado elegível em DXF e nomeia os arquivos de saída para uso na programação de corte a laser.

Este pacote contém cópias da fonte fornecida. A macro original no workspace não foi alterada.

## Processo e fluxo

1. Valida se há uma montagem salva aberta no SOLIDWORKS.
2. Enumera uma vez a árvore de componentes da montagem ativa.
3. Ignora componentes suprimidos e seus descendentes, componentes virtuais, componentes sem caminho e documentos que não sejam peças. Consolida as peças físicas pelo caminho completo.
4. Abre silenciosamente a peça referenciada, se necessário, ativa o documento e seleciona a configuração correspondente. Se o mesmo arquivo aparecer em configurações diferentes, exporta somente a primeira configuração encontrada e registra essa diferença.
5. Lê as propriedades e quantidades da Cut List existente. A macro não chama `UpdateCutList`; atualize a Cut List manualmente antes de executar a macro se o modelo tiver sido alterado.
6. Associa cada corpo de chapa ao `FlatPattern` pela relação usada neste projeto: `FlatPattern -> GetDefinition -> FixedFace2 -> GetBody`. Valida o recurso selecionado e chama `IPartDoc.ExportToDWG2`.
7. Só contabiliza uma exportação como bem-sucedida quando a API retorna sucesso e o DXF existe com tamanho maior que zero. Reativa a montagem principal e informa os totais na Immediate Window e em uma mensagem final.

## Entradas e saídas

- **Entrada:** uma montagem `.SLDASM` salva, com as peças `.SLDPRT` referenciadas disponíveis e as Cut Lists preparadas.
- **Saída:** arquivos DXF na pasta da montagem principal.
- **Formato do nome:** `Material #Espessura NumeroDesenho-IndiceCutList QuantidadeTotalpçs.dxf` (a espessura é incluída quando disponível). O sufixo legado `pçs` e a estrutura do nome são preservados intencionalmente, pois etapas posteriores da programação do laser podem depender deles.
- Os nomes das propriedades da Cut List e os termos de detecção permanecem inalterados por compatibilidade com os documentos SOLIDWORKS existentes. Eles são dados de entrada da API, não texto de interface.

## Confiabilidade, limites e aviso operacional

- O código valida a quantidade e a identidade do objeto selecionado antes da exportação; em seguida, verifica o retorno da API e o tamanho do arquivo.
- Instâncias repetidas do mesmo arquivo físico são consolidadas. Se o mesmo arquivo for referenciado em configurações diferentes, somente a primeira configuração encontrada será processada.
- Se a quantidade da configuração estiver ausente ou for inválida, a macro usa `1`. Se a quantidade da Cut List estiver ausente ou for inválida, usa a quantidade de corpos e, depois, `1`. Confira esses valores se os totais parecerem incorretos.
- A macro não salva documentos nativos do SOLIDWORKS nem atualiza Cut Lists automaticamente.
- **Importante:** após cada tentativa de processamento, a macro fecha sem salvar cada documento de peça processado, inclusive os que já estavam abertos antes da execução. Alterações não salvas nesses documentos podem ser descartadas. A montagem principal permanece aberta e é reativada. Esse comportamento preexistente não foi alterado neste pacote.
- O código não exibe uma prévia explícita dos DXFs com nomes repetidos antes da exportação. Verifique a pasta de destino e proteja os arquivos existentes antes de executar a macro.
- Os diagnósticos são enviados para a Immediate Window; a macro não cria um arquivo de log persistente. O texto de `Err.Description` depende do ambiente local do SOLIDWORKS/VBA.

## Execução e depuração

1. Trabalhe em uma cópia da montagem e confira a pasta de destino.
2. Abra o projeto VBA correto do SOLIDWORKS. Para atualizar um módulo, copie o texto completo da fonte e cole-o no módulo padrão correto. Não use **Import File** neste workspace, pois esse recurso já causou corrupção de caracteres.
3. Compile no editor VBA e confirme no log o marcador `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL`.
4. Abra a Immediate Window (`Ctrl+G` no editor VBA), execute `main` e confira `[VERSION]`, `[SUMMARY]`, `[DXF ERROR]`, `[PART RESULT]` e os diagnósticos de fechamento e restauração.
5. Verifique os DXFs gerados e seus nomes antes de usá-los na fabricação.

## Validação

- **Validação relatada:** o marcador `20261106-02-DXF-MONTAGEM-PASTA-PRINCIPAL` foi testado no SOLIDWORKS 2025 e, segundo o relato, funcionou.

## Arquivos deste pacote

- [AssemblySheetMetalDxfExporter.bas](AssemblySheetMetalDxfExporter.bas) — comentários, diagnósticos e nomes internos em inglês.
- [AssemblySheetMetalDxfExporter.pt-BR.bas](AssemblySheetMetalDxfExporter.pt-BR.bas) — comentários e diagnósticos em português.
- [CONTINUACAO.md](CONTINUACAO.md) — resumo curto para dar continuidade ao trabalho neste pacote.
- Não há arquivo `.swp`: nenhum foi fornecido ou atualizado. As cópias `.bas` e qualquer projeto `.swp` do SOLIDWORKS são artefatos distintos.
