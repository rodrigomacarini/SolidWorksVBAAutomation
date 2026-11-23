# Assembly Drawing PDF Exporter

**Language / Idioma:** English | [Português](#versao-em-portugues)

## Purpose

This SOLIDWORKS VBA macro automates the drawing-delivery step for an assembly. It finds the drawing for the active top-level assembly and matching drawings for physical part and subassembly files, then exports each drawing to PDF.

The expected process benefit is less repetitive manual opening and exporting of drawings. No time-savings or error-reduction measurements are available, so no numerical impact is claimed.

## Inputs and outputs

- **Input:** a saved, active `.SLDASM` assembly and available referenced part/subassembly files.
- **Drawing lookup:** each model must have a `.SLDDRW` with the same base filename in the same folder. The top-level assembly drawing is included when present.
- **Output:** one `.pdf` alongside each matching drawing. `SaveAs3` is called with `ExportData=Nothing`, which the project handoff documents as exporting all sheets.
- Repeated component references and duplicate drawing paths are consolidated.

## Workflow

1. Validate SOLIDWORKS, the active document type, and that the assembly has been saved.
2. Add the main assembly and enumerate the assembly components once.
3. Skip suppressed components and descendants of suppressed parents, virtual components, components without paths, and unsupported model extensions.
4. Find same-name drawings next to each physical part/subassembly and consolidate duplicates.
5. Show a confirmation summary, including how many existing PDFs would be overwritten. The default response is No.
6. Open each drawing silently, confirm its document type, activate it without a forced rebuild, and verify that the active document path matches the expected drawing.
7. Call `IModelDocExtension.SaveAs3` and count an export as successful only if the API returns success and the PDF exists with a non-zero size.
8. Close the drawing and verify its closed state, then reactivate the main assembly and report totals in the VBA Immediate Window and a completion dialog.

## Reliability and limitations

- Diagnostics include a version marker, discovered and missing drawings, skipped components, API return values, PDF file checks, close errors, and a final summary.
- The macro does not save native SOLIDWORKS drawings or the assembly.
- **Important data-loss warning:** every processed drawing is closed without saving, including drawings that were already open before the run. Unsaved changes in those drawings can be discarded. Save or otherwise protect drawing edits before running.
- **Potentially stale output:** the macro intentionally does not force a rebuild before export; it uses the drawing's loaded/saved state. A drawing that is out of date may produce an out-of-date PDF.
- Existing PDFs are overwritten after the user confirms. Review the confirmation and preserve outputs as needed.
- Missing matching drawings are counted and skipped; they do not stop the other exports.
- The macro does not apply a filename prefix filter. It considers the main assembly and physical part/subassembly references whose paths resolve to supported SOLIDWORKS model extensions.

## Run and debug

1. Use a copy of the assembly and review all open drawings for unsaved changes.
2. Confirm referenced models and same-name drawings are available in their expected folders.
3. Open the intended SOLIDWORKS VBA project. If updating a module, copy and paste the complete module into the correct standard module; do not use **Import File** in this workspace because it has caused character corruption.
4. Compile in the VBA editor and confirm marker `20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO`.
5. Open the Immediate Window (`Ctrl+G` in the VBA editor), run `main`, review the overwrite confirmation, and inspect `[VERSION]`, `[ACTIVATE DOC3]`, `[CHAMADA PDF]` / `[PDF EXPORT CALL]`, `[API]`, `[PDF OK]`, close diagnostics, and `[RESUMO]` / `[SUMMARY]`.
6. Inspect the PDFs, including sheet count and drawing revision, before delivery.

### Debugging prompt

> Diagnose the SOLIDWORKS VBA assembly-to-PDF workflow using the exact macro marker, SOLIDWORKS version, anonymized assembly/drawing layout, and complete Immediate Window output. Identify the earliest failing step and distinguish open/activation errors, `SaveAs3` return/errors/warnings, and PDF file validation. Do not assume a drawing was rebuilt. Do not change the PDF overwrite policy or close-without-save behavior without explicit approval.

## Validation and process impact

- **Current source status:** static review only. The latest source marker has not been independently compiled or run in SOLIDWORKS in this task.
- **Historical diagnostic evidence:** the handoff records an earlier report of 44 drawings found, zero PDFs, and 44 failures; the supplied log did not contain enough `SaveAs3` output to prove the exact cause. The source was subsequently changed to activate and verify each drawing before export and to log the direct API call and result. This history is not proof that the current marker succeeds.
- The handoff also records user feedback about using the macro with the assembly as the only open document and previously saved drawings; it does not replace a current test of this source marker.
- **Impact measurement pending:** compare manual and automated elapsed time for equivalent assemblies and drawing counts. Record sample size and period, then compare missing/failed exports and rework. Do not claim a percentage or time saved without those measurements.

## Package files

- [AssemblyDrawingPdfExporter.bas](AssemblyDrawingPdfExporter.bas) — English comments, diagnostics, messages, and internal identifiers.
- [AssemblyDrawingPdfExporter.pt-BR.bas](AssemblyDrawingPdfExporter.pt-BR.bas) — Portuguese source copy, preserved from the workspace source.
- [AssemblyDrawingPdfExporter.swp](AssemblyDrawingPdfExporter.swp) — copy of the same-base `.swp` found in the workspace. Its synchronization with either `.bas` copy has **not** been verified.
- [CONTINUACAO.md](CONTINUACAO.md) — short handoff for future work.

The workspace source `_8.salva_PDF_montagem_inteira.bas` was not modified or moved. The English copy's executable code was compared with the source after normalizing comments, diagnostic text, and internal identifier translations; the comparison passed. This static check is not VBA compilation or SOLIDWORKS execution.

---

# Versao em portugues

## Objetivo

Esta macro VBA para SOLIDWORKS automatiza a etapa de entrega de desenhos de uma montagem. Localiza o desenho da montagem principal e os desenhos correspondentes aos arquivos fisicos de pecas e submontagens, exportando cada desenho para PDF.

O beneficio esperado e reduzir a abertura e exportacao manual repetitiva. Nao existem medicoes de tempo economizado ou reducao de erros; portanto, nao ha alegacao numerica de impacto.

## Entradas e saidas

- **Entrada:** montagem `.SLDASM` ativa e salva, com arquivos referenciados de pecas/submontagens disponiveis.
- **Busca do desenho:** cada modelo deve ter um `.SLDDRW` com o mesmo nome-base e na mesma pasta. O desenho da montagem principal tambem e incluido quando existe.
- **Saida:** um `.pdf` ao lado de cada desenho correspondente. `SaveAs3` e chamado com `ExportData=Nothing`, que o handoff do projeto documenta como exportacao de todas as folhas.
- Referencias repetidas a componentes e caminhos duplicados de desenhos sao consolidados.

## Fluxo

1. Valida a disponibilidade do SOLIDWORKS, o tipo do documento ativo e se a montagem foi salva.
2. Adiciona a montagem principal e enumera uma vez os componentes da montagem.
3. Ignora componentes suprimidos e descendentes de pais suprimidos, componentes virtuais, componentes sem caminho e extensoes de modelo nao suportadas.
4. Procura desenhos de mesmo nome-base ao lado das pecas/submontagens fisicas e consolida duplicatas.
5. Exibe uma confirmacao com a quantidade de PDFs existentes que seriam substituidos. A resposta padrao e Nao.
6. Abre cada desenho silenciosamente, confirma o tipo do documento, ativa sem rebuild forcado e verifica se o caminho do documento ativo e o esperado.
7. Chama `IModelDocExtension.SaveAs3` e so considera sucesso se a API retornar sucesso e o PDF existir com tamanho maior que zero.
8. Fecha o desenho, verifica o fechamento, reativa a montagem principal e informa os totais na Immediate Window e em uma mensagem final.

## Confiabilidade e limites

- Os diagnosticos incluem marcador de versao, desenhos encontrados/ausentes, componentes ignorados, retornos da API, verificacao dos PDFs, erros de fechamento e resumo final.
- A macro nao salva desenhos SOLIDWORKS nativos nem a montagem.
- **Aviso importante de perda de dados:** cada desenho processado e fechado sem salvar, inclusive se ja estava aberto antes da execucao. Alteracoes nao salvas nesses desenhos podem ser descartadas. Salve ou proteja as alteracoes antes de executar.
- **Possibilidade de saida desatualizada:** a macro nao forca rebuild; usa o estado carregado/salvo do desenho. Um desenho desatualizado pode gerar um PDF desatualizado.
- PDFs existentes sao substituidos depois da confirmacao do usuario. Confira a confirmacao e preserve as saidas necessarias.
- Desenhos correspondentes ausentes sao contabilizados e ignorados; nao interrompem as demais exportacoes.
- A macro nao aplica filtro por prefixo de nome. Considera a montagem principal e referencias fisicas de pecas/submontagens com extensoes de modelo SOLIDWORKS suportadas.

## Execucao e depuracao

1. Use uma copia da montagem e verifique se ha desenhos abertos com alteracoes nao salvas.
2. Confira se os modelos referenciados e desenhos com mesmo nome-base estao disponiveis nas pastas esperadas.
3. Abra o projeto VBA pretendido. Para atualizar um modulo, copie e cole o modulo completo no modulo padrao correto; nao use **Import File** neste workspace, pois ja causou corrupcao de caracteres.
4. Compile no editor VBA e confirme o marcador `20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO`.
5. Abra a Immediate Window (`Ctrl+G` no editor VBA), execute `main`, revise a confirmacao de substituicao e confira `[VERSION]`, `[ACTIVATE DOC3]`, `[CHAMADA PDF]`, `[API]`, `[PDF OK]`, diagnosticos de fechamento e `[RESUMO]`.
6. Antes da entrega, confira os PDFs, incluindo quantidade de folhas e revisao dos desenhos.

### Prompt de depuracao

> Diagnostique o fluxo VBA de montagem para PDF do SOLIDWORKS usando o marcador exato da macro, versao do SOLIDWORKS, estrutura anonimizada da montagem/pastas de desenhos e a saida completa da Immediate Window. Identifique a primeira etapa com falha e diferencie erros de abertura/ativacao, retorno/erros/avisos de `SaveAs3` e validacao do arquivo PDF. Nao presuma que houve rebuild. Nao altere a politica de substituicao de PDFs nem o fechamento sem salvar sem aprovacao explicita.

## Validacao e impacto no processo

- **Estado da fonte atual:** somente revisao estatica. O marcador da fonte mais recente nao foi compilado nem executado independentemente no SOLIDWORKS nesta tarefa.
- **Evidencia historica de diagnostico:** o handoff registra um relato anterior de 44 desenhos encontrados, zero PDFs e 44 falhas; o log fornecido nao continha dados suficientes de `SaveAs3` para provar a causa exata. Depois, a fonte passou a ativar e validar cada desenho antes da exportacao e registrar a chamada direta da API e seu retorno. Esse historico nao prova sucesso do marcador atual.
- O handoff tambem registra feedback do usuario sobre executar com somente a montagem aberta e desenhos previamente salvos; isso nao substitui um teste atual deste marcador.
- **Medicao de impacto pendente:** compare o tempo manual e automatizado para montagens e quantidades de desenhos equivalentes. Registre amostra e periodo e compare falhas de exportacao e retrabalho. Nao alegue percentual ou economia sem essas medicoes.

## Arquivos do pacote

- [AssemblyDrawingPdfExporter.bas](AssemblyDrawingPdfExporter.bas) — comentarios, diagnosticos, mensagens e identificadores internos em ingles.
- [AssemblyDrawingPdfExporter.pt-BR.bas](AssemblyDrawingPdfExporter.pt-BR.bas) — copia da fonte em portugues, preservada a partir da fonte do workspace.
- [AssemblyDrawingPdfExporter.swp](AssemblyDrawingPdfExporter.swp) — copia do `.swp` de mesmo nome-base encontrado no workspace. A sincronizacao com qualquer uma das copias `.bas` **nao foi verificada**.
- [CONTINUACAO.md](CONTINUACAO.md) — resumo curto para continuar o trabalho.

A fonte do workspace `_8.salva_PDF_montagem_inteira.bas` nao foi alterada nem movida. O codigo executavel da copia inglesa foi comparado com a fonte apos normalizar comentarios, textos de diagnostico e traducao dos identificadores internos; a comparacao passou. Essa verificacao estatica nao e compilacao VBA nem execucao no SOLIDWORKS.
