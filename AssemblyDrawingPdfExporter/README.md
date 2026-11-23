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

- **User-reported test (2026-11-23):** the user confirmed that this macro was tested successfully in SOLIDWORKS 2025. The exact source marker and test scenario were not specified. The assistant did not observe the run.
- **Static review:** the English presentation copy was compared with the workspace source after normalizing comments, diagnostic text, and internal identifier translations; the comparison passed. This check is separate from the user's execution test.
- **Historical diagnostic evidence:** the handoff records an earlier report of 44 drawings found, zero PDFs, and 44 failures; the supplied log did not contain enough `SaveAs3` output to prove the exact cause. The source was subsequently changed to activate and verify each drawing before export and to log the direct API call and result. The user's later SOLIDWORKS 2025 confirmation indicates a successful test, but does not identify which marker or scenario was tested.
- The handoff also records user feedback about using the macro with the assembly as the only open document and previously saved drawings.
- **Impact measurement pending:** compare manual and automated elapsed time for equivalent assemblies and drawing counts. Record sample size and period, then compare missing/failed exports and rework. Do not claim a percentage or time saved without those measurements.

## Package files

- [AssemblyDrawingPdfExporter.bas](AssemblyDrawingPdfExporter.bas) — English comments, diagnostics, messages, and internal identifiers.
- [AssemblyDrawingPdfExporter.pt-BR.bas](AssemblyDrawingPdfExporter.pt-BR.bas) — Portuguese source copy, preserved from the workspace source.
- [AssemblyDrawingPdfExporter.swp](AssemblyDrawingPdfExporter.swp) — copy of the same-base `.swp` found in the workspace. Its synchronization with either `.bas` copy has **not** been verified.
- [CONTINUACAO.md](CONTINUACAO.md) — short handoff for future work.

The workspace source `_8.salva_PDF_montagem_inteira.bas` was not modified or moved. The English copy's executable code was compared with the source after normalizing comments, diagnostic text, and internal identifier translations; the comparison passed. This static check is not VBA compilation or SOLIDWORKS execution.

---

<a id="versao-em-portugues"></a>

# Versão em português

## Objetivo

Esta macro VBA para SOLIDWORKS automatiza a etapa de entrega dos desenhos de uma montagem. Localiza o desenho da montagem principal e os desenhos correspondentes aos arquivos físicos de peças e submontagens, exportando cada desenho para PDF.

O benefício esperado é reduzir a abertura e a exportação manual repetitiva. Não há medições de tempo economizado ou de redução de erros; portanto, não fazemos alegações numéricas de impacto.

## Entradas e saídas

- **Entrada:** montagem `.SLDASM` ativa e salva, com os arquivos referenciados de peças e submontagens disponíveis.
- **Busca do desenho:** cada modelo deve ter um `.SLDDRW` com o mesmo nome-base e na mesma pasta. O desenho da montagem principal também é incluído, quando existe.
- **Saída:** um `.pdf` ao lado de cada desenho correspondente. `SaveAs3` é chamado com `ExportData=Nothing`, que, segundo o handoff do projeto, exporta todas as folhas.
- Referências repetidas a componentes e caminhos duplicados de desenhos são consolidados.

## Fluxo

1. Valida se o SOLIDWORKS está disponível, se o documento ativo é uma montagem e se ela foi salva.
2. Adiciona a montagem principal à lista e enumera os componentes da montagem uma única vez.
3. Ignora componentes suprimidos e descendentes de pais suprimidos, componentes virtuais, componentes sem caminho e extensões de modelo não suportadas.
4. Procura desenhos com o mesmo nome-base ao lado das peças e submontagens físicas e consolida duplicatas.
5. Exibe uma confirmação com a quantidade de PDFs existentes que serão substituídos. A opção padrão é **Não**.
6. Abre cada desenho silenciosamente, confirma o tipo do documento, ativa-o sem forçar rebuild e verifica se o caminho do documento ativo é o esperado.
7. Chama `IModelDocExtension.SaveAs3` e só considera a exportação bem-sucedida se a API retornar sucesso e o PDF existir com tamanho maior que zero.
8. Fecha o desenho, verifica se foi fechado, reativa a montagem principal e informa os totais na Immediate Window e em uma mensagem final.

## Confiabilidade e limites

- Os diagnósticos incluem o marcador de versão, desenhos encontrados ou ausentes, componentes ignorados, retornos da API, validação dos PDFs, erros ao fechar documentos e um resumo final.
- A macro não salva os desenhos SOLIDWORKS nativos nem a montagem.
- **Aviso importante sobre perda de dados:** cada desenho processado é fechado sem salvar, inclusive se já estava aberto antes da execução. Alterações não salvas nesses desenhos podem ser descartadas. Salve ou proteja as alterações antes de executar a macro.
- **Possibilidade de saída desatualizada:** a macro não força rebuild; usa o estado carregado ou salvo do desenho. Se o desenho estiver desatualizado, o PDF também poderá ficar desatualizado.
- PDFs existentes são substituídos após a confirmação do usuário. Confira a confirmação e preserve as saídas necessárias.
- Desenhos correspondentes ausentes são contabilizados e ignorados; isso não interrompe as demais exportações.
- A macro não aplica filtro por prefixo de nome. Considera a montagem principal e as referências físicas de peças e submontagens com extensões de modelo SOLIDWORKS suportadas.

## Execução e depuração

1. Use uma cópia da montagem e verifique se há desenhos abertos com alterações não salvas.
2. Confira se os modelos referenciados e os desenhos de mesmo nome-base estão disponíveis nas pastas esperadas.
3. Abra o projeto VBA desejado. Para atualizar um módulo, copie e cole o módulo completo no módulo padrão correto. Não use **Import File** neste workspace, pois esse procedimento já causou corrupção de caracteres.
4. Compile no editor VBA e confirme o marcador `20261105-04-PDF-MONTAGEM-SAVEAS-DIRETO`.
5. Abra a Immediate Window (`Ctrl+G` no editor VBA), execute `main`, revise a confirmação de substituição e confira `[VERSION]`, `[ACTIVATE DOC3]`, `[CHAMADA PDF]`, `[API]`, `[PDF OK]`, os diagnósticos de fechamento e `[RESUMO]`.
6. Antes da entrega, confira os PDFs, incluindo a quantidade de folhas e a revisão dos desenhos.

### Prompt de depuração

> Diagnostique o fluxo VBA de exportação de montagem para PDF do SOLIDWORKS usando o marcador exato da macro, a versão do SOLIDWORKS, uma descrição anonimizada da montagem e das pastas de desenhos e a saída completa da Immediate Window. Identifique a primeira etapa que falhou e diferencie erros de abertura ou ativação, retorno/erros/avisos de `SaveAs3` e validação do arquivo PDF. Não presuma que houve rebuild. Não altere a política de substituição de PDFs nem o fechamento sem salvar sem aprovação explícita.

## Validação e impacto no processo

- **Teste relatado pelo usuário (2026-11-23):** o usuário confirmou que testou esta macro com sucesso no SOLIDWORKS 2025. O marcador da fonte e o cenário de teste não foram especificados. O assistente não observou a execução.
- **Revisão estática:** a cópia de apresentação em inglês foi comparada com a fonte do workspace após normalizar comentários, diagnósticos e tradução dos identificadores internos; a comparação passou. Essa verificação é independente do teste relatado pelo usuário.
- **Histórico de diagnóstico:** o handoff registra um relato anterior de 44 desenhos encontrados, zero PDFs e 44 falhas; o log fornecido não continha dados suficientes de `SaveAs3` para provar a causa exata. Depois, a fonte passou a ativar e validar cada desenho antes da exportação e a registrar a chamada direta da API e seu retorno. A confirmação posterior do usuário relata um teste bem-sucedido no SOLIDWORKS 2025, mas não identifica o marcador ou cenário testado.
- O handoff também registra o feedback do usuário sobre executar a macro com somente a montagem aberta e os desenhos previamente salvos.
- **Medição de impacto pendente:** compare o tempo manual e automatizado para montagens e quantidades de desenhos equivalentes. Registre o tamanho da amostra e o período; compare também falhas de exportação e retrabalho. Não alegue percentuais ou economia sem essas medições.

## Arquivos deste pacote

- [AssemblyDrawingPdfExporter.bas](AssemblyDrawingPdfExporter.bas) — comentários, diagnósticos, mensagens e identificadores internos em inglês.
- [AssemblyDrawingPdfExporter.pt-BR.bas](AssemblyDrawingPdfExporter.pt-BR.bas) — cópia em português preservada a partir da fonte do workspace.
- [AssemblyDrawingPdfExporter.swp](AssemblyDrawingPdfExporter.swp) — cópia do `.swp` de mesmo nome-base encontrado no workspace. A sincronização com qualquer uma das cópias `.bas` **não foi verificada**.
- [CONTINUACAO.md](CONTINUACAO.md) — resumo curto para dar continuidade ao trabalho.

A fonte do workspace `_8.salva_PDF_montagem_inteira.bas` não foi alterada nem movida. O código executável da cópia em inglês foi comparado com a fonte após normalizar comentários, textos de diagnóstico e tradução dos identificadores internos; a comparação passou. Essa verificação estática não é compilação VBA nem execução no SOLIDWORKS.
