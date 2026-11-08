# Redação analítica para o RCI

A versão 0.3 preserva a apresentação dos quadros e aprofunda a redação. O corpo destinado ao RCI apresenta valores, participações, diferenças entre grupos, leitura dos pareceres, enquadramento normativo e recomendações. Não descreve nomes de colunas, critérios de extração, arquivos ou procedimentos do gerador.

O caderno de conferência reúne as informações operacionais, a cobertura de cada requisito normativo, as pendências e a rastreabilidade das sínteses. Essa separação permite copiar o texto do relatório completo para o RCI sem transportar instruções de uso da planilha.

A análise usa a IN-001/2018-CGM retificada em 2025 e o Manual fornecido. Onde o Manual reproduz regra antiga ou orientação diferente da IN atualizada, não se aplica automaticamente o trecho antigo. Os limites monetários dependem da atualização federal anual e do enquadramento da concessão. O caderno registra a fonte oficial consultada para 2026.

As observações da coluna AB passaram a ser consultadas para contextualização, conforme a solicitação de análise completa. Não se usa a coluna auxiliar AG como fonte de parecer. As nove sínteses das contas rejeitadas foram revisadas e vinculadas a processo, nota e hash do texto. Não se produzem contagens gerais de causas apenas por coincidência de palavras. O registro de revisão local contém dados reais e é excluído do Git; a ausência dele em outra máquina faz o motor omitir as sínteses individuais não revisadas.

As categorias de exame não são alteradas. Ressalva não é confundida com rejeição e montante de conta rejeitada não é apresentado como dano ou glosa definitiva. Os saldos zerados são interpretados com cautela quando há relato de despesa impugnada. A frequência de registros não é usada para concluir automaticamente por impedimento de novas concessões.

O script analisar.ps1 gera a análise completa após os quadros. Para aprofundar uma execução existente, use `python src/aprofundar_relatorio.py --config config.local.json`. Uma nova pasta é criada para preservar versões anteriores. Para conferir a fidelidade das tabelas e a evidência dos casos, execute `python src/verificar_aprofundamento.py`.
