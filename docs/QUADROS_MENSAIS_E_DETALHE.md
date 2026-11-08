# Quadros mensais e relação individual

O fluxo `analisar.ps1` inclui os complementos mensais e individuais após a análise aprofundada. A fonte permanece intacta, inclusive suas cores, fórmulas e valores.

Os seis quadros mensais leem os valores armazenados em `CONSOLIDADO POR SECRETARIA`, mantêm todas as 21 unidades da fonte e confrontam as 378 células financeiras com os registros do pagamento correspondente. Os 18 totais mensais também são conferidos contra a linha 24. A ausência de valor armazenado ou a divergência entre as abas interrompe o complemento; não há substituição silenciosa de um demonstrativo pelo outro.

O quadro de presença considera o valor adiantado maior que zero em cada mês. Meses sem movimento não demonstram inexistência de solicitação: os relatórios usam a expressão “sem adiantamento registrado”. As siglas e os nomes completos do consolidado são mantidos como escritos na fonte.

A relação individual copia os 13 campos solicitados das colunas B, C, D, E, F, G, H, S, T, U, W, Z e AA, na ordem das linhas originais elegíveis. Não normaliza grafias, mês declarado, natureza ou textos de atraso e parecer. Os valores monetários permanecem numéricos no Excel. O relatório em Word e PDF usa formato monetário brasileiro. Linha original e processo SEI permanecem no arquivo de rastreabilidade de apoio.

`Analise mensal e detalhada.xlsx` contém seis abas mensais, presença e relação individual. As abas mensais têm área de impressão de um quadro por página; a relação individual fica em paisagem, com cabeçalho repetido e sem limite vertical de uma página. A configuração de impressão é aplicada em OOXML após a autoria com artifact-tool, pois não há API de impressão documentada no runtime usado.

`Relatorio completo para o RCI.docx` também recebe os seis quadros e o quadro de presença, com comentários. `Relacao individual dos adiantamentos.docx` é um anexo em paisagem com os treze campos simultâneos, comentários e blocos de registros. Os arquivos de imagem por mês e por parte do anexo permitem inserção manual no RCI.

Depois de gerar, executar `python src/verificar_complementos.py` para conferir exaustivamente os 1.807 campos literais dos 139 registros desta execução, valores mensais, totais, presença e tabelas do Word. A quantidade é variável em execuções futuras. Conferir também a paginação no Word antes de publicar. O manifesto deve ser atualizado depois da exportação final do Word para PDF.

Planilha definitiva recebida em 07/11/2026: 139 registros, 106 processos e 62 identificações nominais. As contagens anteriores de 69 identificações não devem ser reaproveitadas. Os totais financeiros permaneceram em R$ 358.052,24 adiantados, R$ 281.797,29 utilizados e R$ 76.254,95 devolvidos.
