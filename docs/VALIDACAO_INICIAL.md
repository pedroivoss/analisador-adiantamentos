# Validação inicial - 6 de novembro de 2026

A fonte foi relida depois de o operador informar a correção da linha 11. A execução atual tem 139 registros, 106 processos SEI únicos, 69 servidores e 19 secretarias. Os valores são R$ 358.052,24 adiantados, R$ 281.797,29 utilizados e R$ 76.254,95 devolvidos. A diferença financeira é zero. O erro de R$ 3.500,00 identificado na leitura anterior não consta mais do resultado atualizado.

As comparações das células preenchidas no consolidado não apresentaram divergências; três células de totais gerais de período não têm valores armazenados. A base detalhada permite calcular os totais independentemente dessas células.

Restaram 261 alertas, distribuídos em 134 linhas. Alertas não equivalem a irregularidades ou pareceres novos. A maior parte é comparação técnica de vencimentos: aplicação registrada diferente de pagamento +30, devolução em relação a três dias úteis e prestação com prorrogação do vencimento não útil. Feriados/expediente municipal ainda precisam ser validados. Há também datas anteriores ao pagamento, campos contraditórios e uma natureza fora de serviço/consumo.

Foram executados quatro testes: centavos e ausência, calendário com feriado/fim de semana, normalização de exames/datas e integração com dois registros do mesmo SEI, linha de fórmula vazia e devolução anterior ao pagamento. As verificações da execução conferem tipos numéricos no XLSX, ausência de erros de célula, contagens, totais, hashes da fonte e dos arquivos, leitura dos PDFs e estrutura dos DOCX.

Os PDFs foram renderizados com Poppler e verificados visualmente. A ferramenta de renderização DOCX falhou porque o LibreOffice empacotado não está disponível no Windows desta sessão. A tentativa de exportação pelo Word via COM retornou erro de sessão de logon. O acesso pela ferramenta de controle de aplicativos expirou na autorização de acesso ao Word. Portanto os DOCX têm validação estrutural, mas **a paginação dos editáveis precisa ser conferida no Word**. O PDF é uma composição independente do mesmo conteúdo, não uma conversão validada do DOCX.

Ollama respondeu com `qwen3:14b`. A primeira redação livre, mesmo usando dados agregados, incluiu comparações numéricas incorretas; esse modo não foi aprovado. O fluxo foi restringido: o modelo gera somente uma abertura e encerramento sem números, sem receber dados reais. O corpo completo é inserido pelo motor determinístico. O teste dessa versão foi concluído, preservando os números calculados. Não se pode garantir a correção factual de uma redação livre só porque o modelo é local.

As versões intermediárias permanecem separadas. Use a pasta apontada por `ULTIMA_EXECUCAO.txt`. Antes do RCI definitivo, revisar fonte, calendário, classificação não padronizada, datas sinalizadas e correspondência dos exames com os processos. Atualizar os dados e executar novamente após a revisão.
