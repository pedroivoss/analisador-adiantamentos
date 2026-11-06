# Apresentação dos quadros para o RCI

A versão 0.2 muda a apresentação e preserva a fonte. A planilha analítica contém quadros de até seis colunas e largura total equivalente a 480 pontos, compatível com página em retrato. Os títulos e cabeçalhos não usam nomes técnicos com sublinhado. Meses aparecem como jan/26.

O relatório principal reúne os mesmos quadros da planilha, explicação do que mostram, leitura dos resultados, critérios e gráficos. Há duas versões Word: tabelas nativas editáveis e tabelas em imagem para manter o desenho ao copiar. A pasta Quadros para o RCI contém cada tabela em PNG, com título e fonte.

As tabelas de servidores são divididas em partes de até 18 nomes. Os rankings de dez posições preservam todos os empates no corte, portanto podem ter mais de dez linhas. A grafia dos nomes é preservada da fonte. Grafias diferentes não são unidas por semelhança; isso exigiria confirmação de identidade.

Atrasos principais são contados diretamente dos campos S, T e U da planilha. Os números de dias são extraídos dos textos escritos em Sim. Os demais conteúdos ficam em Outros. Prazos recalculados e comparação de calendário aparecem somente na revisão técnica e não alteram a classificação preenchida.

Médias, medianas, percentuais e tempos são indicadores derivados dos valores e datas. O critério acompanha cada quadro. Valores originais, dados ausentes e pareceres não são inventados, completados ou corrigidos automaticamente. A transcrição sem alterações fica em dados/transcricao_fonte.json com o número da linha original.

Antes da entrega, o gerador confere somas gerais e por secretaria e servidor com a fonte. O verificador confere os mesmos dados em Excel e nas tabelas Word, a largura, os cabeçalhos e a existência dos PNGs. O hash SHA-256 permite detectar alteração posterior da planilha original. Rode src/verificar_apresentacao.py para essa conferência.

Na execução de 6 de novembro de 2026, o Word abriu a versão de tabelas editáveis em 18 páginas, uma por quadro, e permitiu exportar para PDF para conferência visual. Cada nova alteração de conteúdo exige nova verificação de paginação.
