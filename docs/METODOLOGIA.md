# Metodologia e limites

## População e rastreabilidade

- Uma linha preenchida com dados de identificação/financeiros é um registro, normalmente por nota e natureza. Linhas com apenas fórmulas de atraso são ignoradas.
- O recorte usa **DATA PAGTO.** entre `inicio` e `fim`, inclusive. O mês é derivado dessa data. Mês declarado divergente gera alerta. Pagamento sem data válida ou fora do período é listado em excluídos.
- Processo único é o SEI após remoção de espaços nas extremidades. Não se removem registros por repetirem o SEI. Possíveis duplicatas de processo + nota + natureza são sinalizadas e mantidas, sem deduplicação destrutiva.
- Identificadores são texto. Nomes/secretarias têm caixa e espaços normalizados; variantes ortográficas não são fundidas automaticamente.
- O número da linha original acompanha a base e os casos. A fonte e os documentos de referência têm hashes SHA-256 no manifesto.
- Colunas OP, empenho, observação e auxiliares não participam da análise. Não se inferem motivos de ressalva ou falhas documentais sem os processos originais.

## Valores

Valores são convertidos em centavos inteiros com arredondamento decimal para evitar resíduos binários. Total concedido significa a coluna `valor adt` da planilha, sem comprovação dos pedidos nem das transferências. Valores ausentes não viram zero. Os totais somam valores disponíveis e mostram a quantidade de preenchidos.

A diferença por registro é adiantado menos utilizado e devolvido. Diferenças não são automaticamente corrigidas ou interpretadas como prejuízo. O consolidado é comparado célula a célula nos meses e no total do período, incluindo a linha de total. Seus valores são caches armazenados no arquivo, não uma recálculação do Excel. Células sem cache são identificadas.

Percentual de utilização agregado = soma utilizada / soma adiantada com cobertura completa para ambas as medidas. Não se usa média simples de percentuais. Médias de valor adiantado, mediana, mínimo, máximo e percentil 90 usam valores válidos por registro. P90 usa o método da posição mais próxima: posição `ceil(0,9*n)` na série ordenada.

Cada grupo mostra registros e processos únicos. Somar processos únicos dos grupos pode contar novamente processos que aparecem em vários grupos. Rankings representam a base disponibilizada; não comprovam a cobertura de todos os adiantamentos da Prefeitura.

## Datas, médias e atrasos

`data_observacao` explicita a situação disponível na fonte. O recorte financeiro é o do pagamento; aplicações, devoluções e prestações posteriores ao fim do semestre são mantidas até a observação. A planilha não contém a data do parecer SMF/CGM, portanto não é possível reconstruir a situação histórica desses exames em 30 de junho.

Durações = data do evento menos data do pagamento, em dias corridos. Datas anteriores ao pagamento ou posteriores à observação são excluídas das médias e sinalizadas. A aplicação é a **última** utilização, não a primeira; não se mede tempo até início da despesa. A data de pagamento é uma aproximação da disponibilização. A prestação informada na planilha não prova o protocolo na SMF (art. 17, §1).

Devolução é aplicável quando valor devolvido > 0. Aplicação é aplicável quando valor utilizado > 0. Valores iguais a zero são não aplicáveis aos respectivos eventos; ausências não são confundidas com esses casos. Campos contraditórios geram alerta. Não se presume pendência vencida só pela falta de data.

O atraso principal é `max(0, data_evento - prazo_registrado)` para eventos avaliáveis. A taxa de atraso divide atrasados por registros com evento e vencimento válidos, não pela base inteira. Média de atraso considera somente atrasados; média de duração considera todos os eventos elegíveis. Extremos preservam empates e indicam servidor, nota, processo e linha em `extremos_prazos.csv`.

## Comparação técnica dos vencimentos

Fontes locais: IN-001/2018-CGM, versão retificada em 2025, art. 6 (PDF p. 3), art. 13 (p. 5), art. 17 (p. 6) e §3 (p. 7).

- Referência da aplicação: pagamento + 30 dias corridos, excluindo o dia inicial.
- Referência da devolução: três dias úteis após o término da aplicação registrado, ou referência de 30 dias se ele faltar.
- Referência da prestação: prazo de aplicação registrado + 10 dias corridos; se não útil, primeiro útil seguinte.
- Úteis: segunda a sexta, excetuando as datas de `feriados`. O calendário só fica validado quando o operador preenche e verifica feriados nacionais, estaduais, municipais, pontos facultativos e expediente aplicável e marca `calendario_validado: true`.

As datas técnicas ficam em colunas separadas e NÃO substituem os prazos registrados, nem são usadas para alterar o parecer ou a contagem principal de atraso. A ausência de calendário validado é explicitada. A regra de fechamento do exercício (art. 18), cancelamentos (art. 17, §2) e autorizações excepcionais não são aplicadas automaticamente; para análises anuais ou casos específicos, revisar antes de concluir.

## Exames

Categorias normalizadas: Aprovado, Ressalva, Rejeitado, Ausente e Não informado. Valores não reconhecidos ficam em Outro. Aprovação sem ressalva é distinta de ressalva. Os percentuais por órgão usam o total de registros, incluindo ausências. Divergência é diferença literal entre as categorias, inclusive ausência. Não significa necessariamente conflito técnico entre pareceres.

Valores envolvidos em categorias são montantes dos registros, não danos, multas ou ressarcimentos. Não se infere alcance, reincidência impeditiva ou legalidade de novas liberações: faltam cronologia dos pareceres, todo o exercício e processos comprobatórios.

## Atualização e IA

Planilha analítica, gráficos e textos são fotografias da execução. Correções manuais nesses produtos não retroalimentam o motor. Para atualizar os cálculos, alterar a fonte e executar novamente. O manifesto documenta os arquivos antes de edições manuais posteriores.

Os textos principais são modelos preenchidos por dados calculados. O Ollama é opcional e local. Na versão atual não recebe dados reais: sugere somente abertura e encerramento sem números, que são validados quanto a formato e ausência de dígitos/percentuais/valores. O corpo do texto é inserido pelo código. Isso evita a repetição dos erros numéricos encontrados no teste inicial de redação livre. O rascunho fica separado, identificado para revisão, e nunca altera os dados. O resumo agregado em dados/ é material de apoio local. Não há integração automática com GPT, Claude ou Gemini nesta versão.
