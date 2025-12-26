# dio-lab-app-excel
Organizador  de Declaração de Imposto de Renda


# Sistema Completo em Microsoft Excel para Cálculo de Imposto de Renda (PF, MEI, PJ)

---

## Estrutura Geral da Planilha Excel

A estrutura do sistema é baseada em múltiplas abas, cada uma dedicada a um tipo de contribuinte: Pessoa Física, MEI e PJ. Cada aba é organizada para permitir o registro, controle e cálculo automático dos dados fiscais relevantes, como CPF/CNPJ, notas fiscais, receitas, despesas dedutíveis e imposto devido. Além disso, há abas auxiliares para listas de validação, tabelas de alíquotas, dashboards e instruções de uso.

A seguir, uma visão geral das abas principais:

| Aba                | Finalidade Principal                                      | Público-Alvo         |
|--------------------|----------------------------------------------------------|----------------------|
| Capa/Introdução    | Página inicial interativa, navegação e instruções         | Todos                |
| Pessoa Física      | Registro de dados, receitas, despesas, cálculo IRPF       | PF                   |
| MEI                | Controle de faturamento, cálculo de rendimentos tributáveis| MEI                  |
| Pessoa Jurídica    | Controle de receitas, despesas, cálculo Simples/Lucro     | PJ                   |
| Notas Fiscais      | Registro e controle de NF-e/NFS-e                        | MEI, PJ              |
| Tabelas de Alíquotas| Regras fiscais atualizadas, faixas, deduções             | Todos                |
| Dashboards         | Resumo visual, gráficos, tabelas dinâmicas                | Todos                |
| Listas e Validações| Categorias, subcategorias, menus suspensos                | Todos                |
| Instruções/Manual  | Explicação detalhada do funcionamento                     | Todos                |

A separação por tipo de contribuinte permite que cada usuário acesse apenas as funcionalidades relevantes ao seu perfil, tornando o preenchimento mais intuitivo e seguro.

---

## Regras Vigentes da Receita Federal para Cálculo do Imposto de Renda (2025-2026)

### Pessoa Física (IRPF)

As regras para o cálculo do Imposto de Renda Pessoa Física (IRPF) em 2025-2026 seguem as faixas e deduções estabelecidas pela Receita Federal. A tabela anual vigente para o exercício de 2026 (ano-calendário 2025) é:

| Base de Cálculo (R$)         | Alíquota (%) | Dedução (R$)   |
|------------------------------|--------------|----------------|
| Até 28.467,20                | 0            | 0              |
| 28.467,21 até 33.919,80      | 7,5          | 2.135,04       |
| 33.919,81 até 45.012,60      | 15,0         | 4.679,03       |
| 45.012,61 até 55.976,16      | 22,5         | 8.054,97       |
| Acima de 55.976,16           | 27,5         | 10.853,78      |

Outros limites importantes:

- Dedução anual por dependente: **R$ 2.275,08**
- Limite anual de despesa com instrução: **R$ 3.561,50**
- Limite anual de desconto simplificado: **R$ 16.754,34**
- Rendimentos previdenciários isentos para maiores de 65 anos: **R$ 1.903,98/mês**

### MEI (Microempreendedor Individual)

O MEI deve declarar IRPF se seus rendimentos tributáveis superarem **R$ 33.888,00** no ano anterior. O cálculo dos rendimentos tributáveis considera:

- **Comércio, Indústria, Transporte de Cargas:** 8% do faturamento é isento, 92% é tributável.
- **Transporte de Passageiros:** 16% isento, 84% tributável.
- **Prestação de Serviços:** 32% isento, 68% tributável.

Além disso, o MEI é obrigado a entregar a **Declaração Anual do Simples Nacional (DASN-SIMEI)** até 31 de maio, independentemente do faturamento.

### Pessoa Jurídica (PJ)

As regras para PJ variam conforme o regime tributário:

- **Simples Nacional:** Alíquotas variam de 4% a 33%, conforme o anexo e faixa de faturamento. O cálculo efetivo considera a receita bruta acumulada, a alíquota da faixa e a parcela a deduzir.
- **Lucro Presumido:** Percentual de presunção sobre a receita, com aplicação de alíquotas específicas para IRPJ e CSLL.
- **Lucro Real:** Apuração do lucro contábil ajustado, dedução de despesas e cálculo do imposto devido.

O limite de faturamento para o Simples Nacional é de **R$ 4,8 milhões/ano**.

---

## Campos Obrigatórios e Validações de Dados

A precisão dos dados fiscais depende da correta validação dos campos obrigatórios. O sistema Excel deve implementar validações automáticas para:

- **CPF:** 11 dígitos, validação de dígito verificador (fórmula ou VBA)
- **CNPJ:** 14 dígitos, validação de dígito verificador (fórmula ou VBA)
- **Tipo de Contribuinte:** Menu suspenso (PF, MEI, PJ)
- **Notas Fiscais:** Número, data, valor, tipo (NF-e, NFS-e), CFOP
- **Receitas:** Fonte, categoria, valor, data
- **Despesas Dedutíveis:** Categoria, valor, comprovante
- **Dependentes:** Nome, CPF, grau de parentesco

A validação pode ser feita por meio de **Validação de Dados** (menu suspenso, listas dependentes com INDIRETO), fórmulas personalizadas (SE, E, OU), e funções LAMBDA para CPF/CNPJ.

---

## Registro e Controle de Notas Fiscais Emitidas

O controle de notas fiscais é essencial para MEI e PJ, e também pode ser útil para PF que recebem rendimentos de aluguel ou serviços. A aba de Notas Fiscais deve permitir:

- Registro de NF-e (produtos) e NFS-e (serviços)
- Identificação do tipo de nota, número, data de emissão, valor, cliente/fornecedor, CFOP
- Status de pagamento (pago, pendente)
- Dashboard com resumo mensal de notas emitidas e recebidas

Exemplo de tabela para controle de notas fiscais:

| Nº NF | Data Emissão | Tipo | Valor | Cliente/Fornecedor | CFOP | Status |
|-------|--------------|------|-------|--------------------|------|--------|
| 12345 | 10/03/2025   | NF-e | 2.500 | ABC Ltda           | 5102 | Pago   |
| 12346 | 12/03/2025   | NFS-e| 1.200 | João Silva         | 5933 | Pendente|

A planilha pode usar **PROCV/XLOOKUP** para buscar dados do cliente/fornecedor, e **SOMASES** para somar valores por período ou categoria.

---

## Controle de Receitas por Fonte e Classificação

A aba de receitas deve permitir o registro detalhado das entradas, com classificação por fonte (salários, serviços, vendas, aluguel, investimentos) e categoria. Recomenda-se o uso de listas suspensas para facilitar o preenchimento e evitar erros de digitação.

Exemplo de campos:

- Data
- Fonte (menu suspenso: salário, serviço, venda, aluguel, investimento)
- Categoria (menu dependente: ex. serviço → consultoria, manutenção, etc.)
- Valor
- Observações

A função **SOMASES** pode ser utilizada para calcular o total de receitas por fonte, categoria ou período. Para dashboards, tabelas dinâmicas permitem visualizar receitas por mês, por fonte e por categoria.

---

## Controle de Despesas Dedutíveis e Comprovantes

O correto registro das despesas dedutíveis é fundamental para reduzir a base de cálculo do imposto. A aba de despesas deve incluir:

- Data
- Categoria (saúde, educação, previdência, pensão alimentícia, dependentes, doações, imóveis)
- Valor
- Comprovante (nº, tipo, instituição)
- Observações

As principais despesas dedutíveis para IRPF em 2025 são:

| Categoria         | Limite de Dedução         | Observações                           |
|-------------------|--------------------------|---------------------------------------|
| Saúde             | Sem limite               | Consultas, exames, internações        |
| Educação          | R$ 3.561,50 por pessoa   | Mensalidades, cursos reconhecidos     |
| Previdência Privada| Até 12% da renda bruta   | Apenas PGBL                           |
| Dependentes       | R$ 2.275,08 por dependente| Filhos, cônjuges, pais, etc.          |
| Pensão alimentícia| Integral                 | Judicialmente determinada             |
| Doações           | Conforme legislação      | Fundos controlados, projetos sociais  |

A planilha deve alertar o usuário caso o valor ultrapasse o limite permitido, usando **validação de dados** e fórmulas **SE**.

---

## Cálculo Automático do Imposto Devido para Pessoa Física

O cálculo do IRPF é feito automaticamente a partir das receitas e despesas registradas. O processo envolve:

1. **Cálculo da Base de Cálculo:**  
   `Base = Rendimentos Tributáveis - Deduções Legais`

2. **Identificação da Faixa de Alíquota:**  
   Utilizar **PROCV/XLOOKUP** para buscar a alíquota e dedução correspondente na tabela da Receita Federal.

3. **Cálculo do Imposto:**  
   `Imposto = (Base x Alíquota) - Dedução`

4. **Aplicação de descontos por dependente, instrução, previdência, etc.**

Exemplo de fórmula aplicada (usando SE e PROCV):

```excel
=SE(Base<=28467,20;0;SE(Base<=33919,80;(Base*7,5%)-2135,04;SE(Base<=45012,60;(Base*15%)-4679,03;SE(Base<=55976,16;(Base*22,5%)-8054,97;(Base*27,5%)-10853,78))))
```

Para facilitar, pode-se usar **SOMARPRODUTO** para calcular o imposto progressivo, especialmente em casos de rendimentos variáveis.

---

## Cálculo Automático do Imposto para MEI

O MEI tem regras específicas:

- **Obrigatoriedade de IRPF:** Se os rendimentos tributáveis (após deduções e aplicação do percentual isento) superarem R$ 33.888,00, o MEI deve declarar IRPF.
- **Cálculo dos rendimentos tributáveis:**  
  - Comércio/Indústria/Transporte de Cargas: 8% isento, 92% tributável
  - Transporte de Passageiros: 16% isento, 84% tributável
  - Serviços: 32% isento, 68% tributável

Exemplo de cálculo para MEI prestador de serviços:

```excel
Rendimento Tributável = Faturamento Bruto - Despesas - (Faturamento Bruto x 32%)
```

A planilha deve orientar o usuário sobre onde informar cada valor na DIRPF (Rendimentos Isentos e Não Tributáveis, Rendimentos Tributáveis Recebidos de PJ pelo Titular).

Além disso, o MEI deve pagar mensalmente o **DAS** (Documento de Arrecadação do Simples Nacional), que inclui contribuição previdenciária e impostos da empresa. O valor é fixo e não depende do faturamento, mas deve ser registrado para controle.

---

## Cálculo Automático do Imposto para PJ (Simples, Lucro Presumido, Lucro Real)

### Simples Nacional

O cálculo do imposto no Simples Nacional envolve:

- Identificação do anexo conforme CNAE (atividade econômica)
- Cálculo da receita bruta acumulada nos últimos 12 meses
- Aplicação da alíquota da faixa correspondente
- Dedução da parcela a deduzir

Fórmula para alíquota efetiva:

```excel
Alíquota Efetiva = ((Receita Bruta x Alíquota) - Parcela a Deduzir) / Receita Bruta
```

Exemplo de tabela para Anexo I (Comércio):

| Faixa | Receita Bruta (12 meses) | Alíquota | Parcela a Deduzir |
|-------|--------------------------|----------|-------------------|
| 1     | Até R$ 180.000,00        | 4%       | 0                 |
| 2     | R$ 180.000,01 a 360.000  | 7,3%     | R$ 5.940,00       |
| ...   | ...                      | ...      | ...               |

A planilha deve usar **PROCV/XLOOKUP** para buscar a alíquota e parcela a deduzir conforme o faturamento e CNAE.

### Lucro Presumido

- Percentual de presunção sobre a receita (ex: 8% para comércio, 32% para serviços)
- Aplicação das alíquotas de IRPJ (15% + adicional) e CSLL (9%)
- Registro de receitas e despesas para cálculo do lucro presumido

### Lucro Real

- Apuração do lucro contábil ajustado
- Dedução de despesas operacionais, financeiras, etc.
- Cálculo do IRPJ e CSLL sobre o lucro real

A aba de PJ deve permitir a escolha do regime tributário via menu suspenso, com fórmulas adaptadas para cada caso.

---

## Uso de Fórmulas Avançadas: XLOOKUP/PROCV, SE, SOMASES, SOMARPRODUTO

O sistema Excel utiliza diversas fórmulas avançadas para automação dos cálculos e validação dos dados:

- **PROCV/XLOOKUP:** Pesquisa de alíquotas, deduções, dados de clientes/fornecedores, categorias
- **SE:** Teste de condições (ex: faixa de renda, limite de dedução)
- **SOMASES:** Soma de receitas/despesas por categoria, período, fonte
- **SOMARPRODUTO:** Cálculo progressivo do imposto, cruzamento de múltiplos critérios
- **INDIRETO:** Listas suspensas dependentes (ex: categoria/subcategoria)
- **TABELAS DINÂMICAS:** Resumo de receitas, despesas, impostos, dashboards interativos

Exemplo prático de uso do XLOOKUP:

```excel
=XLOOKUP(Base;TabelaFaixas[Base];TabelaFaixas[Alíquota];"Não encontrado";0)
```

O XLOOKUP permite buscas flexíveis, inclusive para múltiplos critérios, pesquisa reversa, personalização de mensagens de erro e retorno de múltiplos valores.

---

## Tabelas Dinâmicas e Dashboards para Resumo e Análise Fiscal

As tabelas dinâmicas são fundamentais para análise fiscal, permitindo:

- Resumo de receitas e despesas por mês, categoria, fonte
- Comparativo de impostos pagos por tipo de contribuinte
- Identificação de tendências e padrões de consumo
- Visualização gráfica (colunas, pizza, linhas) dos dados fiscais

O dashboard pode incluir:

- Receita total por mês/ano
- Despesa total por categoria
- Imposto devido por tipo de contribuinte
- Gráficos de evolução do saldo fiscal
- Segmentação de dados para filtros dinâmicos

A atualização automática das tabelas dinâmicas pode ser feita via VBA, garantindo que os dados estejam sempre atualizados ao abrir o arquivo ou modificar os dados de origem.

---

## Validações e Listas Dependentes (INDIRETO) para Categorias e Subcategorias

A criação de listas suspensas dependentes melhora a experiência do usuário e reduz erros de preenchimento. Por exemplo, ao selecionar "Serviço" como categoria de receita, a lista de subcategorias exibe apenas opções relevantes (consultoria, manutenção, etc.).

O método envolve:

- Criação de intervalos nomeados para cada categoria
- Uso da função **INDIRETO** na validação de dados para conectar a lista principal à dependente

Exemplo de fórmula:

```excel
=INDIRETO(A2)
```

Onde A2 contém a categoria selecionada, e o intervalo nomeado corresponde ao texto da célula.

---

## Navegação Facilitada: Botões, Links Internos e Macros Simples (VBA)

A navegação entre abas é facilitada por:

- **Botões de comando:** Criados via menu Desenvolvedor, direcionam o usuário para abas específicas
- **Hiperlinks internos:** Permitem acesso rápido a seções relevantes
- **Macros VBA:** Automatizam abertura da página inicial, ajuste de zoom, ocultação de menus, atualização de tabelas dinâmicas

Exemplo de macro para ativar a aba "Capa" ao abrir o arquivo:

```vba
Private Sub Workbook_Open()
    Sheets("Capa").Activate
    ActiveWindow.Zoom = 100
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
End Sub
```

Esses recursos tornam o sistema mais amigável, profissional e seguro para o usuário final.

---

## Layout Profissional e Formatação (Cores Neutras, Tipografia, Proteção)

O layout da planilha é projetado para ser visualmente agradável e funcional:

- **Cores neutras e leves:** Tons de cinza, azul claro, branco, para facilitar a leitura e evitar fadiga visual
- **Tipografia clara:** Fontes padrão do Excel (Calibri, Arial), tamanhos adequados para títulos e dados
- **Formatação condicional:** Destaque automático de campos obrigatórios, erros de preenchimento, limites de dedução
- **Bordas automáticas:** Separação de seções, tabelas e campos
- **Proteção de células e abas:** Bloqueio de fórmulas e áreas críticas, proteção por senha para dados sensíveis

A personalização pode incluir logotipo, instruções rápidas, campos de contato para suporte, e ocultação de menus para evitar edições acidentais.

---

## Exemplos Práticos de Preenchimento e Fórmulas Aplicadas

### Caso 1: Pessoa Física

- Receita anual: R$ 50.000,00
- Despesas dedutíveis: Saúde (R$ 5.000,00), Educação (R$ 3.000,00), Previdência (R$ 6.000,00)
- Dependentes: 2

Preenchimento:

- Aba PF: Inserir receitas, despesas, dependentes
- Fórmula de cálculo do imposto:

```excel
Base = 50.000 - (5.000 + 3.000 + 6.000 + 2*2.275,08)
Imposto = (Base x Alíquota) - Dedução
```

### Caso 2: MEI Prestador de Serviços

- Faturamento anual: R$ 72.000,00
- Despesas: R$ 15.000,00

Cálculo:

- Parcela isenta: 32% x 72.000 = R$ 23.040,00
- Parcela tributável: (72.000 - 15.000) - 23.040 = R$ 33.960,00

Como o resultado é maior que R$ 33.888,00, o MEI deve declarar IRPF.

### Caso 3: PJ Simples Nacional (Comércio)

- Receita bruta acumulada: R$ 400.000,00
- Anexo I, faixa 2: Alíquota 7,3%, parcela a deduzir R$ 5.940,00

Cálculo:

```excel
Alíquota Efetiva = ((400.000 x 7,3%) - 5.940) / 400.000 = 6,13%
Imposto devido = 400.000 x 6,13% = R$ 24.520,00
```

Esses exemplos são acompanhados de instruções passo a passo e fórmulas aplicadas nas células correspondentes.

---

## Exportação e Integração com Sistemas Contábeis e Receita Federal (SPED, e-CAC)

A planilha pode ser integrada com sistemas contábeis e plataformas da Receita Federal:

- **Exportação para TXT/CSV:** Facilita a importação dos dados para sistemas como SPED, e-CAC
- **Importação de arquivos fiscais:** Ferramentas como TaxSheets permitem editar arquivos do SPED diretamente no Excel, corrigir erros e gerar arquivos prontos para entrega ao fisco
- **Automação de relatórios:** Geração de relatórios em PDF, gráficos para apresentação

Esses recursos aumentam a produtividade, reduzem erros e garantem conformidade com as exigências legais.

---

## Documentação e Instruções de Uso

A aba de instruções/Manual inclui:

- Guia passo a passo para preenchimento de cada aba
- Explicação das fórmulas utilizadas
- Orientações sobre limites de dedução, obrigatoriedade de declaração
- Dicas de validação, proteção e atualização de regras fiscais
- FAQ com dúvidas frequentes

A documentação é clara, objetiva e adaptada para usuários com diferentes níveis de conhecimento em Excel.

---

## Segurança e Privacidade dos Dados Fiscais

A proteção dos dados fiscais é garantida por:

- **Criptografia de arquivos:** Proteção por senha ao abrir o arquivo
- **Proteção de células e abas:** Bloqueio de áreas críticas, restrição de edição
- **Controle de acesso:** Definição de permissões para usuários diferentes
- **Backup automático:** Recomendações para cópias de segurança

Essas medidas evitam acessos não autorizados, alterações acidentais e garantem a confidencialidade das informações.

---

## Atualizações Legais e Manutenção da Planilha

A manutenção da planilha envolve:

- Atualização anual das tabelas de alíquotas, limites e deduções conforme legislação vigente
- Inclusão de novas funcionalidades conforme mudanças fiscais
- Suporte técnico para dúvidas e correções
- Compatibilidade com versões recentes do Excel e outros softwares de planilhas

O usuário é orientado sobre como atualizar as regras e fórmulas, garantindo que o sistema permaneça aderente às normas da Receita Federal.

---

## Recursos e Modelos Existentes

Diversos modelos de planilhas públicas, templates pagos e projetos open source estão disponíveis para download e personalização:

- **ePlanilhas:** Pacote com milhares de planilhas prontas, incluindo modelos de IRPF, carnê-leão, controle financeiro, dashboards
- **GitHub:** Projetos como Organizador-IR, Planilha_IRPF, Gerenciador de Imposto de Renda, com estrutura lógica, validações automáticas e navegação intuitiva
- **TaxSheets:** Ferramenta para edição de arquivos SPED no Excel, integração com sistemas fiscais
- **Planilhas.VC, SmartPlanilhas:** Modelos para controle de notas fiscais, dashboard de NF, relatórios de despesas

Esses recursos podem ser usados como ponto de partida para personalização e adaptação às necessidades específicas do usuário.

---

**Este relatório serve como guia completo para a implementação, uso e manutenção de um sistema Excel para cálculo de Imposto de Renda, pronto para download e adaptação conforme as necessidades do contribuinte brasileiro.**
Perfeito, Gabriel! Vou começar a desenvolver um sistema completo no Excel para cálculo de Imposto de Renda voltado para Pessoa Física, MEI e Pessoa Jurídica. Isso incluirá:

- Estrutura de planilhas separadas por tipo de contribuinte
- Campos para CPF, CNPJ, notas fiscais, receitas, despesas e deduções
- Cálculo automático do imposto devido com base nas regras da Receita Federal
- Validações de dados, menus suspensos e fórmulas como `SE`, `SOMASES`, `XLOOKUP`, `PROCV`, entre outras
- Layout profissional com cores neutras e leves
- Exemplos práticos de preenchimento e explicações de cada fórmula
