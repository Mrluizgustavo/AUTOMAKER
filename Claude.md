# Análise Técnica: AutoMaker

**Visão Geral do Sistema**
O **AutoMaker** é uma aplicação desktop monolítica desenvolvida em Python para automação de back-office, desenhada primariamente para setores de Recursos Humanos e Financeiro. 

**Objetivo Central**
Consolidar, tratar e unificar dados financeiros e operacionais dispersos em múltiplas planilhas (folha de pagamento, impostos, férias, rescisões, vale-transporte e almoxarifado) e automatizar a geração de documentos legais físicos, como formulários de telegrama dos Correios.

**Mecânica de Funcionamento**

| Módulo | Execução Técnica | Output |
| :--- | :--- | :--- |
| **Core UI** | Instanciação do loop principal via `tkinter`. Configuração de frames estáticos de navegação e injeção de callbacks das *views* secundárias através de matriz de dados gerada dinamicamente no `interface_master.py`. | Interface gráfica de controle. |
| **Engenharia de Dados (Despesas)** | Ingestão de múltiplos arquivos Excel via `pandas.read_excel`. Normalização vetorial rigorosa: remoção de whitespaces, conversão de tipagem, extração avançada de datas via *Regex* e padronização de nomenclatura (ex: chaves de "LOJA"). Realização de merges múltiplos (`LEFT JOIN`) e agrupamentos vetoriais (`groupby`). Implementa regras de negócio fixas via código, como abstração de R$ 95.000 em impostos para a unidade central (ADM). | Relatórios consolidados segmentados por unidade (LOJA) ou total global, prontos para consumo da camada de *Report*. |
| **Automação de Documentos (Telegrama)** | Carregamento de *asset* estático (template PDF do formulário padrão dos Correios). Injeção de strings tratadas via quebra em bloco (`textwrap.dedent`) mapeadas por coordenadas predefinidas. | Arquivo `.pdf` achatado contendo o texto legal formatado para disparo/impressão. |

**Estrutura de Diretórios e Interação Arquitetural**

A arquitetura do projeto segue um padrão de **módulos independentes orquestrados por uma interface desacoplada**. A raiz atua como *bootstrap*, enquanto os processos de negócio são contidos em subdiretórios próprios.

**`/ (Raiz)`**
* **`main.py`**: Ponto de entrada (*entry-point*). Configura a instância master do Tkinter e despacha para a interface de menu.
* **`README.md`**: Documentação técnica do sistema.

**`/services`**
* **`interface_master.py`**: Roteador da camada visual. Constrói componentes via sub-frames, barra lateral e aplica propriedades visuais do tema. Vincula comandos de UI aos scripts de execução (módulo de despesas e telegrama).
* **`ui_theme.py`** *(inferido)*: Arquivo de configuração de estilo global, retendo paletas hexadecimais, tipografia e funções puras de limpeza de janelas.

**`/services/despesas`**
* **`main.py`**: Controlador de fluxo do processamento de planilhas. Declara hardcodes de roteamento dos arquivos I/O. Orquestra a injeção dos dados no processador e aciona sequencialmente as sub-rotinas de compilação e relatório de loja a loja.
* **`/services/despesas/services`**
    * **`processador.py`**: Núcleo computacional do sistema. Comporta toda a lógica com bibliotecas matemáticas e de dados (`pandas`, `numpy`). Executa rotinas de limpeza, tratamento dinâmico de cabeçalhos corrompidos, padronização referencial de ADMs e criação de dicionários estatísticos estruturados.
    * **`reporter.py`**: Módulo de escrita. Pega as tabelas higienizadas em memória do processador e persiste as métricas calculadas na saída `.xlsx`.
    * **`dashboard_despesas.py`**: Script de visualização que integra os dados processados na visualização gerencial da ferramenta.

**`/services/telegrama`**
* **`main_telegrama.py`**: Controlador de testes e rotinas de backend para o preenchimento dos PDFs. Submete *strings* *mockadas* ou brutas para injeção.
* **`/services/telegrama/input`**
    * **`Formulário de telegrama - correios.pdf`**: Máscara oficial usada pelo script gerador para a sobreposição de *text box*.
* **`/services/telegrama/services`**
    * **`reporter.py`**: Processador do documento. Localiza coordenadas `X, Y` no template em branco e gera uma nova camada unificada de documento.
    * **`tela_telegrama.py`**: *View* gráfica acessada pelo usuário para *input* dinâmico dos dados demissionários ou comunicados que irão popular o telegrama gerado.