# Suite de Produção

Aplicativo desktop em Python para extrair relatórios de produção e lançar quantidades em planilhas Excel, com conferência dos resultados antes da gravação.

A interface reúne três ferramentas e um editor de configurações. O projeto foi desenvolvido para uma rotina de produção de móveis e usa layouts específicos de planilhas.

## Recursos

- **Extrator de Produção:** seleciona funcionários e um intervalo de dias para gerar um relatório em texto.
- **Cadeiras / Mesas:** interpreta relatórios colados na interface e associa os produtos aos catálogos das planilhas.
- **Mensal por Funcionário:** associa funcionário, produto e dia para lançar quantidades nas abas individuais.
- **Conferência:** busca aproximada de nomes, revisão de correspondências e correção manual dos resultados.
- **Configurações:** grupos de funcionários, exclusões, regras de produtos, limites de reconhecimento e histórico de arquivos.
- **Saída:** grava cópias com o sufixo `_ATUALIZADO`.

## Requisitos

- Python **3.10 ou superior**, com Tkinter.
- Dependências: `openpyxl` e `rapidfuzz`.
- Windows é o ambiente principal dos atalhos `.bat`. Nos demais sistemas, execute pelo terminal e disponibilize o Tkinter.
- Planilhas `.xlsx` ou `.xlsm`; `.xls` não é aceito.

## Instalação no Windows

1. Baixe o projeto pelo botão **Code → Download ZIP** no GitHub e extraia a pasta.
2. Instale o Python com a opção de adicioná-lo ao PATH.
3. Execute `INSTALAR_DEPENDENCIAS.bat`.
4. Execute `INICIAR.bat`.

Deixe os seis arquivos `.py` e o `requirements.txt` na mesma pasta dos atalhos.

### Pelo terminal

Na pasta do projeto:

```bash
python -m venv .venv
```

Ative o ambiente:

```powershell
# Windows (Prompt de Comando)
.venv\Scripts\activate.bat
```

```bash
# Linux / macOS
source .venv/bin/activate
```

Instale e abra:

```bash
python -m pip install -r requirements.txt
python suite_producao.py
```

## Primeiro uso

1. Abra **Configurações** e cadastre os grupos e funcionários que você utiliza. A distribuição começa sem nomes reais.
2. Confira as regras automáticas de produtos: os padrões incluídos são voltados a móveis e devem ser adaptados ao seu catálogo.
3. Escolha a ferramenta e selecione as planilhas compatíveis.
4. Gere ou cole o relatório, analise as correspondências e confira os itens pendentes.
5. Grave o resultado e confira a cópia `_ATUALIZADO`.

Exemplo de texto, com dados fictícios:

```text
01/10

**FUNCIONARIO EXEMPLO**
5 CADEIRA MODELO A
2 MESA MODELO B

---
```

## Layouts das planilhas

O aplicativo não interpreta qualquer planilha automaticamente:

| Ferramenta | Estrutura esperada |
| --- | --- |
| Extrator | Uma aba por funcionário; célula com `PRODUTOS`; linha de dias de 1 a 31 logo abaixo; produtos e quantidades nas linhas seguintes. |
| Cadeiras / Mesas | Aba `PRODUÇÃO-DIA-MÊS-QUANT`; cabeçalho `ESPECIFICAÇÃO` na coluna C; dias nas colunas I a AE desse cabeçalho; produtos a partir de duas linhas abaixo. |
| Individual | Uma aba por funcionário; indicação de `DIA` no início da aba; dias na linha 2 a partir da coluna B; nomes dos produtos na coluna A a partir da linha 3 (constantes `DIA_HEADER_ROW` e `PRODUTO_COL`). |

O reconhecimento aproximado pode exigir correções. Revise os resultados antes de gravar e mantenha backups das planilhas. A preservação de VBA em `.xlsm` não garante a preservação de todo recurso avançado do Excel.

## Configurações e dados

As preferências são salvas fora do código:

- Windows: `%APPDATA%\SuiteProducao\config.json`.
- Linux / macOS: `$XDG_CONFIG_HOME/SuiteProducao/config.json`, ou `~/.config/SuiteProducao/config.json`.
- Opcional: defina `SUITE_PRODUCAO_CONFIG_DIR` para escolher outra pasta.

Atualizar o código não apaga as configurações existentes. A interface permite exportar e importar backups.

As planilhas são escolhidas localmente. O `.gitignore` exclui planilhas de trabalho, configurações e arquivos temporários para evitar seu envio acidental ao repositório.

## Arquivos

| Arquivo | Função |
| --- | --- |
| `suite_producao.py` | Aplicativo principal. |
| `extrator_producao.py` | Extração de relatórios. |
| `lancador_producao.py` | Lançamentos em Cadeiras / Mesas. |
| `lancador_individual.py` | Lançamentos por funcionário. |
| `configuracao.py` | Persistência e edição de preferências. |
| `tema.py` | Interface, estilos e tarefas em segundo plano. |
| `requirements.txt` | Dependências externas. |

Cada uma das três ferramentas também pode ser executada diretamente com `python nome_do_arquivo.py`.

## Licença

Este projeto ainda não define uma licença de uso ou redistribuição. Tornar o repositório público permite consultar o código, mas não equivale a conceder uma licença de código aberto.
