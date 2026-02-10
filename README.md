<p align="center">
  <img src="assets/DomBot_New.png" alt="DomBot Logo" width="120">
</p>

<h1 align="center">DomBot - Folha de Ponto</h1>

<p align="center">
  Automacao RPA para geracao e publicacao de Folhas de Ponto no sistema Dominio Folha (Thomson Reuters).
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/GUI-CustomTkinter-1ABC9C?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/automation-pywinauto-F39C12?style=for-the-badge" alt="pywinauto">
  <img src="https://img.shields.io/badge/license-MIT-2ECC71?style=for-the-badge" alt="License">
</p>

---

## Sobre

O **DomBot - Folha de Ponto** automatiza o processo repetitivo de gerar relatorios de Folha de Ponto para diversas empresas no sistema **Dominio Folha**. A partir de uma planilha Excel contendo os dados das empresas, o bot executa todo o fluxo de forma autonoma:

1. Troca de empresa via `F8`
2. Navegacao ate o Gerenciador de Relatorios Integrados
3. Preenchimento dos campos de data
4. Execucao e publicacao do relatorio na categoria **Pessoal/Folha de Ponto**
5. Exportacao e salvamento do PDF

## Funcionalidades

- **Interface grafica moderna** com tema escuro (CustomTkinter)
- **Painel de estatisticas** em tempo real (total, sucesso, erros, empresa atual, tempo decorrido)
- **Barra de progresso** com porcentagem
- **Sistema de logs** com cores por nivel (sucesso, erro, aviso, info)
- **Controles de execucao** - Iniciar, Pausar/Retomar e Parar
- **Preview da planilha** antes de processar
- **Exportacao de logs** para arquivo `.txt`
- **Logs persistentes** em arquivo (sucesso e erro separados por data)
- **Tratamento de janelas auxiliares** (Avisos de Vencimento, dialogos de erro)

## Pre-requisitos

- **Windows 10/11**
- **Python 3.10+**
- **Dominio Folha** (Thomson Reuters) instalado e aberto

## Instalacao

```bash
# Clone o repositorio
git clone https://github.com/seu-usuario/DomBot_FPonto.git
cd DomBot_FPonto

# Instale as dependencias
pip install customtkinter pandas pywinauto pywin32 pillow openpyxl
```

## Planilha Excel

A planilha deve conter as seguintes colunas:

| Coluna | Descricao | Exemplo |
|---|---|---|
| `Nº` | Codigo da empresa no Dominio | `123` |
| `EMPRESA` | Nome da empresa (opcional, para logs) | `Empresa XYZ Ltda` |
| `data inicio` | Data inicial do periodo | `01/01/2025` |
| `data final` | Data final do periodo | `31/01/2025` |
| `nome pdf` | Nome do arquivo PDF a ser salvo | `FolhaPonto_Jan2025` |

> A linha 1 e o cabecalho. O processamento inicia a partir da linha 2 por padrao.

## Uso

1. Abra o **Dominio Folha** e faca login
2. Execute o bot:

```bash
python DomBot-FolhaPonto.py
```

3. Na interface, clique em **Procurar** e selecione a planilha Excel
4. Ajuste a **linha inicial** se necessario (padrao: 2)
5. Clique em **Iniciar**

## Estrutura do Projeto

```
DomBot_FPonto/
├── DomBot-FolhaPonto.py   # Aplicacao principal (GUI + automacao)
├── assets/
│   ├── DomBot_New.png     # Logo do bot
│   └── favicon.ico        # Icone da janela
├── logs/                  # Logs gerados automaticamente
│   ├── success_YYYY-MM-DD.log
│   └── error_YYYY-MM-DD.log
├── .gitignore
└── README.md
```

## Fluxo da Automacao

```
┌─────────────────────┐
│  Carregar planilha  │
└────────┬────────────┘
         v
┌─────────────────────┐
│ Conectar ao Dominio │
└────────┬────────────┘
         v
┌─────────────────────┐     ┌──────────────────────┐
│  F8 - Trocar        │────>│  Fechar Avisos de    │
│  Empresa            │     │  Vencimento (se ha)  │
└────────┬────────────┘     └──────────┬───────────┘
         v                             v
┌─────────────────────┐     ┌──────────────────────┐
│  Abrir Relatorios   │────>│  Navegar na arvore   │
│  Integrados         │     │  de relatorios       │
└────────┬────────────┘     └──────────┬───────────┘
         v                             v
┌─────────────────────┐     ┌──────────────────────┐
│  Preencher datas    │────>│  Executar relatorio  │
└────────┬────────────┘     └──────────┬───────────┘
         v                             v
┌─────────────────────┐     ┌──────────────────────┐
│  Publicar documento │────>│  Gerar e salvar PDF  │
│  (Pessoal/FPonto)   │     │                      │
└────────┬────────────┘     └──────────┬───────────┘
         v                             v
┌─────────────────────┐     ┌──────────────────────┐
│  Fechar janelas     │────>│  Proxima empresa     │
└─────────────────────┘     └──────────────────────┘
```

## Dependencias

| Pacote | Versao | Uso |
|---|---|---|
| `customtkinter` | >= 5.0 | Interface grafica moderna |
| `pandas` | >= 2.0 | Leitura de planilhas Excel |
| `pywinauto` | >= 0.6.8 | Automacao de janelas Windows |
| `pywin32` | >= 306 | API nativa do Windows (win32gui) |
| `Pillow` | >= 10.0 | Manipulacao de imagens (logo) |
| `openpyxl` | >= 3.1 | Engine de leitura `.xlsx` |

## Observacoes

- **Nao mova o mouse** durante a execucao — o bot interage com a interface via cliques e teclas
- O Dominio Folha deve estar **aberto e logado** antes de iniciar
- Logs de sucesso e erro sao salvos automaticamente na pasta `logs/`
- Use o botao **Pausar** para interromper temporariamente sem perder o progresso

---

<p align="center">
  Desenvolvido para automatizar processos contabeis repetitivos no Dominio Folha.
</p>
