# Automação de Boletos de IPTU

Scripts que baixam os boletos de IPTU de centenas de lotes no portal da prefeitura, organizam os PDFs por comprador e enviam cada boleto por e-mail com comprovante do envio.

![Python](https://img.shields.io/badge/Python-3.12-3776AB) ![Playwright](https://img.shields.io/badge/Playwright-async-2EAD33) ![openpyxl](https://img.shields.io/badge/openpyxl-planilhas-217346) ![status](https://img.shields.io/badge/status-Conclu%C3%ADdo-blue) ![licença](https://img.shields.io/badge/licen%C3%A7a-propriet%C3%A1ria-lightgrey)

> Projeto real. Este repositório é uma **cópia de portfólio**: segredos, servidores e nomes de empresas foram removidos ou trocados pelo nome fictício "Natela".

## Sumário

1. [O que o projeto faz](#1-o-que-o-projeto-faz)
2. [Instalação e requisitos](#2-instalação-e-requisitos)
3. [Como rodar e usar](#3-como-rodar-e-usar)
4. [Demo visual](#4-demo-visual)
5. [Suporte e ajuda](#5-suporte-e-ajuda)
6. [Contribuição e manutenção](#6-contribuição-e-manutenção)
7. [Licença e status](#7-licença-e-status)

## 1. O que o projeto faz

**O que resolve:** Todo ano, a equipe de uma loteadora emitia um a um os boletos de IPTU de cada lote no portal da prefeitura e mandava por e-mail para cada comprador.

**Como resolve:** Um conjunto de scripts lê a planilha de lotes, baixa o boleto de cada CCI no portal, confere e organiza os arquivos, gera relatórios e envia os e-mails pelo Outlook Web, salvando um print de cada envio como prova.

**Para quem é:** Loteadoras e administradoras que emitem e enviam boletos de IPTU em volume.

**Por que importa:** Dias de trabalho manual viram uma execução acompanhada, com retomada de onde parou e prova de cada envio.

### Principais benefícios

| Baixa centenas de boletos | Envia com prova de envio | Retoma de onde parou |
|---|---|---|

### Funcionalidades

- `baixar_iptu.py`: download por CCI com modos debug, teste, headless e retomada por linha
- Diagnóstico com print quando o portal muda (sem PDF, sem prévia, modal de débitos)
- `vistoria_iptus.py`, `checagem_final.py`, `corrigir_estrutura.py`: conferem e organizam as pastas
- `relatorio_completo.py`, `gerar_dashboard.py`: relatórios e painel HTML do andamento
- `enviar_iptu.py`: envio pelo Outlook Web com modo teste, envio por CCI e print de comprovação
- `reprocessar_falhas.py`, `restaurar_marcadores.py`: recuperação de execuções interrompidas

## 2. Instalação e requisitos

### Pré-requisitos

| Ferramenta | Versão |
|---|---|
| Python | 3.12 |
| Playwright | com Chromium |
| openpyxl | 3.x |

### Instalação

```bash
git clone https://github.com/jhonatancamargo121/automacao-iptu.git
cd automacao-iptu
pip install playwright openpyxl
python -m playwright install chromium
```

Os scripts leem a planilha de lotes (`IPTU - CIDADE EXEMPLO.xlsx`) da pasta do projeto ou de Downloads. Nenhuma planilha real está no repositório.

### Variáveis de ambiente

| Variável | Para que serve |
|---|---|
| `NOMES_PLANILHA` | Nomes aceitos para a planilha de lotes |
| `REMETENTE` | Conta do Outlook usada no envio |
| `--teste N` | Processa só N lotes |
| `--producao` | Envia para o e-mail real do comprador |

A lista completa, com comentários, está no `.env.example`. Nenhum valor real é versionado.

## 3. Como rodar e usar

### Quick start

```bash
python baixar_iptu.py --teste 1
python enviar_iptu.py --teste 1
```

Na primeira execução do envio o navegador abre para login e 2FA no Outlook; a sessão fica salva para as próximas.

### Exemplo de uso

Requisição:

```bash
python baixar_iptu.py --teste 5
```

Resposta esperada:

```text
Baixa os boletos dos 5 primeiros lotes pendentes, salva os PDFs na pasta de cada comprador e atualiza a planilha.
```

## 4. Demo visual

Fluxo principal:

```mermaid
flowchart LR
  PL[Planilha de lotes] --> B[Baixar boleto por CCI]
  B --> V[Vistoria e organização das pastas]
  V --> R[Relatório / dashboard]
  V --> E[Enviar por e-mail]
  E --> P[Print do envio como prova]
```

> Prints e GIF com dados fictícios estão no roadmap. Para ver o sistema rodando, siga o [Quick start](#3-como-rodar-e-usar).

## 5. Suporte e ajuda

- **Documentação:**
  - [Uso de cada script (cabeçalho dos arquivos)](baixar_iptu.py)
- **Issues:** abra uma [issue](https://github.com/jhonatancamargo121/automacao-iptu/issues) para reportar problemas ou sugerir melhorias.
- **Discussions:** dúvidas gerais em [Discussions](https://github.com/jhonatancamargo121/automacao-iptu/discussions).
- **Contato:** pelo perfil [@jhonatancamargo121](https://github.com/jhonatancamargo121)

## 6. Contribuição e manutenção

### Como contribuir

1. Faça um **fork** do repositório
2. Crie uma **branch** a partir da `main` (`git checkout -b feat/minha-melhoria`)
3. Abra um **Pull Request** descrevendo o que mudou e por quê

### Padrões

- Commits no padrão [Conventional Commits](https://www.conventionalcommits.org/pt-br/) (`feat:`, `fix:`, `docs:`…)
- Lint e formatação passando antes do PR
- Testes rodando antes do PR

### Testes

Este projeto ainda não tem suíte automatizada; ela está no roadmap.

### Mantenedores

| | |
|---|---|
| **Jhonatan Camargo** ([@jhonatancamargo121](https://github.com/jhonatancamargo121)) | Mantenedor |

## 7. Licença e status

### Licença

Código proprietário, publicado para fins de portfólio. Todos os direitos reservados; veja o arquivo [LICENSE](LICENSE).

### Status

![status](https://img.shields.io/badge/status-Conclu%C3%ADdo-blue)

### Próximos passos

- [ ] Arquivo de configuração único no lugar das constantes
- [ ] Testes da leitura da planilha
