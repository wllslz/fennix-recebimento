# Fennix Brasil — Sistema de Recebimento, Controle de Validade e Operação APB

> Aplicação web operacional para a rotina de recebimento de cargas da **Fennix Brasil**, construída com **Python (Flask)** e implantada em arquitetura *serverless* na **Vercel** (`https://fennix-recebimento.vercel.app`).

## Sumário

1. [Porquê o Sistema Foi Criado](#porquê-o-sistema-foi-criado)
2. [Como Foi Construído e Decisões de Engenharia](#como-foi-construído-e-decisões-de-engenharia)
3. [O Que o Aplicativo Faz (Aba 1 e Aba 2)](#o-que-o-aplicativo-faz-aba-1-e-aba-2)
4. [Arquitetura e Fluxo de Funcionamento](#arquitetura-e-fluxo-de-funcionamento)
5. [Estrutura do Repositório](#estrutura-do-repositório)
6. [Resiliência Mobile, Anti-413, OCR e Queda de Energia](#resiliência-mobile-anti-413-ocr-e-queda-de-energia)
7. [Como Executar e Testar Localmente](#como-executar-e-testar-localmente)

---

## Porquê o Sistema Foi Criado

No fluxo diário de recebimento de mercadorias, conferência de documentos de motoristas e validação de *shelf life* (prazo de validade), os processos manuais apresentavam gargalos recorrentes:

- **Lentidão e repetição manual:** Preencher dados de nota fiscal, fornecedor, transportadora, placa, horários de chegada e descarga, operadores APB e anexar CNH de motoristas exigia tempo e atenção constante na doca.
- **Cálculo manual de validade mínima:** Conferir manualmente a margem mínima aceitável para diferentes prazos de validade aumentava a chance de erro humano em lotes críticos.
- **Uso no celular em ambiente operacional:** A execução via smartphone na operação logística esbarrava em bloqueios de câmera, fotos pesadas (5–12 MB) que estouravam o limite de payload da Vercel (`HTTP 413`) e risco de perda de preenchimento por queda de energia ou fechamento acidental do navegador.

O sistema foi criado para unificar essas rotinas em uma interface *mobile-first* rápida, auditável e tolerante a falhas.

---

## Como Foi Construído e Decisões de Engenharia

- **Backend em Python + Flask (`app.py`):** Responsável pelas rotas principais, regras de cálculo de validade mínima (`REGRAS_SHELF` / `calcular_val_min`), preenchimento de modelos `.docx` via `docxtpl`, atualização de planilhas `.xlsx` via `openpyxl`, renderização de imagens `.png` via `Pillow` com fontes `DejaVuSans` e tratamento global de erros HTTP em JSON (`413` e `500`).
- **Adaptação Serverless para Vercel (`vercel.json` e `IS_VERCEL`):** Como a execução *serverless* utiliza sistema de arquivos efêmero, caminhos de saída e banco SQLite de apoio são direcionados para `/tmp` quando `VERCEL` está ativo, evitando erros de permissão ou diretório ausente.
- **Processamento no navegador (`templates/index.html`):** Para reduzir latência e respeitar o limite de requisições da Vercel (~4,5 MB), a compressão de imagem é feita no cliente via HTML5 `<canvas>`, a leitura OCR da CNH roda no navegador com `Tesseract.js` e o auto-save contínuo persiste os rascunhos em `localStorage`.

---

## O Que o Aplicativo Faz (Aba 1 e Aba 2)

### Aba 1 — Informativo de Recebimento, Operadores APB e CNH

- **Dados operacionais da carga:** Registra nota fiscal, fornecedor, transportadora, placa do veículo, data de recebimento, horários de chegada e descarga, conferente, qualidade e motorista.
- **Operadores APB (carga e descarga):** Permite informar até 4 operadores (Nome e CPF) ou deixar em branco para preenchimento manual na folha impressa.
- **Foto da CNH com seletor flexível e OCR:** Remove restrição de câmera forçada no celular (liberando Câmera, Galeria e Arquivos), exibe pré-visualização, executa OCR com recortes direcionados e filtragem de datas para sugerir **Nome completo do motorista** e **Nº Registro CNH**, e comprime a imagem antes do envio.
- **Geração de `.docx` e resumo para WhatsApp:** Preenche `template_recebimento_fennix.docx` e monta mensagem pronta para copiar ou compartilhar no WhatsApp.

### Aba 2 — Controle de Validade (*Shelf Life*) e Planilha Excel

- **Tabela de tolerância de *shelf life*:** Calcula automaticamente a validade mínima exigida a partir do prazo total de validade do produto:

| Shelf Life Total | Validade Mínima Exigida |
| :---: | :---: |
| 60 dias | 30 dias |
| 75 dias | 45 dias |
| 90 dias | 60 dias |
| 120 dias | 78 dias |
| 150 dias | 105 dias |
| 180 dias | 120 dias |
| 270 dias | 200 dias |
| 360 dias | 270 dias |

- **Catálogo de produtos (`static/catalogo_produtos.json`):** Apoia o preenchimento rápido por código.
- **Exportação dupla (`.xlsx` + `.png`):** Gera a planilha atualizada a partir de `base_recebimento.xlsx` e renderiza uma imagem PNG legível da tabela para envio rápido no WhatsApp.

---

## Arquitetura e Fluxo de Funcionamento

1. **Preenchimento e Auto-Save:** Cada evento `input`/`change` na Aba 1 ou Aba 2 salva o estado atual em `localStorage`. Ao reabrir a página após queda de energia ou recarregamento, os dados são restaurados automaticamente (e só são apagados ao confirmar a limpeza do formulário).
2. **Anexo e OCR da CNH:** Ao selecionar a imagem, o navegador exibe o preview e aciona a leitura OCR em recortes (faixa superior de nome, visão geral e lateral vertical de registro), preenchendo motorista e registro CNH e atualizando o rascunho.
3. **Compressão e Envio (`POST /gerar`):** A imagem é redimensionada/comprimida via `<canvas>` (até `1400px`, JPEG) antes do envio via `FormData`. O frontend usa `lerJsonSeguro` para validar `resp.ok` e tratar falhas HTTP com clareza.
4. **Geração de Planilha (`POST /gerar-planilha`):** Recebe `data_ref`, `somente_abaixo` e `itens`, calcula prazos e validade mínima, atualiza o `.xlsx` via `openpyxl` e gera o `.png` via `Pillow`.

---

## Estrutura do Repositório

- `app.py`: Aplicação Flask, regras de shelf life, geração `.docx`/`.xlsx`/`.png` e handlers JSON de erro (`413` e `500`).
- `vercel.json`: Configuração de deploy *serverless* na Vercel.
- `requirements.txt`: Dependências Python (`Flask`, `docxtpl`, `python-docx`, `Pillow`, `openpyxl`).
- `template_recebimento_fennix.docx`: Modelo Word oficial do informativo de recebimento.
- `base_recebimento.xlsx`: Planilha modelo para controle de validade.
- `static/catalogo_produtos.json`: Base de produtos para autocompletar.
- `static/fonts/DejaVuSans.ttf` e `DejaVuSans-Bold.ttf`: Fontes usadas na renderização das imagens PNG.
- `templates/index.html`: Interface web, compressão de imagem, OCR de CNH e persistência em `localStorage`.
- `README.md` e `CHANGELOG.md`: Documentação técnica e histórico de commits.

---

## Resiliência Mobile, Anti-413, OCR e Queda de Energia

- **Galeria e arquivos no celular:** Remoção de `capture="environment"` no input `foto_cnh`.
- **Prevenção ao erro HTTP 413:** Compressão *client-side* em `<canvas>` antes do upload e handler JSON `413`/`500` no Flask.
- **Proteção contra queda de energia:** Auto-save contínuo em `localStorage` (`fennix_recebimento_rascunho_aba1_v1` e `fennix_recebimento_rascunho_aba2_v1`).
- **Histórico completo:** Veja [`CHANGELOG.md`](./CHANGELOG.md) para os detalhes de cada commit.

---

## Como Executar e Testar Localmente

```bash
cd fennix_recebimento_vercel
python -m venv venv
source venv/Scripts/activate
pip install -r requirements.txt
python app.py
```

Acesse `http://127.0.0.1:5000` no navegador.
