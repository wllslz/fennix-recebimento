# Fennix Brasil — Automação de Recebimento, Controle de Validade e Operação APB

Sistema web desenvolvido em **Python (Flask)** e implantado na **Vercel** (`https://fennix-recebimento.vercel.app`) para automatizar os registros operacionais de recebimento de cargas, geração de informativo com CNH/operadores APB e controle de validade (*shelf life*) da **Fennix Brasil**.

## Funcionalidades Implementadas

1. **Informativo de Recebimento e Operadores APB (Aba 1):**
   - Preenchimento guiado de NF, fornecedor, transportadora, placa, horários e até 4 operadores APB (Nome e CPF).
   - **Upload Mobile Flexível:** Seletor nativo liberado para **Câmera, Galeria e Arquivos** (sem bloqueio em câmera direta).
   - **Compressão Automática Anti-413 (`<canvas>`):** Redimensiona fotos de alta resolução (5–12 MB) para até `1400px` em JPEG (`0.80`, ~250 KB) antes do envio, evitando `HTTP 413 (Request Entity Too Large)` na Vercel.
   - Preenche `template_recebimento_fennix.docx` via `docxtpl` e prepara resumo para WhatsApp.
2. **Controle de Validade / Shelf Life e Planilha Excel (Aba 2):**
   - Cálculo automático da validade mínima aceitável (`60->30`, `75->45`, `90->60`, `120->78`, `150->105`, `180->120`, `270->200`, `360->270` dias).
   - Exportação dupla em `.xlsx` (`openpyxl`) e imagem PNG (`Pillow` + `DejaVuSans`) pronta para WhatsApp.
3. **Resiliência contra Queda de Energia (`localStorage`):**
   - Auto-save contínuo de todos os campos da Aba 1 e linhas da Aba 2 a cada digitação, restaurando tudo automaticamente caso falte energia ou a aba seja fechada.
