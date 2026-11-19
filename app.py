import os
import io
import json
import base64
import subprocess
import shutil
import sqlite3
from datetime import datetime, date
from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image, ImageOps, ImageDraw, ImageFont
import openpyxl

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template_recebimento_fennix.docx")
BASE_XLSX_PATH = os.path.join(BASE_DIR, "base_recebimento.xlsx")
CATALOGO_PATH = os.path.join(BASE_DIR, "static", "catalogo_produtos.json")

IS_VERCEL = bool(os.environ.get("VERCEL"))
if IS_VERCEL:
    OUTPUT_DIR = "/tmp/gerados"
    DB_PATH = "/tmp/historico_recebimento.db"
else:
    OUTPUT_DIR = os.path.join(BASE_DIR, "gerados")
    DB_PATH = os.path.join(BASE_DIR, "historico_recebimento.db")

os.makedirs(OUTPUT_DIR, exist_ok=True)

REGRAS_SHELF = [
    (60, 30),
    (75, 45),
    (90, 60),
    (120, 78),
    (150, 105),
    (180, 120),
    (270, 200),
    (360, 270),
]

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS recebimentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                criado_em TEXT,
                data_recebimento TEXT,
                nf TEXT,
                cte TEXT,
                bonus TEXT,
                lacre TEXT,
                container TEXT,
                hora_chegada TEXT,
                hora_inicio TEXT,
                hora_fim TEXT,
                hora_saida TEXT,
                placa_carreta TEXT,
                placa_veiculo TEXT,
                transportadora TEXT,
                fornecedor TEXT,
                motorista TEXT,
                cnh TEXT,
                conferente TEXT,
                qualidade TEXT,
                paletes TEXT,
                arquivo_docx TEXT,
                arquivo_pdf TEXT
            )
        """)

init_db()

def carregar_fontes():
    """Carrega fontes TrueType limpas (DejaVuSans / Arial / Calibri) para renderizar a tabela estilo Excel."""
    candidatos_bold = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "C:\\Windows\\Fonts\\calibrib.ttf",
        "C:\\Windows\\Fonts\\arialbd.ttf",
    ]
    candidatos_reg = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "C:\\Windows\\Fonts\\calibri.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ]
    f_bold = None
    f_reg = None
    for p in candidatos_bold:
        if os.path.exists(p):
            f_bold = ImageFont.truetype(p, 21)
            break
    for p in candidatos_reg:
        if os.path.exists(p):
            f_reg = ImageFont.truetype(p, 20)
            break
    if f_bold is None:
        f_bold = ImageFont.load_default()
    if f_reg is None:
        f_reg = ImageFont.load_default()
    return f_bold, f_reg

def calcular_val_min(dif_dias):
    """Replica exatamente =VLOOKUP(F2, DADOS!A:C, 3, 0) da planilha original."""
    for dias, val_min in REGRAS_SHELF:
        if dias == dif_dias:
            return val_min
    # Fallback aproximado caso a validade total não seja exata
    escolhido = None
    for dias, val_min in REGRAS_SHELF:
        if dif_dias >= dias:
            escolhido = val_min
    return escolhido

def formatar_qtd_br(qtd_raw):
    s = str(qtd_raw).strip()
    if not s:
        return ""
    try:
        val = float(s.replace(",", "."))
        if val.is_integer():
            return str(int(val))
        return f"{val:g}".replace(".", ",")
    except Exception:
        return s

def renderizar_imagem_tabela_excel(linhas):
    """Gera uma imagem PNG fiel à captura de tela da tabela filtrada do Excel (pronta para enviar no WhatsApp)."""
    f_bold, f_reg = carregar_fontes()

    # Larguras proporcionais exatas às colunas visíveis da captura de tela:
    # Col A (Código): 145px | Col B (DESCRIÇÃO): 680px | Col C (QTD): 175px | Col D (FABRICAÇÃO): 215px | Col E (VALIDADE): 205px | Col I (ABAIXO): 135px
    col_widths = [145, 680, 175, 215, 205, 135]
    headers = ["", "DESCRIÇÃO", "QTD", "FABRICAÇÃO", "VALIDADE", "ABAIXO"]
    total_w = sum(col_widths)
    header_h = 46
    row_h = 42
    total_h = header_h + max(len(linhas), 1) * row_h

    img = Image.new("RGB", (total_w, total_h), "#D9D9D9")
    draw = ImageDraw.Draw(img)

    # Cores idênticas à captura de tela enviada:
    # Coluna A do cabeçalho: laranja vivo (#ED7D31), Colunas B..I do cabeçalho: terracota/marrom-alaranjado (#C65911)
    # Linhas de dados: cinza claro uniforme (#D9D9D9) com grade fina (#808080) e borda externa verde Excel (#217346)
    cor_header_a = "#ED7D31"
    cor_header_resto = "#C65911"
    cor_linha = "#D9D9D9"
    cor_grade = "#8C8C8C"
    cor_texto = "#000000"

    # Desenhar cabeçalho
    x = 0
    for idx, (w, titulo) in enumerate(zip(col_widths, headers)):
        bg = cor_header_a if idx == 0 else cor_header_resto
        draw.rectangle([x, 0, x + w, header_h], fill=bg, outline=cor_grade, width=1)

        if titulo:
            bbox = draw.textbbox((0, 0), titulo, font=f_bold)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            # Centraliza descontando o botão de filtro à direita (24px)
            tx = x + (w - 22 - tw) // 2
            ty = (header_h - th) // 2 - 2
            draw.text((tx, ty), titulo, fill=cor_texto, font=f_bold)

        # Desenhar botão quadrado de filtro do Excel no canto inferior direito da célula do cabeçalho
        btn_size = 20
        bx1 = x + w - btn_size - 4
        by1 = header_h - btn_size - 4
        bx2 = x + w - 4
        by2 = header_h - 4
        draw.rectangle([bx1, by1, bx2, by2], fill="#E6E6E6", outline="#7A7A7A", width=1)
        # Triângulo preto apontando para baixo
        cx = (bx1 + bx2) // 2
        cy = (by1 + by2) // 2
        draw.polygon([(cx - 4, cy - 2), (cx + 4, cy - 2), (cx, cy + 3)], fill="#333333")

        x += w

    # Desenhar linhas de dados
    for r_idx, item in enumerate(linhas):
        y1 = header_h + r_idx * row_h
        y2 = y1 + row_h
        vals = [
            str(item.get("codigo", "")),
            str(item.get("descricao", "")),
            str(item.get("qtd_fmt", "")),
            str(item.get("fab_fmt", "")),
            str(item.get("val_fmt", "")),
            str(item.get("abaixo", "")),
        ]
        x = 0
        for c_idx, (w, val) in enumerate(zip(col_widths, vals)):
            draw.rectangle([x, y1, x + w, y2], fill=cor_linha, outline=cor_grade, width=1)
            if val:
                bbox = draw.textbbox((0, 0), val, font=f_reg)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                # Se a descrição for muito longa, ajusta com margem mínima
                if tw > w - 16:
                    tx = x + 8
                else:
                    tx = x + (w - tw) // 2
                ty = y1 + (row_h - th) // 2 - 2
                draw.text((tx, ty), val, fill=cor_texto, font=f_reg)
            x += w

    # Borda externa fina estilo seleção do Excel
    draw.rectangle([0, 0, total_w - 1, total_h - 1], outline="#217346", width=2)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()

def gerar_xlsx_recebimento(linhas, data_ref_str):
    """Gera uma cópia da planilha base preenchida na aba Folha4 com todas as fórmulas originais."""
    wb = openpyxl.load_workbook(BASE_XLSX_PATH)
    ws = wb["Folha4"]

    # Atualiza data de referência em J1 (=TODAY() ou data informada)
    try:
        dt_ref = datetime.strptime(data_ref_str, "%d/%m/%Y")
        ws["J1"].value = dt_ref
    except Exception:
        pass

    # Limpa linhas antigas de dados (linhas 2 até 60) e garante que todas fiquem visíveis
    for r in range(2, 65):
        ws.row_dimensions[r].hidden = False
        ws.cell(r, 1).value = None
        ws.cell(r, 3).value = None
        ws.cell(r, 4).value = None
        ws.cell(r, 5).value = None

    # Preenche a partir da linha 2 com os itens informados
    for idx, item in enumerate(linhas):
        r = 2 + idx
        cod_str = str(item.get("codigo", "")).strip()
        ws.cell(r, 1).value = int(cod_str) if cod_str.isdigit() else cod_str
        try:
            ws.cell(r, 3).value = float(str(item.get("qtd", "0")).replace(",", "."))
        except Exception:
            ws.cell(r, 3).value = item.get("qtd", "")

        if item.get("fab_iso"):
            ws.cell(r, 4).value = datetime.strptime(item["fab_iso"], "%Y-%m-%d")
            ws.cell(r, 4).number_format = "dd/mm/yyyy"
        if item.get("val_iso"):
            ws.cell(r, 5).value = datetime.strptime(item["val_iso"], "%Y-%m-%d")
            ws.cell(r, 5).number_format = "dd/mm/yyyy"

    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()

def preparar_imagem_cnh(file_storage, temp_path):
    """Corrige rotação EXIF da câmera do celular e calcula dimensões ideais (mm) para caber em 1 página."""
    img = Image.open(file_storage.stream)
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail((1400, 1400), Image.Resampling.LANCZOS)
    img.save(temp_path, format="JPEG", quality=90)

    w_px, h_px = img.size
    aspect = h_px / float(w_px)
    target_w = 98.0
    target_h = target_w * aspect
    if target_h > 55.0:
        target_h = 55.0
        target_w = target_h / aspect
    return target_w, target_h

def converter_para_pdf(caminho_docx, pasta_saida):
    nome_pdf = os.path.splitext(os.path.basename(caminho_docx))[0] + ".pdf"
    caminho_pdf = os.path.join(pasta_saida, nome_pdf)

    if os.name == "nt":
        try:
            import pythoncom
            from docx2pdf import convert
            pythoncom.CoInitialize()
            convert(caminho_docx, caminho_pdf)
            if os.path.exists(caminho_pdf):
                return nome_pdf
        except Exception:
            pass

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", pasta_saida, caminho_docx],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=30,
            )
            if os.path.exists(caminho_pdf):
                return nome_pdf
        except Exception:
            pass

    return None

def montar_texto_whatsapp(dados):
    return (
        f"*INFORMATIVO DO RECEBIMENTO - FENNIX FRIOS*\n"
        f"📅 *Data:* {dados['data']}\n"
        f"🏭 *Fornecedor:* {dados['fornecedor']}\n"
        f"📄 *NF:* {dados['nf']}\n"
        f"🚚 *CT-e:* {dados['cte'] or '-'} | *Bônus:* {dados['bonus'] or '-'}\n"
        f"🔒 *Lacre:* {dados['lacre'] or '-'} | *Container:* {dados['container']}\n"
        f"🚛 *Placa:* {dados['placa_veiculo']} (Carreta: {dados['placa_carreta']})\n"
        f"🏢 *Transportadora:* {dados['transportadora']}\n"
        f"👤 *Motorista:* {dados['motorista']} (CNH: {dados['cnh']})\n"
        f"📦 *Paletes:* {dados['paletes']}\n"
        f"⏱️ *Horários:* Chegada {dados['hora_chegada'] or '--:--'} | Início {dados['hora_inicio'] or '--:--'} | Fim {dados['hora_fim'] or '--:--'} | Saída {dados['hora_saida'] or '--:--'}\n"
        f"✅ *Conferente:* {dados['conferente']} | *Qualidade:* {dados['qualidade']}"
    )

@app.route("/")
def index():
    historico = []
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            historico = conn.execute(
                "SELECT * FROM recebimentos ORDER BY id DESC LIMIT 15"
            ).fetchall()
    except Exception:
        pass
    return render_template("index.html", historico=historico, is_vercel=IS_VERCEL)

@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form
    traco_nome = "____________________________"
    traco_cpf = "_____________________"

    dados = {
        "data": form.get("data", "").strip(),
        "nf": form.get("nf", "").strip(),
        "cte": form.get("cte", "").strip(),
        "bonus": form.get("bonus", "").strip(),
        "lacre": form.get("lacre", "").strip(),
        "container": (form.get("container", "").strip() or "XXX").upper(),
        "hora_chegada": form.get("hora_chegada", "").strip(),
        "hora_inicio": form.get("hora_inicio", "").strip(),
        "hora_fim": form.get("hora_fim", "").strip(),
        "hora_saida": form.get("hora_saida", "").strip(),
        "placa_carreta": (form.get("placa_carreta", "").strip() or "XXX-XXXX").upper(),
        "placa_veiculo": form.get("placa_veiculo", "").strip().upper(),
        "transportadora": form.get("transportadora", "").strip().upper(),
        "fornecedor": form.get("fornecedor", "").strip().upper(),
        "motorista": form.get("motorista", "").strip().upper(),
        "cnh": form.get("cnh", "").strip(),
        "conferente": form.get("conferente", "").strip().upper(),
        "qualidade": form.get("qualidade", "").strip().upper(),
        "paletes": form.get("paletes", "").strip(),
        "op1_nome": form.get("op1_nome", "").strip().upper() or traco_nome,
        "op1_cpf": form.get("op1_cpf", "").strip() or traco_cpf,
        "op2_nome": form.get("op2_nome", "").strip().upper() or traco_nome,
        "op2_cpf": form.get("op2_cpf", "").strip() or traco_cpf,
        "op3_nome": form.get("op3_nome", "").strip().upper() or traco_nome,
        "op3_cpf": form.get("op3_cpf", "").strip() or traco_cpf,
        "op4_nome": form.get("op4_nome", "").strip().upper() or traco_nome,
        "op4_cpf": form.get("op4_cpf", "").strip() or traco_cpf,
    }

    doc = DocxTemplate(TEMPLATE_PATH)

    temp_img = None
    foto = request.files.get("foto_cnh")
    if foto and foto.filename:
        temp_img = os.path.join(OUTPUT_DIR, f"_tmp_cnh_{datetime.now().timestamp()}.jpg")
        w_mm, h_mm = preparar_imagem_cnh(foto, temp_img)
        dados["foto_cnh"] = InlineImage(doc, temp_img, width=Mm(w_mm), height=Mm(h_mm))
    else:
        dados["foto_cnh"] = ""

    doc.render(dados)

    data_slug = dados["data"].replace("/", ".") or datetime.now().strftime("%d.%m.%Y")
    forn_slug = "".join(c for c in dados["fornecedor"] if c.isalnum() or c in (" ", "-", "_")).strip() or "CARGA"
    base_nome = f"INFO RECEBIMENTO - {data_slug} - {forn_slug}"
    nome_docx = f"{base_nome}.docx"
    caminho_docx = os.path.join(OUTPUT_DIR, nome_docx)
    doc.save(caminho_docx)

    if temp_img and os.path.exists(temp_img):
        os.remove(temp_img)

    with open(caminho_docx, "rb") as f:
        docx_b64 = base64.b64encode(f.read()).decode("utf-8")

    nome_pdf = converter_para_pdf(caminho_docx, OUTPUT_DIR)
    pdf_b64 = None
    if nome_pdf:
        caminho_pdf = os.path.join(OUTPUT_DIR, nome_pdf)
        if os.path.exists(caminho_pdf):
            with open(caminho_pdf, "rb") as f:
                pdf_b64 = base64.b64encode(f.read()).decode("utf-8")

    texto_wpp = montar_texto_whatsapp(dados)

    rec_id = 0
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.execute("""
                INSERT INTO recebimentos (
                    criado_em, data_recebimento, nf, cte, bonus, lacre, container,
                    hora_chegada, hora_inicio, hora_fim, hora_saida,
                    placa_carreta, placa_veiculo, transportadora, fornecedor,
                    motorista, cnh, conferente, qualidade, paletes, arquivo_docx, arquivo_pdf
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                dados["data"], dados["nf"], dados["cte"], dados["bonus"], dados["lacre"], dados["container"],
                dados["hora_chegada"], dados["hora_inicio"], dados["hora_fim"], dados["hora_saida"],
                dados["placa_carreta"], dados["placa_veiculo"], dados["transportadora"], dados["fornecedor"],
                dados["motorista"], dados["cnh"], dados["conferente"], dados["qualidade"], dados["paletes"],
                nome_docx, nome_pdf
            ))
            rec_id = cur.lastrowid
    except Exception:
        pass

    return jsonify({
        "ok": True,
        "id": rec_id,
        "arquivo_docx": nome_docx,
        "arquivo_pdf": nome_pdf,
        "docx_base64": docx_b64,
        "pdf_base64": pdf_b64,
        "url_docx": f"/baixar/{rec_id}/docx" if rec_id else None,
        "url_pdf": f"/baixar/{rec_id}/pdf" if (rec_id and nome_pdf) else None,
        "texto_whatsapp": texto_wpp
    })

@app.route("/gerar-planilha", methods=["POST"])
def gerar_planilha():
    payload = request.get_json(force=True)
    data_ref_str = payload.get("data_ref", "").strip()
    somente_abaixo = bool(payload.get("somente_abaixo", True))
    itens_raw = payload.get("itens", [])

    try:
        dt_ref = datetime.strptime(data_ref_str, "%d/%m/%Y").date()
    except Exception:
        dt_ref = date.today()
        data_ref_str = dt_ref.strftime("%d/%m/%Y")

    with open(CATALOGO_PATH, "r", encoding="utf-8") as f:
        catalogo = json.load(f)

    linhas_calculadas = []
    linhas_filtradas = []

    for it in itens_raw:
        cod = str(it.get("codigo", "")).strip()
        qtd_raw = str(it.get("qtd", "")).strip()
        fab_iso = str(it.get("fab", "")).strip()
        val_iso = str(it.get("val", "")).strip()
        desc_manual = str(it.get("descricao", "")).strip().upper()

        if not cod and not desc_manual:
            continue

        descricao = desc_manual or catalogo.get(cod, "PRODUTO NÃO CADASTRADO")
        qtd_fmt = formatar_qtd_br(qtd_raw)

        fab_fmt = ""
        val_fmt = ""
        abaixo_val = ""
        abaixo_num = None

        if fab_iso and val_iso:
            try:
                dt_fab = datetime.strptime(fab_iso, "%Y-%m-%d").date()
                dt_val = datetime.strptime(val_iso, "%Y-%m-%d").date()
                fab_fmt = dt_fab.strftime("%d/%m/%Y")
                val_fmt = dt_val.strftime("%d/%m/%Y")

                # Fórmulas idênticas à Folha4:
                # DIF = VALIDADE - FABRICAÇÃO
                # BASE 1 = DIF - (HOJE - FABRICAÇÃO) = VALIDADE - HOJE
                # BASE 2 = VLOOKUP(DIF, DADOS!A:C, 3, 0)
                # ABAIXO = BASE 1 - BASE 2
                dif = (dt_val - dt_fab).days
                base1 = dif - (dt_ref - dt_fab).days
                base2 = calcular_val_min(dif)
                if base2 is not None:
                    abaixo_num = base1 - base2
                    abaixo_val = str(abaixo_num)
            except Exception:
                pass

        registro = {
            "codigo": cod,
            "descricao": descricao,
            "qtd": qtd_raw,
            "qtd_fmt": qtd_fmt,
            "fab_iso": fab_iso,
            "val_iso": val_iso,
            "fab_fmt": fab_fmt,
            "val_fmt": val_fmt,
            "abaixo": abaixo_val,
            "abaixo_num": abaixo_num,
        }
        linhas_calculadas.append(registro)

        if somente_abaixo:
            if abaixo_num is not None and abaixo_num < 0:
                linhas_filtradas.append(registro)
        else:
            linhas_filtradas.append(registro)

    if not linhas_filtradas and linhas_calculadas:
        linhas_filtradas = linhas_calculadas

    png_bytes = renderizar_imagem_tabela_excel(linhas_filtradas)
    png_b64 = base64.b64encode(png_bytes).decode("utf-8")

    xlsx_bytes = gerar_xlsx_recebimento(linhas_calculadas, data_ref_str)
    xlsx_b64 = base64.b64encode(xlsx_bytes).decode("utf-8")

    data_slug = data_ref_str.replace("/", ".")
    return jsonify({
        "ok": True,
        "total_itens": len(linhas_calculadas),
        "total_exibidos": len(linhas_filtradas),
        "arquivo_png": f"VALIDADE_RECEBIMENTO_{data_slug}.png",
        "arquivo_xlsx": f"RECEBIMENTO_{data_slug}.xlsx",
        "png_base64": png_b64,
        "xlsx_base64": xlsx_b64,
    })

@app.route("/baixar/<int:rec_id>/<tipo>")
def baixar_historico(rec_id, tipo):
    col = "arquivo_pdf" if tipo == "pdf" else "arquivo_docx"
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(f"SELECT {col} FROM recebimentos WHERE id = ?", (rec_id,)).fetchone()
    if not row or not row[0]:
        return "Arquivo não disponível", 404
    caminho = os.path.join(OUTPUT_DIR, row[0])
    if not os.path.exists(caminho):
        return "Arquivo não encontrado no disco", 404
    return send_file(caminho, as_attachment=True, download_name=row[0])

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
