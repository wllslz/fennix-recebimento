import os
import io
import base64
import subprocess
import shutil
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, send_file, jsonify
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image, ImageOps

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "template_recebimento_fennix.docx")

IS_VERCEL = bool(os.environ.get("VERCEL"))
if IS_VERCEL:
    OUTPUT_DIR = "/tmp/gerados"
    DB_PATH = "/tmp/historico_recebimento.db"
else:
    OUTPUT_DIR = os.path.join(BASE_DIR, "gerados")
    DB_PATH = os.path.join(BASE_DIR, "historico_recebimento.db")

os.makedirs(OUTPUT_DIR, exist_ok=True)

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
    """Converte o DOCX para PDF quando Word (Windows) ou LibreOffice está disponível."""
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
