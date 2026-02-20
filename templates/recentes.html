import os
import re
import io
import uuid
import tempfile
import subprocess
from decimal import Decimal, InvalidOperation
from datetime import datetime, date
from zoneinfo import ZoneInfo

import psycopg2
from flask import Flask, request, render_template, send_file, redirect, url_for, flash

from docx import Document
from docx.shared import Inches

# ----------------------------
# Config
# ----------------------------
TZ_BR = ZoneInfo("America/Sao_Paulo")
UPLOAD_MAX_MB = 8

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

app.config["MAX_CONTENT_LENGTH"] = UPLOAD_MAX_MB * 1024 * 1024

# ----------------------------
# Utils: DB
# ----------------------------
def db_conn():
    dsn = os.environ.get("DATABASE_URL")
    if not dsn:
        raise RuntimeError("DATABASE_URL não encontrada (adicione o PostgreSQL no Railway).")
    return psycopg2.connect(dsn)

def ensure_schema():
    """
    Garante tabela 'propostas' com colunas esperadas:
    id (uuid), cliente, cpf, modelo, franquia (int), valor (numeric), pdf (bytea), created_at (timestamp)
    """
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS propostas (
                id UUID PRIMARY KEY,
                cliente TEXT NOT NULL,
                cpf TEXT NOT NULL,
                modelo TEXT NOT NULL,
                franquia INTEGER NOT NULL,
                valor NUMERIC(12,2) NOT NULL,
                pdf BYTEA NOT NULL,
                created_at TIMESTAMP NOT NULL DEFAULT NOW()
            );
        """)

        # Caso sua tabela seja antiga e falte alguma coluna, tenta adicionar sem quebrar
        cur.execute("""
            DO $$
            BEGIN
                IF NOT EXISTS (
                    SELECT 1 FROM information_schema.columns
                    WHERE table_name='propostas' AND column_name='created_at'
                ) THEN
                    ALTER TABLE propostas ADD COLUMN created_at TIMESTAMP NOT NULL DEFAULT NOW();
                END IF;

                IF NOT EXISTS (
                    SELECT 1 FROM information_schema.columns
                    WHERE table_name='propostas' AND column_name='pdf'
                ) THEN
                    ALTER TABLE propostas ADD COLUMN pdf BYTEA;
                END IF;
            END $$;
        """)

def salvar_proposta(cliente, cpf, modelo, franquia_int, valor_decimal, pdf_bytes):
    ensure_schema()
    pid = uuid.uuid4()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            INSERT INTO propostas (id, cliente, cpf, modelo, franquia, valor, pdf, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, NOW());
        """, (str(pid), cliente, cpf, modelo, int(franquia_int), valor_decimal, psycopg2.Binary(pdf_bytes)))
    return str(pid)

def listar_propostas():
    ensure_schema()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, created_at
            FROM propostas
            ORDER BY created_at DESC
            LIMIT 50;
        """)
        rows = cur.fetchall()

    out = []
    for r in rows:
        out.append({
            "id": str(r[0]),
            "cliente": r[1],
            "cpf": r[2],
            "modelo": r[3],
            "franquia": int(r[4]),
            "valor": Decimal(str(r[5])),
            "created_at": r[6],
        })
    return out

def obter_pdf_proposta(pid):
    ensure_schema()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT pdf FROM propostas WHERE id=%s;", (pid,))
        row = cur.fetchone()
        return row[0] if row else None

def deletar_proposta(pid):
    ensure_schema()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE id=%s;", (pid,))

def limpar_propostas_expiradas(dias=10):
    ensure_schema()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute(f"DELETE FROM propostas WHERE created_at < NOW() - INTERVAL '{int(dias)} days';")

# ----------------------------
# Utils: texto / datas / números
# ----------------------------
MESES = [
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
]

def now_br():
    return datetime.now(TZ_BR)

def data_extenso(dt: date):
    return f"{dt.day} de {MESES[dt.month-1]} de {dt.year}"

def data_extenso_por_digitos(s: str):
    """
    Aceita:
    - "20022026"
    - "20/02/2026"
    - "20-02-2026"
    - "20/02/26"
    """
    s = (s or "").strip()
    if not s:
        raise ValueError("Data vazia")

    digits = re.sub(r"\D", "", s)  # só números

    if len(digits) == 8:
        dd = int(digits[0:2])
        mm = int(digits[2:4])
        yyyy = int(digits[4:8])
    elif len(digits) == 6:
        dd = int(digits[0:2])
        mm = int(digits[2:4])
        yy = int(digits[4:6])
        yyyy = 2000 + yy
    else:
        raise ValueError("Data inválida")

    dt = date(yyyy, mm, dd)
    return data_extenso(dt)

def format_brl(valor: Decimal):
    # 200.00 -> "200,00"
    s = f"{valor:.2f}"
    return s.replace(".", ",")

UNIDADES = ["zero","um","dois","três","quatro","cinco","seis","sete","oito","nove"]
DEZ_A_DEZENOVE = ["dez","onze","doze","treze","quatorze","quinze","dezesseis","dezessete","dezoito","dezenove"]
DEZENAS = ["","dez","vinte","trinta","quarenta","cinquenta","sessenta","setenta","oitenta","noventa"]
CENTENAS = ["","cem","cento","duzentos","trezentos","quatrocentos","quinhentos","seiscentos","setecentos","oitocentos","novecentos"]

def numero_por_extenso_ate_999(n: int):
    if n < 0 or n > 999:
        raise ValueError("fora do range 0-999")
    if n < 10:
        return UNIDADES[n]
    if 10 <= n < 20:
        return DEZ_A_DEZENOVE[n-10]
    if n < 100:
        d = n // 10
        u = n % 10
        if u == 0:
            return DEZENAS[d]
        return f"{DEZENAS[d]} e {UNIDADES[u]}"
    # 100-999
    if n == 100:
        return "cem"
    c = n // 100
    resto = n % 100
    if resto == 0:
        return CENTENAS[c]
    return f"{CENTENAS[c]} e {numero_por_extenso_ate_999(resto)}"

def numero_por_extenso(n: int):
    # suficiente pra franquia e valores mensais do seu caso
    if n < 0:
        return "zero"
    if n <= 999:
        return numero_por_extenso_ate_999(n)
    if n <= 999999:
        mil = n // 1000
        resto = n % 1000
        mil_txt = "mil" if mil == 1 else f"{numero_por_extenso_ate_999(mil)} mil"
        if resto == 0:
            return mil_txt
        # regra simples com "e" para português
        conj = " e " if resto < 100 else " "
        return f"{mil_txt}{conj}{numero_por_extenso_ate_999(resto)}"
    return str(n)

def parse_valor_decimal(s: str) -> Decimal:
    s = (s or "").strip()
    s = s.replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        v = Decimal(s)
        return v.quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        raise ValueError("Valor inválido")

# ----------------------------
# Utils: DOCX placeholder replace (body + tables + header/footer)
# ----------------------------
def iter_paragraphs(container):
    """
    Retorna todos os parágrafos de:
    - Document / _Cell / Header / Footer
    incluindo os parágrafos dentro de tabelas.
    """
    # parágrafos diretos
    for p in getattr(container, "paragraphs", []):
        yield p
    # tabelas
    for t in getattr(container, "tables", []):
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                # tabelas aninhadas
                for p in iter_paragraphs(cell):
                    yield p

def iter_all_paragraphs(doc: Document):
    # body
    for p in iter_paragraphs(doc):
        yield p
    # header/footer
    for sec in doc.sections:
        for p in iter_paragraphs(sec.header):
            yield p
        for p in iter_paragraphs(sec.footer):
            yield p

def replace_placeholders(doc: Document, mapping: dict):
    """
    Substitui placeholders como {{ NOME }} preservando o estilo principal do parágrafo.
    Funciona mesmo quando o Word quebra em vários "runs".
    """
    for p in iter_all_paragraphs(doc):
        full = "".join(run.text for run in p.runs)
        if "{{" not in full:
            continue

        new_full = full
        for k, v in mapping.items():
            new_full = new_full.replace(k, str(v))

        if new_full == full:
            continue

        # mantém estilo do primeiro run
        if p.runs:
            p.runs[0].text = new_full
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(new_full)

def insert_image_at_placeholder(doc: Document, placeholder: str, image_path: str, width_inches: float = 4.8):
    """
    Insere imagem onde encontrar o placeholder (em body/tables/header/footer).
    width_inches controlado para evitar virar 2 páginas.
    """
    if not image_path or not os.path.exists(image_path):
        return

    for p in iter_all_paragraphs(doc):
        full = "".join(run.text for run in p.runs)
        if placeholder not in full:
            continue

        # remove placeholder do texto
        new_full = full.replace(placeholder, "").strip()
        if p.runs:
            p.runs[0].text = new_full
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(new_full)

        # insere imagem no mesmo parágrafo
        run = p.add_run()
        run.add_picture(image_path, width=Inches(width_inches))

        # reduz espaçamento (ajuda a não estourar página)
        try:
            p.paragraph_format.space_before = 0
            p.paragraph_format.space_after = 0
        except Exception:
            pass

# ----------------------------
# Utils: DOCX -> PDF (LibreOffice)
# ----------------------------
def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nolockcheck",
         "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    return pdf_path

# ----------------------------
# Rotas
# ----------------------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "GET":
        return render_template("proposta.html")

    cliente = (request.form.get("cliente") or "").strip()
    cpf = (request.form.get("cpf") or "").strip()
    modelo = (request.form.get("modelo") or "").strip()
    franquia = (request.form.get("franquia") or "").strip()
    valor_in = (request.form.get("valor") or "").strip()

    if not all([cliente, cpf, modelo, franquia, valor_in]):
        flash("Preencha todos os campos.")
        return redirect(url_for("proposta"))

    try:
        franquia_int = int(re.sub(r"\D", "", franquia))
    except Exception:
        flash("Franquia inválida.")
        return redirect(url_for("proposta"))

    try:
        valor_decimal = parse_valor_decimal(valor_in)
    except Exception:
        flash("Valor inválido.")
        return redirect(url_for("proposta"))

    valor_fmt = f"R$ {format_brl(valor_decimal)} ({numero_por_extenso(int(valor_decimal))} reais)"

    # imagem opcional
    img_file = request.files.get("imagem")
    img_tmp_path = None
    if img_file and img_file.filename:
        ext = os.path.splitext(img_file.filename)[1].lower()
        if ext not in [".png", ".jpg", ".jpeg", ".webp"]:
            flash("Imagem inválida (use PNG/JPG/WEBP).")
            return redirect(url_for("proposta"))
        fd, img_tmp_path = tempfile.mkstemp(suffix=ext)
        os.close(fd)
        img_file.save(img_tmp_path)

    # template proposta
    template_path = os.path.join(os.getcwd(), "template.docx")
    if not os.path.exists(template_path):
        # caso seu template tenha outro nome, ajuste aqui
        template_path = os.path.join(os.getcwd(), "proposta_template.docx")

    if not os.path.exists(template_path):
        flash("Template da proposta não encontrado (template.docx).")
        return redirect(url_for("proposta"))

    with tempfile.TemporaryDirectory() as tmp:
        doc = Document(template_path)

        hoje = now_br().date()
        mapping = {
            "{{ CLIENTE }}": cliente,
            "{{ CPF }}": cpf,
            "{{ MODELO }}": modelo,
            "{{ FRANQUIA }}": str(franquia_int),
            "{{ FRANQUIA_EXTENSO }}": numero_por_extenso(franquia_int),
            "{{ VALOR }}": valor_fmt,
            "{{ DATA }}": data_extenso(hoje),  # data do topo (se estiver no header)
        }

        replace_placeholders(doc, mapping)

        # imagem (se o template tiver {{ IMAGEM }})
        if img_tmp_path:
            insert_image_at_placeholder(doc, "{{ IMAGEM }}", img_tmp_path, width_inches=4.8)

        docx_out = os.path.join(tmp, "proposta_gerada.docx")
        doc.save(docx_out)

        pdf_out = docx_para_pdf(docx_out, tmp)

        with open(pdf_out, "rb") as f:
            pdf_bytes = f.read()

    # salva no banco com valor NUMÉRICO (sem texto)
    pid = salvar_proposta(cliente, cpf, modelo, franquia_int, valor_decimal, pdf_bytes)

    filename = f"Proposta {cliente}.pdf"
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )

@app.route("/contrato", methods=["GET", "POST"])
def contrato():
    prefill = {
        "denominacao": request.args.get("cliente", ""),
        "cpf": request.args.get("cpf", ""),
        "equipamento": request.args.get("modelo", ""),
        "franquia": request.args.get("franquia", ""),
        "valor": request.args.get("valor", ""),
    }

    if request.method == "GET":
        return render_template("contrato.html", prefill=prefill)

    denominacao = (request.form.get("denominacao") or "").strip()
    cpf = (request.form.get("cpf") or "").strip()
    endereco = (request.form.get("endereco") or "").strip()
    telefone = (request.form.get("telefone") or "").strip()
    email = (request.form.get("email") or "").strip()
    equipamento = (request.form.get("equipamento") or "").strip()
    acessorios = (request.form.get("acessorios") or "").strip()
    data_inicio_in = (request.form.get("data_inicio") or "").strip()
    data_termino_in = (request.form.get("data_termino") or "").strip()
    franquia_in = (request.form.get("franquia") or "").strip()
    valor_in = (request.form.get("valor") or "").strip()

    if not all([denominacao, cpf, endereco, telefone, email, equipamento, data_inicio_in, data_termino_in, franquia_in, valor_in]):
        flash("Preencha todos os campos do contrato.")
        return redirect(url_for("contrato"))

    try:
        data_inicio = data_extenso_por_digitos(data_inicio_in)
        data_termino = data_extenso_por_digitos(data_termino_in)
    except Exception:
        flash("Data inválida. Digite só números (ex: 20022026) ou com /.")
        return redirect(url_for("contrato"))

    try:
        franquia_int = int(re.sub(r"\D", "", franquia_in))
    except Exception:
        flash("Franquia inválida.")
        return redirect(url_for("contrato"))

    try:
        valor_decimal = parse_valor_decimal(valor_in)
    except Exception:
        flash("Valor inválido.")
        return redirect(url_for("contrato"))

    franquia_formatada = f"{franquia_int}"
    franquia_extenso = numero_por_extenso(franquia_int)

    valor_mensal_formatado = format_brl(valor_decimal)
    valor_mensal_extenso = f"{numero_por_extenso(int(valor_decimal))} reais"

    data_assinatura = data_extenso(now_br().date())

    template_path = os.path.join(os.getcwd(), "contrato_template.docx")
    if not os.path.exists(template_path):
        flash("Template do contrato não encontrado (contrato_template.docx).")
        return redirect(url_for("contrato"))

    with tempfile.TemporaryDirectory() as tmp:
        doc = Document(template_path)
        mapping = {
            "{{ DENOMINACAO }}": denominacao,
            "{{ CPF }}": cpf,
            "{{ ENDERECO }}": endereco,
            "{{ TELEFONE }}": telefone,
            "{{ EMAIL }}": email,
            "{{ EQUIPAMENTO }}": equipamento,
            "{{ ACESSORIOS }}": acessorios,

            "{{ DATA_INICIO }}": data_inicio,
            "{{ DATA_TERMINO }}": data_termino,

            "{{ FRANQUIA_FORMATADA }}": franquia_formatada,
            "{{ FRANQUIA_EXTENSO }}": franquia_extenso,

            "{{ VALOR_MENSAL_FORMATADO }}": valor_mensal_formatado,
            "{{ VALOR_MENSAL_EXTENSO }}": valor_mensal_extenso,

            "{{ DATA_ASSINATURA }}": data_assinatura,
        }

        replace_placeholders(doc, mapping)

        docx_out = os.path.join(tmp, "contrato_gerado.docx")
        doc.save(docx_out)

        pdf_out = docx_para_pdf(docx_out, tmp)
        with open(pdf_out, "rb") as f:
            pdf_bytes = f.read()

    filename = f"Contrato {denominacao}.pdf"
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )

@app.route("/recentes")
def recentes():
    limpar_propostas_expiradas(dias=10)
    propostas = listar_propostas()

    # formata valor pra tela
    for p in propostas:
        p["valor_fmt"] = f"R$ {format_brl(p['valor'])}"
        try:
            p["created_fmt"] = p["created_at"].astimezone(TZ_BR).strftime("%d/%m/%Y %H:%M")
        except Exception:
            p["created_fmt"] = str(p["created_at"])

    return render_template("recentes.html", propostas=propostas)

@app.route("/recentes/<pid>/pdf")
def recentes_pdf(pid):
    pdf = obter_pdf_proposta(pid)
    if not pdf:
        return "PDF não encontrado", 404

    return send_file(
        io.BytesIO(pdf),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"Proposta {pid}.pdf"
    )

@app.route("/recentes/<pid>/delete", methods=["POST"])
def recentes_delete(pid):
    deletar_proposta(pid)
    return redirect(url_for("recentes"))

@app.route("/recentes/<pid>/contrato")
def recentes_contrato(pid):
    # puxa dados da proposta pra pré-preencher o contrato
    ensure_schema()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT cliente, cpf, modelo, franquia, valor
            FROM propostas WHERE id=%s
        """, (pid,))
        row = cur.fetchone()

    if not row:
        return "Proposta não encontrada", 404

    cliente, cpf, modelo, franquia, valor = row
    return redirect(url_for(
        "contrato",
        cliente=cliente,
        cpf=cpf,
        modelo=modelo,
        franquia=str(franquia),
        valor=str(valor)
    ))

# ----------------------------
# Main
# ----------------------------
if __name__ == "__main__":
    app.run(debug=True)
