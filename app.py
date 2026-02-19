import os
import re
import io
import tempfile
import subprocess
from datetime import datetime
from decimal import Decimal, InvalidOperation

import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Flask, render_template, request, send_file, redirect, url_for, abort, flash

from docx import Document
from docx.shared import Inches

try:
    from num2words import num2words
except Exception:
    num2words = None


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")


# -----------------------------
# DB
# -----------------------------
def db_conn():
    dburl = os.environ.get("DATABASE_URL")
    if not dburl:
        raise RuntimeError("DATABASE_URL não encontrada (adicione o PostgreSQL no Railway).")
    return psycopg2.connect(dburl)


def init_db():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS propostas (
            id SERIAL PRIMARY KEY,
            cliente TEXT NOT NULL,
            cpf TEXT NOT NULL,
            modelo TEXT NOT NULL,
            franquia INTEGER NOT NULL,
            valor NUMERIC(12,2) NOT NULL,
            pdf BYTEA NOT NULL,
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        );
        """)
        conn.commit()


def salvar_proposta(cliente, cpf, modelo, franquia_int, valor_num, pdf_bytes):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            INSERT INTO propostas (cliente, cpf, modelo, franquia, valor, pdf)
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (cliente, cpf, modelo, franquia_int, valor_num, psycopg2.Binary(pdf_bytes)))
        conn.commit()


def listar_propostas(limit=100):
    with db_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, created_at
            FROM propostas
            ORDER BY created_at DESC
            LIMIT %s
        """, (limit,))
        return cur.fetchall()


def limpar_propostas_expiradas(dias=10):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE created_at < NOW() - INTERVAL %s", (f"{dias} days",))
        conn.commit()


def pegar_pdf_proposta(pid: int) -> bytes | None:
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT pdf FROM propostas WHERE id=%s", (pid,))
        row = cur.fetchone()
        if not row:
            return None
        return bytes(row[0])


def excluir_proposta(pid: int):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE id=%s", (pid,))
        conn.commit()


# -----------------------------
# DOCX helpers (substituição confiável)
# -----------------------------
def replace_text_everywhere(doc: Document, mapping: dict[str, str]):
    # método “confiável”: substitui no texto do parágrafo e recria em 1 run
    # (mantém estilo do parágrafo, evita placeholders quebrados em vários runs)
    def _replace_paragraph(p):
        txt = p.text
        changed = False
        for k, v in mapping.items():
            if k in txt:
                txt = txt.replace(k, v)
                changed = True
        if changed:
            for r in p.runs:
                r.text = ""
            if p.runs:
                p.runs[0].text = txt
            else:
                p.add_run(txt)

    for p in doc.paragraphs:
        _replace_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_paragraph(p)


def insert_image_at_placeholder(doc: Document, placeholder: str, image_path: str, width_inches: float = 2.2) -> bool:
    if not image_path or not os.path.exists(image_path):
        return False

    def _handle_paragraph(p):
        if placeholder in p.text:
            # limpa o texto e coloca a imagem no mesmo parágrafo
            for r in p.runs:
                r.text = ""
            run = p.runs[0] if p.runs else p.add_run("")
            run.add_picture(image_path, width=Inches(width_inches))
            return True
        return False

    for p in doc.paragraphs:
        if _handle_paragraph(p):
            return True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if _handle_paragraph(p):
                        return True

    return False


# -----------------------------
# PDF conversion
# -----------------------------
def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nolockcheck", "--convert-to", "pdf",
         "--outdir", out_dir, docx_path],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    return os.path.join(out_dir, base + ".pdf")


# -----------------------------
# Format / parsing
# -----------------------------
MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}


def data_extenso_por_digitos(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^\d]", "", s)
    if len(s) == 8:
        dd = int(s[0:2]); mm = int(s[2:4]); yyyy = int(s[4:8])
    elif len(s) == 6:
        dd = int(s[0:2]); mm = int(s[2:4]); yy = int(s[4:6])
        yyyy = 2000 + yy
    else:
        raise ValueError("Data inválida (use DDMMAAAA, ex: 20022026).")

    datetime(yyyy, mm, dd)
    return f"{dd} de {MESES[mm]} de {yyyy}"


def parse_int(s: str) -> int:
    s = (s or "").strip()
    s = re.sub(r"[^\d]", "", s)
    if not s:
        raise ValueError("Número inválido.")
    return int(s)


def parse_money_to_decimal(s: str) -> Decimal:
    s = (s or "").strip()
    s = s.replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        val = Decimal(s)
    except InvalidOperation:
        raise ValueError("Valor inválido (ex: 150 ou 150,00).")
    return val.quantize(Decimal("0.01"))


def brl_fmt(valor: Decimal) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def numero_por_extenso(n: int) -> str:
    if not num2words:
        return str(n)
    return num2words(n, lang="pt_BR")


def dinheiro_por_extenso(valor: Decimal) -> str:
    reais = int(valor)
    if not num2words:
        return "reais"
    if reais == 1:
        return "um real"
    return f"{num2words(reais, lang='pt_BR')} reais"


# -----------------------------
# Routes
# -----------------------------
@app.before_request
def _startup():
    try:
        init_db()
    except Exception:
        pass


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "GET":
        return render_template("proposta.html")

    prefill = dict(request.form)
    try:
        cliente = (request.form.get("cliente") or "").strip()
        cpf = (request.form.get("cpf") or "").strip()
        modelo = (request.form.get("modelo") or "").strip()
        franquia_int = parse_int(request.form.get("franquia") or "")
        valor_num = parse_money_to_decimal(request.form.get("valor") or "")

        if not cliente or not cpf or not modelo:
            raise ValueError("Preencha cliente, CPF/CNPJ e modelo.")

        valor_ext = dinheiro_por_extenso(valor_num)
        valor_fmt = f"R$ {brl_fmt(valor_num)} ({valor_ext})"

        template_path = os.path.join(os.path.dirname(__file__), "template.docx")
        if not os.path.exists(template_path):
            raise RuntimeError("template.docx não encontrado no projeto.")

        doc = Document(template_path)

        mapping = {
            "{{ CLIENTE }}": cliente,
            "{{ CPF }}": cpf,
            "{{ MODELO }}": modelo,
            "{{ FRANQUIA }}": str(franquia_int),
            "{{ VALOR }}": valor_fmt,
            "{{ DATA }}": datetime.now().strftime("%d/%m/%Y"),
        }
        replace_text_everywhere(doc, mapping)

        # IMAGEM: upload do formulário (opcional)
        image_file = request.files.get("imagem")
        image_path = None
        with tempfile.TemporaryDirectory() as tmp:
            if image_file and image_file.filename:
                image_path = os.path.join(tmp, "img_upload")
                image_file.save(image_path)
                insert_image_at_placeholder(doc, "{{ IMAGEM }}", image_path, width_inches=2.2)
            else:
                # fallback: se existir static/logo.png
                static_logo = os.path.join(os.path.dirname(__file__), "static", "logo.png")
                if os.path.exists(static_logo):
                    insert_image_at_placeholder(doc, "{{ IMAGEM }}", static_logo, width_inches=2.2)

            docx_saida = os.path.join(tmp, "proposta_gerada.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        salvar_proposta(cliente, cpf, modelo, franquia_int, valor_num, pdf_bytes)

        nome_arquivo = f"Proposta ({cliente}).pdf"
        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf", as_attachment=True, download_name=nome_arquivo)

    except Exception as e:
        flash(str(e))
        return render_template("proposta.html", prefill=prefill)


@app.route("/recentes")
def recentes():
    try:
        limpar_propostas_expiradas()
        props = listar_propostas()
        return render_template("recentes.html", propostas=props)
    except Exception as e:
        return f"Erro em recentes: {e}", 500


@app.route("/proposta/<int:pid>/ver")
def proposta_ver(pid: int):
    pdf = pegar_pdf_proposta(pid)
    if not pdf:
        abort(404)
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=False)


@app.route("/proposta/<int:pid>/baixar")
def proposta_baixar(pid: int):
    pdf = pegar_pdf_proposta(pid)
    if not pdf:
        abort(404)
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name=f"Proposta ({pid}).pdf")


@app.route("/proposta/<int:pid>/excluir", methods=["POST"])
def proposta_excluir(pid: int):
    excluir_proposta(pid)
    return redirect(url_for("recentes"))


@app.route("/contrato", methods=["GET", "POST"])
def contrato():
    if request.method == "GET":
        return render_template("contrato.html", prefill={})

    prefill = dict(request.form)
    try:
        denominacao = (request.form.get("denominacao") or "").strip()
        cpf = (request.form.get("cpf") or "").strip()
        endereco = (request.form.get("endereco") or "").strip()
        telefone = (request.form.get("telefone") or "").strip()
        email = (request.form.get("email") or "").strip()
        equipamento = (request.form.get("equipamento") or "").strip()
        acessorios = (request.form.get("acessorios") or "").strip()

        data_inicio_ext = data_extenso_por_digitos(request.form.get("data_inicio") or "")
        data_termino_ext = data_extenso_por_digitos(request.form.get("data_termino") or "")

        franquia_int = parse_int(request.form.get("franquia") or "")
        valor_num = parse_money_to_decimal(request.form.get("valor") or "")

        franquia_ext = numero_por_extenso(franquia_int)
        valor_ext = dinheiro_por_extenso(valor_num)
        valor_fmt = f"R$ {brl_fmt(valor_num)} ({valor_ext})"

        data_assinatura_ext = data_extenso_por_digitos(datetime.now().strftime("%d%m%Y"))

        template_path = os.path.join(os.path.dirname(__file__), "contrato_template.docx")
        if not os.path.exists(template_path):
            raise RuntimeError("contrato_template.docx não encontrado no projeto.")

        doc = Document(template_path)

        mapping = {
            "{{ DENOMINACAO }}": denominacao,
            "{{ CPF }}": cpf,
            "{{ ENDERECO }}": endereco,
            "{{ TELEFONE }}": telefone,
            "{{ EMAIL }}": email,
            "{{ EQUIPAMENTO }}": equipamento,
            "{{ ACESSORIOS }}": acessorios,
            "{{ DATA_INICIO }}": data_inicio_ext,
            "{{ DATA_TERMINO }}": data_termino_ext,
            "{{ FRANQUIA }}": f"{franquia_int} ({franquia_ext})",
            "{{ VALOR }}": valor_fmt,
            "{{ DATA_ASSINATURA }}": data_assinatura_ext,
            "{{ ASSINATURA }}": denominacao,
        }
        replace_text_everywhere(doc, mapping)

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "contrato_gerado.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        nome_arquivo = f"Contrato ({denominacao}).pdf"
        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf", as_attachment=True, download_name=nome_arquivo)

    except Exception as e:
        flash(str(e))
        return render_template("contrato.html", prefill=prefill)


if __name__ == "__main__":
    app.run(debug=True)
