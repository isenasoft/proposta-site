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

try:
    from num2words import num2words
except Exception:
    num2words = None  # se não tiver, ainda roda sem extenso


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
    # Cria a tabela se não existir (não quebra se já existir)
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


def listar_propostas(limit=50):
    with db_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, created_at
            FROM propostas
            ORDER BY created_at DESC
            LIMIT %s
        """, (limit,))
        return cur.fetchall()


def limpar_propostas_expiradas(dias=10):
    # apaga antigas automaticamente
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
# DOCX helpers (mantém estilo)
# -----------------------------
def replace_in_paragraph(paragraph, mapping: dict[str, str]):
    # substitui preservando runs
    for key, val in mapping.items():
        if key in paragraph.text:
            # junta os runs, substitui, e redistribui (simples e funciona bem na maioria dos templates)
            full = "".join(run.text for run in paragraph.runs)
            if key not in full:
                continue
            full = full.replace(key, val)
            # limpa e coloca no primeiro run
            for run in paragraph.runs:
                run.text = ""
            paragraph.runs[0].text = full


def replace_in_doc(doc: Document, mapping: dict[str, str]):
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, mapping)


# -----------------------------
# PDF conversion
# -----------------------------
def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    # LibreOffice headless
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
    """
    Aceita: 20022026, 20/02/2026, 20-02-2026, 20/02/26
    Retorna: 20 de fevereiro de 2026
    """
    s = (s or "").strip()
    s = re.sub(r"[^\d]", "", s)  # deixa só números
    if len(s) == 8:
        dd = int(s[0:2]); mm = int(s[2:4]); yyyy = int(s[4:8])
    elif len(s) == 6:
        dd = int(s[0:2]); mm = int(s[2:4]); yy = int(s[4:6])
        yyyy = 2000 + yy
    else:
        raise ValueError("Data inválida (use DDMMAAAA, ex: 20022026).")

    datetime(yyyy, mm, dd)  # valida
    return f"{dd} de {MESES[mm]} de {yyyy}"


def parse_int(s: str) -> int:
    s = (s or "").strip()
    s = re.sub(r"[^\d]", "", s)
    if not s:
        raise ValueError("Número inválido.")
    return int(s)


def parse_money_to_decimal(s: str) -> Decimal:
    """
    Aceita: "150", "150,00", "150.00", "R$ 150,00"
    Retorna Decimal('150.00')
    """
    s = (s or "").strip()
    s = s.replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")  # BR -> decimal
    try:
        val = Decimal(s)
    except InvalidOperation:
        raise ValueError("Valor inválido (ex: 150 ou 150,00).")
    return val.quantize(Decimal("0.01"))


def brl_fmt(valor: Decimal) -> str:
    # 150.00 -> "150,00"
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def numero_por_extenso(n: int) -> str:
    if not num2words:
        return str(n)
    return num2words(n, lang="pt_BR")


def dinheiro_por_extenso(valor: Decimal) -> str:
    # "150.00" -> "cento e cinquenta reais"
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
    # garante DB
    try:
        init_db()
    except Exception:
        # não derruba tudo só por isso
        pass


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "GET":
        return render_template("proposta.html")

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
            # fallback caso use outro nome
            template_path = os.path.join(os.path.dirname(__file__), "contrato_template.docx")

        doc = Document(template_path)

        mapping = {
            "{{ CLIENTE }}": cliente,
            "{{ CPF }}": cpf,
            "{{ MODELO }}": modelo,
            "{{ FRANQUIA }}": str(franquia_int),
            "{{ VALOR }}": valor_fmt,
            "{{ DATA }}": datetime.now().strftime("%d/%m/%Y"),
        }
        replace_in_doc(doc, mapping)

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "proposta_gerada.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        # SALVA NO BANCO: valor NUMÉRICO (não o texto formatado!)
        salvar_proposta(cliente, cpf, modelo, franquia_int, valor_num, pdf_bytes)

        nome_arquivo = f"Proposta ({cliente}).pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=nome_arquivo
        )

    except Exception as e:
        # não “apaga” o formulário: volta com erro
        flash(str(e))
        return render_template("proposta.html")


@app.route("/recentes")
def recentes():
    try:
        limpar_propostas_expiradas()
        props = listar_propostas()
        return render_template("recentes.html", propostas=props)
    except Exception as e:
        return f"Erro em recentes: {e}", 500


@app.route("/proposta/<int:pid>")
def proposta_view(pid: int):
    pdf = pegar_pdf_proposta(pid)
    if not pdf:
        abort(404)
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=False)


@app.route("/proposta/<int:pid>/download")
def proposta_download(pid: int):
    pdf = pegar_pdf_proposta(pid)
    if not pdf:
        abort(404)
    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True, download_name=f"Proposta ({pid}).pdf")


@app.route("/proposta/<int:pid>/excluir", methods=["POST"])
def proposta_excluir(pid: int):
    try:
        excluir_proposta(pid)
        return redirect(url_for("recentes"))
    except Exception as e:
        return f"Erro ao excluir: {e}", 500


@app.route("/contrato", methods=["GET", "POST"])
def contrato():
    if request.method == "GET":
        # garante prefill SEMPRE
        return render_template("contrato.html", prefill={})

    prefill = dict(request.form)  # para não limpar se der erro
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
            "{{ ASSINATURA }}": denominacao,  # repete nome na assinatura
        }
        replace_in_doc(doc, mapping)

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "contrato_gerado.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        nome_arquivo = f"Contrato ({denominacao}).pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=nome_arquivo
        )

    except Exception as e:
        flash(str(e))
        return render_template("contrato.html", prefill=prefill)


if __name__ == "__main__":
    app.run(debug=True)
