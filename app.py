import os
import re
import tempfile
import subprocess
from io import BytesIO
from datetime import datetime

import psycopg2
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from num2words import num2words


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev")

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# -----------------------
# Helpers (DB)
# -----------------------
def db_conn():
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL não encontrada (adicione o PostgreSQL no Railway e conecte a variável).")
    return psycopg2.connect(db_url)


def init_db():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS propostas (
                id SERIAL PRIMARY KEY,
                cliente TEXT NOT NULL,
                cpf TEXT NOT NULL,
                modelo TEXT NOT NULL,
                franquia INTEGER NOT NULL,
                valor TEXT NOT NULL,
                pdf BYTEA NOT NULL,
                criado_em TIMESTAMP NOT NULL DEFAULT NOW()
            );
        """)
        conn.commit()


def limpar_propostas_expiradas():
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE created_at < NOW() - INTERVAL '10 days';")
        conn.commit()


def listar_propostas():
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, criado_em
            FROM propostas
            ORDER BY criado_em DESC;
        """)
        rows = cur.fetchall()
    props = []
    for r in rows:
        props.append({
            "id": r[0],
            "cliente": r[1],
            "cpf": r[2],
            "modelo": r[3],
            "franquia": r[4],
            "valor": r[5],
            "criado_em": r[6],
        })
    return props


def salvar_proposta(cliente, cpf, modelo, franquia, valor_str, pdf_bytes):
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            INSERT INTO propostas (cliente, cpf, modelo, franquia, valor, pdf)
            VALUES (%s,%s,%s,%s,%s,%s)
            RETURNING id;
        """, (cliente, cpf, modelo, int(franquia), valor_str, psycopg2.Binary(pdf_bytes)))
        pid = cur.fetchone()[0]
        conn.commit()
    return pid


def carregar_pdf_proposta(pid):
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT cliente, pdf FROM propostas WHERE id=%s;", (pid,))
        row = cur.fetchone()
    return row  # (cliente, pdf) or None


def carregar_dados_proposta(pid):
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
            SELECT cliente, cpf, modelo, franquia, valor
            FROM propostas
            WHERE id=%s;
        """, (pid,))
        row = cur.fetchone()
    if not row:
        return None
    return {
        "cliente": row[0],
        "cpf": row[1],
        "modelo": row[2],
        "franquia": row[3],
        "valor": row[4],
    }


def excluir_proposta(pid):
    init_db()
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE id=%s;", (pid,))
        conn.commit()


# -----------------------
# Helpers (format)
# -----------------------
def data_formatada_extenso(dt: datetime) -> str:
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
        7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"


def parse_data_usuario(valor: str) -> datetime:
    """
    Aceita: 20022026, 20/02/2026, 20/02/26, 20-02-2026, etc.
    """
    s = re.sub(r"\D", "", (valor or "").strip())
    if len(s) == 8:
        dd = int(s[0:2]); mm = int(s[2:4]); yyyy = int(s[4:8])
    elif len(s) == 6:  # ddmmyy -> assume 20yy
        dd = int(s[0:2]); mm = int(s[2:4]); yy = int(s[4:6])
        yyyy = 2000 + yy
    else:
        raise ValueError("Data inválida (use DDMMAAAA ou DD/MM/AAAA ou DD/MM/AA)")
    return datetime(yyyy, mm, dd)


def franquia_com_extenso(n: str) -> str:
    n_int = int(re.sub(r"\D", "", n))
    ext = num2words(n_int, lang="pt_BR")
    return f"{n_int} ({ext})"


def formatar_valor_com_reais(valor: str) -> str:
    v = float(str(valor).replace(".", "").replace(",", "."))
    extenso = num2words(v, lang="pt_BR")
    dinheiro = f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    # garante "reais" no final
    return f"{dinheiro} ({extenso} reais)"


def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nolockcheck", "--convert-to", "pdf",
         "--outdir", out_dir, docx_path],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    return os.path.join(out_dir, f"{base}.pdf")


# -----------------------
# Routes
# -----------------------
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
    valor = (request.form.get("valor") or "").strip()

    imagem = request.files.get("imagem")

    if not all([cliente, cpf, modelo, franquia, valor]):
        flash("Preencha todos os campos.")
        return redirect(url_for("proposta"))

    data_hoje = data_formatada_extenso(datetime.now())
    valor_fmt = formatar_valor_com_reais(valor)

    tpl_path = "template.docx"
    if not os.path.exists(tpl_path):
        return "template.docx não encontrado no projeto.", 500

    doc = DocxTemplate(tpl_path)

    ctx = {
        "DATA": data_hoje,
        "CLIENTE": cliente,
        "CPF": cpf,
        "MODELO": modelo,
        "FRANQUIA": franquia,
        "VALOR": valor_fmt,
    }

    # imagem opcional
    if imagem and imagem.filename:
        img_path = os.path.join(UPLOAD_FOLDER, imagem.filename)
        imagem.save(img_path)
        ctx["IMAGEM"] = InlineImage(doc, img_path, width=Mm(70))

    doc.render(ctx)

    with tempfile.TemporaryDirectory() as tmp:
        docx_out = os.path.join(tmp, "proposta_gerada.docx")
        doc.save(docx_out)

        try:
            pdf_out = docx_para_pdf(docx_out, tmp)
        except FileNotFoundError:
            return "LibreOffice não encontrado (soffice). Verifique o deploy.", 500

        with open(pdf_out, "rb") as f:
            pdf_bytes = f.read()

    # salva no banco para Recentes
    salvar_proposta(cliente, cpf, modelo, franquia, valor_fmt, pdf_bytes)

    bio = BytesIO(pdf_bytes)
    bio.seek(0)
    return send_file(bio, mimetype="application/pdf", as_attachment=True,
                     download_name=f"Proposta {cliente}.pdf")


@app.route("/recentes")
def recentes():
    limpar_propostas_expiradas()
    propostas = listar_propostas()
    return render_template("recentes.html", propostas=propostas)


@app.route("/proposta_view/<int:pid>")
def proposta_view(pid):
    row = carregar_pdf_proposta(pid)
    if not row:
        return "Proposta não encontrada (talvez expirou).", 404

    _, pdf = row
    pdf_bytes = pdf.tobytes() if hasattr(pdf, "tobytes") else bytes(pdf)

    bio = BytesIO(pdf_bytes)
    bio.seek(0)
    # abre no navegador (sem baixar)
    return send_file(bio, mimetype="application/pdf", as_attachment=False)


@app.route("/proposta_download/<int:pid>")
def proposta_download(pid):
    row = carregar_pdf_proposta(pid)
    if not row:
        return "Proposta não encontrada (talvez expirou).", 404

    cliente, pdf = row
    pdf_bytes = pdf.tobytes() if hasattr(pdf, "tobytes") else bytes(pdf)

    bio = BytesIO(pdf_bytes)
    bio.seek(0)
    return send_file(bio, mimetype="application/pdf", as_attachment=True,
                     download_name=f"Proposta {cliente}.pdf")


@app.route("/proposta_delete/<int:pid>", methods=["POST"])
def proposta_delete(pid):
    excluir_proposta(pid)
    return redirect(url_for("recentes"))


@app.route("/contrato", methods=["GET", "POST"])
def contrato():
    # sempre passa prefill pro HTML (corrige o erro 'prefill is undefined')
    prefill = {}

    # se vier de uma proposta recente (ex: /contrato?proposta_id=3)
    proposta_id = request.args.get("proposta_id")
    if proposta_id and proposta_id.isdigit():
        dados = carregar_dados_proposta(int(proposta_id))
        if dados:
            prefill = {
                "denominacao": dados["cliente"],
                "cpf": dados["cpf"],
                "equipamento": dados["modelo"],
                "franquia": str(dados["franquia"]),
                "valor": re.sub(r"[^\d,\.]", "", dados["valor"]) or "",
            }

    if request.method == "GET":
        return render_template("contrato.html", prefill=prefill)

    # POST
    denominacao = (request.form.get("denominacao") or "").strip()
    cpf = (request.form.get("cpf") or "").strip()
    endereco = (request.form.get("endereco") or "").strip()
    telefone = (request.form.get("telefone") or "").strip()
    email = (request.form.get("email") or "").strip()
    equipamento = (request.form.get("equipamento") or "").strip()
    acessorios = (request.form.get("acessorios") or "").strip()

    data_inicio_raw = (request.form.get("data_inicio") or "").strip()
    data_termino_raw = (request.form.get("data_termino") or "").strip()

    franquia_raw = (request.form.get("franquia") or "").strip()
    valor_raw = (request.form.get("valor") or "").strip()

    if not all([denominacao, cpf, endereco, telefone, email, equipamento, acessorios,
                data_inicio_raw, data_termino_raw, franquia_raw, valor_raw]):
        flash("Preencha todos os campos do contrato.")
        return redirect(url_for("contrato"))

    try:
        dt_ini = parse_data_usuario(data_inicio_raw)
        dt_fim = parse_data_usuario(data_termino_raw)
    except ValueError as e:
        return str(e), 400

    data_inicio_ext = data_formatada_extenso(dt_ini)
    data_termino_ext = data_formatada_extenso(dt_fim)

    franquia_fmt = franquia_com_extenso(franquia_raw)
    valor_fmt = formatar_valor_com_reais(valor_raw)

    data_assinatura = data_formatada_extenso(datetime.now())

    tpl_path = "contrato_template.docx"
    if not os.path.exists(tpl_path):
        return "contrato_template.docx não encontrado no projeto.", 500

    doc = DocxTemplate(tpl_path)
    ctx = {
        "DENOMINACAO": denominacao,
        "CPF": cpf,
        "ENDERECO": endereco,
        "TELEFONE": telefone,
        "EMAIL": email,
        "EQUIPAMENTO": equipamento,
        "ACESSORIOS": acessorios,
        "DATA_INICIO": data_inicio_ext,
        "DATA_TERMINO": data_termino_ext,
        "FRANQUIA": franquia_fmt,
        "VALOR": valor_fmt,
        "DATA_ASSINATURA": data_assinatura,
        "ASSINATURA_CLIENTE": denominacao,
    }

    doc.render(ctx)

    with tempfile.TemporaryDirectory() as tmp:
        docx_out = os.path.join(tmp, "contrato_gerado.docx")
        doc.save(docx_out)

        try:
            pdf_out = docx_para_pdf(docx_out, tmp)
        except FileNotFoundError:
            return "LibreOffice não encontrado (soffice). Verifique o deploy.", 500

        with open(pdf_out, "rb") as f:
            pdf_bytes = f.read()

    bio = BytesIO(pdf_bytes)
    bio.seek(0)
    return send_file(bio, mimetype="application/pdf", as_attachment=True,
                     download_name=f"Contrato {denominacao}.pdf")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=True)
