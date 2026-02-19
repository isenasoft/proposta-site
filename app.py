from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from datetime import datetime
from num2words import num2words
from pathlib import Path
import os
import re
import tempfile
import subprocess
import psycopg2
import psycopg2.extras
from io import BytesIO

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ================= BANCO =================

def db_conn():
    url = os.environ.get("DATABASE_URL")
    if not url:
        raise RuntimeError("DATABASE_URL não encontrada.")
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    return psycopg2.connect(url, connect_timeout=3)

def db_init():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS propostas (
            id SERIAL PRIMARY KEY,
            cliente TEXT NOT NULL,
            cpf TEXT,
            modelo TEXT,
            franquia INTEGER,
            valor NUMERIC,
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            pdf BYTEA NOT NULL
        );
        """)
        # migração "anti-bug": se você já tinha criado_em em algum momento, tenta copiar
        cur.execute("ALTER TABLE propostas ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ;")
        try:
            # se existir criado_em, copia para created_at quando created_at estiver nulo
            cur.execute("UPDATE propostas SET created_at = criado_em WHERE created_at IS NULL;")
        except Exception:
            conn.rollback()
        conn.commit()

def limpar_propostas_expiradas():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE created_at < NOW() - INTERVAL '10 days';")
        conn.commit()


# ================= FUNÇÕES =================

def data_formatada():
    meses = {
        1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
        7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"
    }
    hoje = datetime.now()
    return f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

def limpar_nome_arquivo(txt: str) -> str:
    txt = (txt or "").strip()
    txt = re.sub(r"[\\/:*?\"<>|]+", "", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt[:80] if txt else "Cliente"

def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    env = os.environ.copy()
    user_install = f"file://{out_dir}/lo-profile"
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--nologo",
            "--nolockcheck",
            f"-env:UserInstallation={user_install}",
            "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", out_dir,
            docx_path
        ],
        check=True,
        env=env
    )
    return str(Path(out_dir) / (Path(docx_path).stem + ".pdf"))

def formatar_valor_reais(valor):
    v = float(valor)
    v_fmt = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    ext = num2words(int(round(v)), lang="pt_BR")
    return f"R$ {v_fmt} ({ext} reais)"

def so_digitos(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def data_extenso_por_digitos(valor: str) -> str:
    v = so_digitos(valor)

    # aceita 20/02/26 (DDMMAA)
    if len(v) == 6:
        dia = int(v[0:2])
        mes = int(v[2:4])
        ano = 2000 + int(v[4:6])
    # aceita 20/02/2026 (DDMMAAAA) ou 20022026
    elif len(v) == 8:
        dia = int(v[0:2])
        mes = int(v[2:4])
        ano = int(v[4:8])
    else:
        raise ValueError("Data inválida (use 20022026, 20/02/2026 ou 20/02/26)")

    meses = {
        1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
        7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"
    }
    return f"{dia} de {meses[mes]} de {ano}"

def formatar_inteiro_ptbr(n: int) -> str:
    return f"{n:,}".replace(",", ".")

def franquia_formatada_e_extenso(valor: str):
    n = int(so_digitos(valor))
    return formatar_inteiro_ptbr(n), num2words(n, lang="pt_BR")

def valor_formatado_e_extenso(valor: str):
    v = float(valor)
    v_fmt = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    ext = num2words(int(round(v)), lang="pt_BR")
    return v_fmt, f"{ext} reais"


# ================= ROTAS =================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "POST":
        cliente = request.form["cliente"]
        cpf = request.form["cpf"]
        modelo = request.form["modelo"]
        franquia_raw = request.form["franquia"]
        valor_raw = request.form["valor"]

        imagem = request.files.get("imagem")
        imagem_path = None
        if imagem and imagem.filename:
            imagem_path = os.path.join(UPLOAD_FOLDER, imagem.filename)
            imagem.save(imagem_path)

        doc = DocxTemplate("template.docx")
        imagem_doc = InlineImage(doc, imagem_path, width=Mm(80)) if imagem_path else ""

        context = {
            "DATA": data_formatada(),
            "CLIENTE": cliente,
            "CPF": cpf,
            "MODELO": modelo,
            "FRANQUIA": franquia_raw,
            "VALOR": formatar_valor_reais(valor_raw),
            "IMAGEM": imagem_doc
        }

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "proposta_gerada.docx")
            doc.render(context)
            doc.save(docx_saida)

            pdf_path = docx_para_pdf(docx_saida, tmp)
            pdf_bytes = open(pdf_path, "rb").read()
            @app.route("/proposta_view/<int:pid>")
def proposta_view(pid):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT pdf FROM propostas WHERE id=%s;", (pid,))
        row = cur.fetchone()
        if not row:
            return "Proposta não encontrada (talvez expirou).", 404

        pdf_bytes = row[0]
        bio = BytesIO(pdf_bytes.tobytes() if hasattr(pdf_bytes, "tobytes") else bytes(pdf_bytes))
        bio.seek(0)
        return send_file(bio, mimetype="application/pdf")

            # salva no banco
            db_init()
            franquia_int = int(so_digitos(franquia_raw)) if so_digitos(franquia_raw) else None
            valor_num = float(valor_raw)

            with db_conn() as conn, conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO propostas (cliente, cpf, modelo, franquia, valor, pdf) VALUES (%s, %s, %s, %s, %s, %s);",
                    (cliente, cpf, modelo, franquia_int, valor_num, psycopg2.Binary(pdf_bytes))
                )
                conn.commit()

            nome = limpar_nome_arquivo(cliente)
            return send_file(pdf_path, as_attachment=True, download_name=f"Proposta - {nome}.pdf")

    return render_template("proposta.html")


@app.route("/recentes")
def recentes():
    db_init()
    limpar_propostas_expiradas()

    propostas = []
    with db_conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
        cur.execute("SELECT id, cliente, created_at FROM propostas ORDER BY created_at DESC;")
        for r in cur.fetchall():
            propostas.append({
                "id": r["id"],
                "cliente": r["cliente"],
                "data": r["created_at"].strftime("%d/%m/%Y %H:%M")
            })

    return render_template("recentes.html", propostas=propostas)


@app.route("/proposta_pdf/<int:pid>")
def proposta_pdf(pid):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT cliente, pdf FROM propostas WHERE id=%s;", (pid,))
        row = cur.fetchone()
        if not row:
            return "Proposta não encontrada (talvez expirou).", 404

        cliente, pdf_bytes = row[0], row[1]
        nome = limpar_nome_arquivo(cliente)

        bio = BytesIO(pdf_bytes.tobytes() if hasattr(pdf_bytes, "tobytes") else bytes(pdf_bytes))
        bio.seek(0)
        return send_file(bio, as_attachment=True, download_name=f"Proposta - {nome}.pdf", mimetype="application/pdf")


@app.route("/contrato_de_proposta/<int:pid>")
def contrato_de_proposta(pid):
    with db_conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
        cur.execute("SELECT cliente, cpf, modelo, franquia, valor FROM propostas WHERE id=%s;", (pid,))
        r = cur.fetchone()
        if not r:
            return "Proposta não encontrada (talvez expirou).", 404

        prefill = {
            "denominacao": r["cliente"] or "",
            "cpf_cnpj": r["cpf"] or "",
            "equipamento": r["modelo"] or "",
            "franquia": str(r["franquia"] or ""),
            "valor_mensal": str(r["valor"] or ""),
        }

    return render_template("contrato.html", prefill=prefill)


@app.route("/contrato", methods=["GET","POST"])
def contrato():
    if request.method == "POST":
        denominacao = request.form["denominacao"]
        cpf_cnpj = request.form["cpf_cnpj"]
        endereco = request.form["endereco"]
        telefone = request.form["telefone"]
        email = request.form["email"]
        equipamento = request.form["equipamento"]
        acessorios = request.form["acessorios"]

        data_inicio = data_extenso_por_digitos(request.form["data_inicio"])
        data_termino = data_extenso_por_digitos(request.form["data_termino"])

        franquia_fmt, franquia_ext = franquia_formatada_e_extenso(request.form["franquia"])
        valor_fmt, valor_ext = valor_formatado_e_extenso(request.form["valor_mensal"])

        doc = DocxTemplate("contrato_template.docx")
        context = {
            "DENOMINACAO": denominacao,
            "CPF_CNPJ": cpf_cnpj,
            "ENDERECO": endereco,
            "TELEFONE": telefone,
            "EMAIL": email,
            "EQUIPAMENTO": equipamento,
            "ACESSORIOS": acessorios,
            "DATA_INICIO": data_inicio,
            "DATA_TERMINO": data_termino,
            "FRANQUIA_FORMATADA": franquia_fmt,
            "FRANQUIA_EXTENSO": franquia_ext,
            "VALOR_MENSAL_FORMATADO": valor_fmt,
            "VALOR_MENSAL_EXTENSO": valor_ext,
            "DATA_ASSINATURA": data_formatada(),
        }

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "contrato_gerado.docx")
            doc.render(context)
            doc.save(docx_saida)

            pdf_saida = docx_para_pdf(docx_saida, tmp)
            nome = limpar_nome_arquivo(denominacao)
            return send_file(pdf_saida, as_attachment=True, download_name=f"Contrato - {nome}.pdf")

    # IMPORTANTE: sempre manda prefill vazio pro template não quebrar
    return render_template("contrato.html", prefill={})


@app.route("/proposta_delete/<int:pid>", methods=["POST"])
def proposta_delete(pid):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE id=%s;", (pid,))
        conn.commit()
    return ("", 204)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
