import os
import re
import io
import base64
import tempfile
import subprocess
from datetime import datetime, date

import psycopg2
from psycopg2.extras import RealDictCursor

from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from docx import Document
from docx.shared import Inches
from num2words import num2words

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret")


# ----------------------------
# DB
# ----------------------------
def db_conn():
    db_url = os.environ.get("DATABASE_URL")
    if not db_url:
        raise RuntimeError("DATABASE_URL não encontrada (conecte o Postgres e crie a variável DATABASE_URL no Railway).")
    return psycopg2.connect(db_url, sslmode="require")


def garantir_tabela():
    """Cria a tabela se não existir (não altera colunas existentes)."""
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS propostas (
            id SERIAL PRIMARY KEY,
            cliente TEXT NOT NULL,
            cpf TEXT,
            modelo TEXT,
            franquia INTEGER,
            valor NUMERIC(12,2),
            pdf BYTEA NOT NULL,
            created_at TIMESTAMP NOT NULL DEFAULT NOW()
        );
        """)
        conn.commit()


def salvar_proposta(cliente, cpf, modelo, franquia_int, valor_num, pdf_bytes):
    """SALVA NO BANCO VALOR COMO NUMÉRICO (ex: 150.00), não texto."""
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO propostas (cliente, cpf, modelo, franquia, valor, pdf, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, NOW())
            RETURNING id;
            """,
            (cliente, cpf, modelo, franquia_int, valor_num, psycopg2.Binary(pdf_bytes)),
        )
        pid = cur.fetchone()[0]
        conn.commit()
        return pid


def listar_propostas():
    with db_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, created_at
            FROM propostas
            ORDER BY created_at DESC
            LIMIT 50;
        """)
        return cur.fetchall()


def pegar_pdf_proposta(pid: int) -> bytes | None:
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT pdf FROM propostas WHERE id=%s;", (pid,))
        row = cur.fetchone()
        return row[0] if row else None


def pegar_proposta(pid: int):
    with db_conn() as conn, conn.cursor(cursor_factory=RealDictCursor) as cur:
        cur.execute("""
            SELECT id, cliente, cpf, modelo, franquia, valor, created_at
            FROM propostas
            WHERE id=%s;
        """, (pid,))
        return cur.fetchone()


def limpar_propostas_expiradas(dias=10):
    """Remove propostas antigas usando created_at (não existe criado_em no seu banco)."""
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE created_at < NOW() - INTERVAL %s;", (f"{dias} days",))
        conn.commit()


# ----------------------------
# Utils de formatação
# ----------------------------
MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

def formatar_data_extenso(d: date) -> str:
    return f"{d.day:02d} de {MESES[d.month]} de {d.year}"

def normalizar_data_digitos(s: str) -> str:
    """Aceita '20022026' ou '20/02/2026' e devolve '20022026'."""
    digits = re.sub(r"\D", "", (s or "").strip())
    return digits

def data_extenso_por_digitos(s: str) -> str:
    digits = normalizar_data_digitos(s)
    if len(digits) != 8:
        raise ValueError("Data inválida (use DDMMAAAA)")
    dd = int(digits[0:2])
    mm = int(digits[2:4])
    yyyy = int(digits[4:8])
    d = date(yyyy, mm, dd)  # pode lançar ValueError se data impossível
    return formatar_data_extenso(d)

def int_por_texto(n: int) -> str:
    # num2words em pt_BR
    return num2words(n, lang="pt_BR")

def valor_por_extenso_reais(valor: float) -> str:
    # Ex: 150.00 -> "cento e cinquenta reais"
    inteiro = int(round(valor))
    txt = num2words(inteiro, lang="pt_BR")
    return f"{txt} reais"

def formatar_moeda_brl(valor: float) -> str:
    # 150 -> "R$ 150,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def slug_nome_arquivo(nome: str) -> str:
    nome = (nome or "").strip()
    nome = re.sub(r"[^\w\s\-]", "", nome, flags=re.UNICODE)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome[:80] if nome else "Cliente"


# ----------------------------
# DOCX -> PDF (LibreOffice)
# ----------------------------
def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nolockcheck",
         "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, f"{base}.pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("Falha ao converter para PDF.")
    return pdf_path


# ----------------------------
# Preencher placeholders no docx
# ----------------------------
def substituir_texto_doc(doc: Document, mapa: dict[str, str]):
    """Substitui texto preservando estilos (reconstrói runs)."""
    for p in doc.paragraphs:
        full = "".join(r.text for r in p.runs)
        if not full:
            continue
        novo = full
        mudou = False
        for k, v in mapa.items():
            if k in novo:
                novo = novo.replace(k, v)
                mudou = True
        if mudou:
            # limpa runs e coloca tudo no 1º run (mantém estilo do primeiro)
            if p.runs:
                p.runs[0].text = novo
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.add_run(novo)

    # também em tabelas
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    full = "".join(r.text for r in p.runs)
                    if not full:
                        continue
                    novo = full
                    mudou = False
                    for k, v in mapa.items():
                        if k in novo:
                            novo = novo.replace(k, v)
                            mudou = True
                    if mudou:
                        if p.runs:
                            p.runs[0].text = novo
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.add_run(novo)

def inserir_imagem_doc(doc: Document, placeholder: str, img_bytes: bytes | None, largura_polegadas=4.8):
    if not img_bytes:
        # se não veio imagem, só remove placeholder
        substituir_texto_doc(doc, {placeholder: ""})
        return

    # acha o parágrafo com o placeholder
    for p in doc.paragraphs:
        if placeholder in p.text:
            # limpa texto
            for r in p.runs:
                r.text = r.text.replace(placeholder, "")
            # insere imagem no mesmo parágrafo
            run = p.add_run()
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            try:
                tmp.write(img_bytes)
                tmp.close()
                run.add_picture(tmp.name, width=Inches(largura_polegadas))
            finally:
                try:
                    os.unlink(tmp.name)
                except:
                    pass
            break


# ----------------------------
# Rotas
# ----------------------------
@app.route("/")
def index():
    garantir_tabela()
    return render_template("index.html")


@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    garantir_tabela()
    if request.method == "GET":
        return render_template("proposta.html")

    try:
        cliente = request.form.get("cliente", "").strip()
        cpf = request.form.get("cpf", "").strip()
        modelo = request.form.get("modelo", "").strip()
        franquia_int = int(re.sub(r"\D", "", request.form.get("franquia", "0")) or "0")
        valor_num = float((request.form.get("valor", "0") or "0").replace(",", "."))
        validade = request.form.get("validade", "").strip()  # se você usa no template
        contrato = request.form.get("contrato", "").strip()  # se você usa no template

        if not cliente:
            raise ValueError("Cliente é obrigatório.")

        # imagem opcional
        img_file = request.files.get("imagem")
        img_bytes = img_file.read() if img_file and img_file.filename else None

        # Formatações pro DOCX
        hoje = date.today()
        data_topo = hoje.strftime("%d/%m/%Y")  # topo automático (igual você pediu)
        valor_fmt = f"{formatar_moeda_brl(valor_num)} ({valor_por_extenso_reais(valor_num)})"
        franquia_fmt = f"{franquia_int} ({int_por_texto(franquia_int)})"

        # Carrega template de proposta
        template_path = os.path.join(os.getcwd(), "template.docx")
        if not os.path.exists(template_path):
            raise RuntimeError("template.docx não encontrado na raiz do projeto.")

        doc = Document(template_path)
        mapa = {
            "{{ CLIENTE }}": cliente,
            "{{ CPF }}": cpf,
            "{{ MODELO }}": modelo,
            "{{ FRANQUIA }}": franquia_fmt,
            "{{ VALOR }}": valor_fmt,
            "{{ DATA }}": data_topo,
            "{{ VALIDADE }}": validade,
            "{{ CONTRATO }}": contrato,
        }
        substituir_texto_doc(doc, mapa)
        inserir_imagem_doc(doc, "{{ IMAGEM }}", img_bytes, largura_polegadas=5.6)  # maior (ajustável)

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "proposta_gerada.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        # SALVA NO BANCO (valor_num é número)
        pid = salvar_proposta(cliente, cpf, modelo, franquia_int, valor_num, pdf_bytes)

        filename = f"Proposta {slug_nome_arquivo(cliente)}.pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        flash(str(e))
        return redirect(url_for("proposta"))


@app.route("/recentes")
def recentes():
    garantir_tabela()
    try:
        limpar_propostas_expiradas(10)
    except Exception:
        # se der qualquer erro de limpeza, não derruba a página
        pass

    propostas = listar_propostas()
    return render_template("recentes.html", propostas=propostas)


@app.route("/proposta/<int:pid>/pdf")
def proposta_pdf(pid):
    garantir_tabela()
    pdf = pegar_pdf_proposta(pid)
    if not pdf:
        return "Não encontrada.", 404
    prop = pegar_proposta(pid)
    cliente = prop["cliente"] if prop else "Cliente"
    filename = f"Proposta {slug_nome_arquivo(cliente)}.pdf"
    return send_file(
        io.BytesIO(pdf),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )


@app.route("/contrato", methods=["GET", "POST"])
def contrato():
    garantir_tabela()

    # prefill vindo de uma proposta recente
    prefill = {}
    prefill_id = request.args.get("prefill_id")
    if prefill_id:
        p = pegar_proposta(int(prefill_id))
        if p:
            prefill = {
                "denominacao": p.get("cliente") or "",
                "cpf_cnpj": p.get("cpf") or "",
                "equipamento": p.get("modelo") or "",
                "franquia": str(p.get("franquia") or ""),
                "valor_mensal": str(p.get("valor") or ""),
            }

    if request.method == "GET":
        return render_template("contrato.html", prefill=prefill)

    try:
        denominacao = request.form.get("denominacao", "").strip()
        cpf_cnpj = request.form.get("cpf_cnpj", "").strip()
        endereco = request.form.get("endereco", "").strip()
        telefone = request.form.get("telefone", "").strip()
        email = request.form.get("email", "").strip()
        equipamento = request.form.get("equipamento", "").strip()
        acessorios = request.form.get("acessorios", "").strip()

        data_inicio = data_extenso_por_digitos(request.form.get("data_inicio", ""))
        data_termino = data_extenso_por_digitos(request.form.get("data_termino", ""))

        franquia_int = int(re.sub(r"\D", "", request.form.get("franquia", "0")) or "0")
        valor_num = float((request.form.get("valor_mensal", "0") or "0").replace(",", "."))

        hoje_extenso = formatar_data_extenso(date.today())
        franquia_fmt = f"{franquia_int} ({int_por_texto(franquia_int)})"
        valor_fmt = f"{formatar_moeda_brl(valor_num)} ({valor_por_extenso_reais(valor_num)})"

        template_path = os.path.join(os.getcwd(), "contrato_template.docx")
        if not os.path.exists(template_path):
            raise RuntimeError("contrato_template.docx não encontrado na raiz do projeto.")

        doc = Document(template_path)

        mapa = {
            "{{ DENOMINACAO }}": denominacao,
            "{{ CPF_CNPJ }}": cpf_cnpj,
            "{{ ENDERECO }}": endereco,
            "{{ TELEFONE }}": telefone,
            "{{ EMAIL }}": email,
            "{{ EQUIPAMENTO }}": equipamento,
            "{{ ACESSORIOS }}": acessorios,
            "{{ DATA_INICIO }}": data_inicio,
            "{{ DATA_TERMINO }}": data_termino,
            "{{ FRANQUIA }}": franquia_fmt,
            "{{ VALOR_MENSAL }}": valor_fmt,
            "{{ DATA_ASSINATURA }}": hoje_extenso,
            "{{ ASSINATURA }}": denominacao,  # repete o nome na assinatura
        }
        substituir_texto_doc(doc, mapa)

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "contrato_gerado.docx")
            doc.save(docx_saida)
            pdf_path = docx_para_pdf(docx_saida, tmp)
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()

        filename = f"Contrato {slug_nome_arquivo(denominacao)}.pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        flash(str(e))
        # volta mantendo prefill
        return redirect(url_for("contrato", prefill_id=prefill_id) if prefill_id else url_for("contrato"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
