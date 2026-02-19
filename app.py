import os
import tempfile
import subprocess
from datetime import datetime, timedelta
from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate
from num2words import num2words
import psycopg2

app = Flask(__name__)

# ==============================
# CONFIG BANCO
# ==============================

def db_conn():
    url = os.getenv("DATABASE_URL")
    if not url:
        raise RuntimeError("DATABASE_URL não encontrada.")
    return psycopg2.connect(url)

def criar_tabela():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("""
        CREATE TABLE IF NOT EXISTS propostas (
            id SERIAL PRIMARY KEY,
            cliente TEXT,
            cpf TEXT,
            modelo TEXT,
            franquia TEXT,
            valor TEXT,
            criado_em TIMESTAMP DEFAULT NOW()
        )
        """)
        conn.commit()

def limpar_propostas_expiradas():
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE criado_em < NOW() - INTERVAL '10 days'")
        conn.commit()

criar_tabela()

# ==============================
# FUNÇÕES AUXILIARES
# ==============================

def data_formatada():
    meses = {
        1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",
        5:"Maio",6:"Junho",7:"Julho",8:"Agosto",
        9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"
    }
    hoje = datetime.now()
    return f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

def data_extenso_por_digitos(valor: str) -> str:
    v = ''.join(filter(str.isdigit, valor))

    if len(v) == 6:
        dia = int(v[0:2])
        mes = int(v[2:4])
        ano = 2000 + int(v[4:6])
    elif len(v) == 8:
        dia = int(v[0:2])
        mes = int(v[2:4])
        ano = int(v[4:8])
    else:
        raise ValueError("Data inválida")

    meses = {
        1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",
        5:"Maio",6:"Junho",7:"Julho",8:"Agosto",
        9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"
    }

    return f"{dia} de {meses[mes]} de {ano}"

def formatar_valor(valor):
    valor_float = float(valor)
    extenso = num2words(valor_float, lang='pt_BR')
    return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X",".") + f" ({extenso} reais)"

def formatar_numero_extenso(numero):
    n = int(numero)
    return f"{n} ({num2words(n, lang='pt_BR')})"

def docx_para_pdf(docx_path, out_dir):
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True
    )
    return docx_path.replace(".docx", ".pdf")

# ==============================
# ROTAS
# ==============================

@app.route("/")
def index():
    return render_template("index.html")

# ==============================
# PROPOSTA
# ==============================

@app.route("/proposta", methods=["GET","POST"])
def proposta():
    if request.method == "POST":
        cliente = request.form["cliente"]
        cpf = request.form["cpf"]
        modelo = request.form["modelo"]
        franquia = request.form["franquia"]
        valor = request.form["valor"]

        with db_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO propostas (cliente, cpf, modelo, franquia, valor)
                VALUES (%s,%s,%s,%s,%s)
            """, (cliente, cpf, modelo, franquia, valor))
            conn.commit()

        doc = DocxTemplate("template.docx")

        context = {
            "DATA": data_formatada(),
            "CLIENTE": cliente,
            "CPF": cpf,
            "MODELO": modelo,
            "FRANQUIA": formatar_numero_extenso(franquia),
            "VALOR": formatar_valor(valor),
        }

        doc.render(context)

        with tempfile.TemporaryDirectory() as tmp:
            caminho_docx = os.path.join(tmp, "proposta.docx")
            doc.save(caminho_docx)
            pdf = docx_para_pdf(caminho_docx, tmp)

            return send_file(
                pdf,
                as_attachment=True,
                download_name=f"Proposta {cliente}.pdf"
            )

    return render_template("proposta.html")

# ==============================
# CONTRATO
# ==============================

@app.route("/contrato", methods=["GET","POST"])
def contrato():
    if request.method == "POST":
        cliente = request.form["cliente"]
        cpf = request.form["cpf"]
        endereco = request.form["endereco"]
        telefone = request.form["telefone"]
        email = request.form["email"]
        equipamento = request.form["equipamento"]
        acessorios = request.form["acessorios"]
        data_inicio = data_extenso_por_digitos(request.form["data_inicio"])
        data_fim = data_extenso_por_digitos(request.form["data_fim"])
        franquia = formatar_numero_extenso(request.form["franquia"])
        valor = formatar_valor(request.form["valor"])

        doc = DocxTemplate("contrato_template.docx")

        context = {
            "CLIENTE": cliente,
            "CPF": cpf,
            "ENDERECO": endereco,
            "TELEFONE": telefone,
            "EMAIL": email,
            "EQUIPAMENTO": equipamento,
            "ACESSORIOS": acessorios,
            "DATA_INICIO": data_inicio,
            "DATA_FIM": data_fim,
            "FRANQUIA": franquia,
            "VALOR": valor,
            "DATA_ASSINATURA": data_formatada(),
            "ASSINATURA": cliente
        }

        doc.render(context)

        with tempfile.TemporaryDirectory() as tmp:
            caminho_docx = os.path.join(tmp, "contrato.docx")
            doc.save(caminho_docx)
            pdf = docx_para_pdf(caminho_docx, tmp)

            return send_file(
                pdf,
                as_attachment=True,
                download_name=f"Contrato {cliente}.pdf"
            )

    return render_template("contrato.html")

# ==============================
# RECENTES
# ==============================

@app.route("/recentes")
def recentes():
    limpar_propostas_expiradas()

    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("SELECT id, cliente, criado_em FROM propostas ORDER BY criado_em DESC")
        dados = cur.fetchall()

    return render_template("recentes.html", propostas=dados)

@app.route("/excluir/<int:id>")
def excluir(id):
    with db_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM propostas WHERE id=%s", (id,))
        conn.commit()
    return redirect(url_for("recentes"))

# ==============================

if __name__ == "__main__":
    app.run()
