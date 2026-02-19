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

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


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


def data_extenso_por_digitos(ddmmaaaa: str) -> str:
    v = so_digitos(ddmmaaaa)
    if len(v) != 8:
        raise ValueError("Data inválida (use DDMMAAAA)")
    dia = int(v[0:2])
    mes = int(v[2:4])
    ano = int(v[4:8])
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


# ---------- PROPOSTA ----------
@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "POST":
        cliente = request.form["cliente"]
        cpf = request.form["cpf"]
        modelo = request.form["modelo"]
        franquia = request.form["franquia"]
        valor = formatar_valor_reais(request.form["valor"])
        imagem = request.files["imagem"]

        imagem_path = os.path.join(UPLOAD_FOLDER, imagem.filename)
        imagem.save(imagem_path)

        doc = DocxTemplate("template.docx")
        imagem_doc = InlineImage(doc, imagem_path, width=Mm(80))

        context = {
            "DATA": data_formatada(),
            "CLIENTE": cliente,
            "CPF": cpf,
            "MODELO": modelo,
            "FRANQUIA": franquia,
            "VALOR": valor,
            "IMAGEM": imagem_doc
        }

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "proposta_gerada.docx")
            doc.render(context)
            doc.save(docx_saida)

            pdf_saida = docx_para_pdf(docx_saida, tmp)

            nome = limpar_nome_arquivo(cliente)
            return send_file(
                pdf_saida,
                as_attachment=True,
                download_name=f"Proposta - {nome}.pdf"
            )

    return render_template("proposta.html")


# ---------- CONTRATO ----------
@app.route("/contrato", methods=["GET", "POST"])
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
            return send_file(
                pdf_saida,
                as_attachment=True,
                download_name=f"Contrato - {nome}.pdf"
            )

    return render_template("contrato.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
