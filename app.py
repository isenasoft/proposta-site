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

def data_formatada():
    meses = {
        1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
        7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"
    }
    hoje = datetime.now()
    return f"{hoje.day} de {meses[hoje.month]} de {hoje.year}"

def formatar_valor(valor):
    valor_float = float(valor)
    inteiro = int(round(valor_float))
    extenso = num2words(inteiro, lang="pt_BR")
    valor_formatado = f"{valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    return f"R$ {valor_formatado} ({extenso} reais)"

def limpar_nome_arquivo(txt: str) -> str:
    txt = txt.strip()
    txt = re.sub(r"[\\/:*?\"<>|]+", "", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt[:80] if txt else "Cliente"

def docx_para_pdf(docx_path: str, out_dir: str) -> str:
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nolockcheck", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True
    )
    pdf_path = str(Path(out_dir) / (Path(docx_path).stem + ".pdf"))
    return pdf_path

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/proposta", methods=["GET", "POST"])
def proposta():
    if request.method == "POST":
        cliente = request.form["cliente"]
        cpf = request.form["cpf"]
        modelo = request.form["modelo"]
        franquia = request.form["franquia"]
        valor = formatar_valor(request.form["valor"])
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
            download_nome = f"Proposta - {nome}.pdf"

            return send_file(pdf_saida, as_attachment=True, download_name=download_nome)

    return render_template("proposta.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
