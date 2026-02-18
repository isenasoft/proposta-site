from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
from datetime import datetime
from num2words import num2words
import os

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
    extenso = num2words(valor_float, lang='pt_BR')
    return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X",".") + f" ({extenso} reais)"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/proposta", methods=["GET","POST"])
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
        context = {
            "DATA": data_formatada(),
            "CLIENTE": cliente,
            "CPF": cpf,
