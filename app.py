import re
from num2words import num2words
from docxtpl import DocxTemplate
import tempfile
import os

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
    # 1000 -> 1.000
    return f"{n:,}".replace(",", ".")

def franquia_formatada_e_extenso(valor: str):
    n = int(so_digitos(valor))
    return formatar_inteiro_ptbr(n), num2words(n, lang="pt_BR")

def valor_formatado_e_extenso(valor: str):
    v = float(valor)
    v_fmt = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    ext = num2words(int(round(v)), lang="pt_BR")
    return v_fmt, ext

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
            "DATA_ASSINATURA": data_formatada(),  # automática
        }

        with tempfile.TemporaryDirectory() as tmp:
            docx_saida = os.path.join(tmp, "contrato_gerado.docx")
            doc.render(context)
            doc.save(docx_saida)

            pdf_saida = docx_para_pdf(docx_saida, tmp)

            nome_limpo = limpar_nome_arquivo(denominacao)
            return send_file(pdf_saida, as_attachment=True, download_name=f"Contrato - {nome_limpo}.pdf")

    return render_template("contrato.html")
