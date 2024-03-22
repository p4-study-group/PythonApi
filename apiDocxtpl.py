from flask import Flask, request, jsonify
from flask_cors import CORS
from docxtpl import DocxTemplate, InlineImage
import os
import random

app = Flask(__name__)
CORS(app)

json_data = {
    "lista": [
        {"itens": "x", "DSCServico": "descriçao um", "VCUnid": 1, "VCQuant": 4, "VCValorTotal": 1},
        {"itens": "y", "DSCServico": "descriçao dois", "VCUnid": 6, "VCQuant": 9, "VCValorTotal": 54},
        {"itens": "z", "DSCServico": "descriçao três", "VCUnid": 3, "VCQuant": 6, "VCValorTotal": 18},
    ]
}

def create_table_data(data):
    table_data = []
    for item in data.get("lista", []):
        table_data.append({
            'itens': item.get('itens', ''),
            'DSCServico': item.get('DSCServico', ''),
            'VCUnid': item.get('VCUnid', ''),
            'VCQuant': item.get('VCQuant', ''),
            'VCVU': item.get('VCVU', ''),
            'VCValorTotal': item.get('VCValorTotal', '')
        })
    return table_data

vendas_produtos = []
for row in range(4):
    costPu = random.randint(1, 15)
    nUnits = random.randint(10, 20)
    vendas_produtos.append({"name": "Produto Nº "+str(row+1),
                         "cPu": costPu, "nUnits": nUnits, "revenue": costPu*nUnits})


def main():
    template_file_path = "Proposta com tabela.docx"
    output_file_path = "Proposta criada.docx"

    doc = DocxTemplate(template_file_path)

    tags_lists = {
        "TXTRS": "Nome razão abc",
        "NOMREPRESENTANTECLIENTE": "P4pro max ",
        "vendas_produtos": vendas_produtos,
        "bannerImg": InlineImage(doc,"img/p4.png"),
        "table_data": create_table_data(json_data),
    }

  
    doc.render(tags_lists)

    doc.save(output_file_path)

@app.route('/api/generate-word', methods=['POST'])
def this_page():
   try:
       main()
       return jsonify({"status": "success", "message": "Sucesso, arquivo criado"})
   except Exception as error_mensage:
       return jsonify({"status": "error", "message": str(error_mensage)})

if __name__ == '__main__':
  app.run(debug=True)
