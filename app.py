from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
import io
import json

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

LINHAS_POR_DIA = 11
TURNO_OFFSET = {'T3 DIA ANTERIOR': 3, 'TURNO 1': 4, 'TURNO 2': 5, 'TURNO 3': 6}

POSTO_CONFIGS = {
    'VILAS':      {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'ELO CREDITO':7,'VISA DEBITO':8,'ELO DEBITO':9,'PIX':11},
    'CENTRO':     {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'VISA DEBITO':7,'ELO CREDITO':8,'ELO DEBITO':9,'PIX':11},
    'BURAQUINHO': {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'ELO CREDITO':7,'VISA DEBITO':8,'ELO DEBITO':9,'PIX':11},
}

TAXAS = {
    'VILAS':      {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0142,8:0.0078,9:0.0082,11:0.0074},
    'CENTRO':     {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0078,8:0.0142,9:0.0082,11:0.0074},
    'BURAQUINHO': {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0142,8:0.0078,9:0.0082,11:0.0074},
}

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/preencher', methods=['POST'])
def preencher():
    try:
        planilha = request.files.get('planilha')
        dados_json = request.form.get('dados')

        if not planilha or not dados_json:
            return jsonify({'erro': 'Arquivo ou dados não enviados'}), 400

        dados = json.loads(dados_json)
        posto = dados['posto']
        aba = dados['aba']
        dias = dados['dias']

        col_map = POSTO_CONFIGS.get(posto)
        taxas = TAXAS.get(posto)

        if not col_map:
            return jsonify({'erro': f'Posto {posto} não reconhecido'}), 400

        wb = openpyxl.load_workbook(planilha)

        if aba not in wb.sheetnames:
            return jsonify({'erro': f'Aba {aba} não encontrada'}), 400

        ws = wb[aba]

        # Preencher cada dia
        for dia_info in dias:
            dia = dia_info['dia']
            base = (dia - 1) * LINHAS_POR_DIA
            matriz = dia_info['matriz']

            for turno, valores in matriz.items():
                if turno not in TURNO_OFFSET:
                    continue
                row = base + TURNO_OFFSET[turno]
                for col_name, valor in valores.items():
                    if col_name in col_map and valor > 0:
                        ws.cell(row=row, column=col_map[col_name]).value = valor

        # Atualizar taxas em todos os dias do mês
        for d in range(1, 32):
            base_d = (d - 1) * LINHAS_POR_DIA
            row_liq = base_d + 8
            row_bruto = base_d + 7
            for col_idx, taxa in taxas.items():
                bruto_ref = ws.cell(row=row_bruto, column=col_idx).coordinate
                ws.cell(row=row_liq, column=col_idx).value = f'={bruto_ref}-({bruto_ref}*{taxa})'

        # Salvar em memória e enviar
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        nome_arquivo = planilha.filename.replace('.xlsx', f'_DIA{dias[0]["dia"]}.xlsx')

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nome_arquivo
        )

    except Exception as e:
        return jsonify({'erro': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
