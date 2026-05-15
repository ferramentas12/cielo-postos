from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
import io
import json
from datetime import datetime

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

def to_min(s):
    h, m = map(int, s.split(':'))
    return h * 60 + m

def parse_hora(val):
    s = str(val).strip()
    import re
    m = re.match(r'^(\d{1,2}):(\d{2})', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    try:
        f = float(s)
        if f < 1:
            return round(f * 1440)
    except:
        pass
    return None

def classifica_turno(min_val, t3af, t1i, t1f, t2i, t2f, t3i):
    if min_val <= t3af: return 'T3 DIA ANTERIOR'
    if t1i <= min_val <= t1f: return 'TURNO 1'
    if t2i <= min_val <= t2f: return 'TURNO 2'
    if min_val >= t3i: return 'TURNO 3'
    return None

def map_bandeira(band, forma):
    b = band.lower()
    f = (forma or '').lower()
    is_d = 'débit' in f or 'debito' in f or 'débito' in f
    if 'visa' in b: return 'VISA DEBITO' if is_d else 'VISA CREDITO'
    if 'master' in b: return 'MASTER DEBITO' if is_d else 'MASTER CREDITO'
    if 'elo' in b: return 'ELO DEBITO' if is_d else 'ELO CREDITO'
    if 'amex' in b or 'american' in b: return 'AMEX'
    if 'hiper' in b: return 'HIPER'
    return None

def parse_data(val):
    """Extrai dia do mês de um valor de data"""
    s = str(val).strip()
    import re
    # formato DD/MM/YYYY
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return int(m.group(1))
    # formato YYYY-MM-DD
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return int(m.group(3))
    # número serial do Excel
    try:
        f = float(s)
        if f > 40000:
            from datetime import date
            delta = date(1899, 12, 30)
            from datetime import timedelta
            d = delta + timedelta(days=int(f))
            return d.day
    except:
        pass
    return None

def encontra_header(rows):
    for i, row in enumerate(rows):
        r = [str(c).strip().lower() for c in row]
        hi = next((j for j, c in enumerate(r) if c == 'hora da venda'), -1)
        if hi != -1:
            return {
                'idx': i,
                'hora': hi,
                'data': next((j for j, c in enumerate(r) if c == 'data da venda'), -1),
                'valor': next((j for j, c in enumerate(r) if c == 'valor bruto'), -1),
                'band': next((j for j, c in enumerate(r) if c == 'bandeira'), -1),
                'status': next((j for j, c in enumerate(r) if c == 'status da venda'), -1),
                'forma': next((j for j, c in enumerate(r) if c == 'forma de pagamento'), -1),
            }
    return None

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/preencher', methods=['POST'])
def preencher():
    try:
        planilha = request.files.get('planilha')
        cartoes = request.files.get('cartoes')
        pix = request.files.get('pix')
        dados_json = request.form.get('dados')

        if not planilha or not dados_json:
            return jsonify({'erro': 'Arquivo ou dados não enviados'}), 400

        dados = json.loads(dados_json)
        posto = dados['posto']
        aba = dados['aba']
        dias_config = dados['dias']  # lista de {dia, turnos: {t3af, t1i, t1f, t2i, t2f, t3i}}
        modo = dados.get('modo', 1)

        col_map = POSTO_CONFIGS.get(posto)
        taxas = TAXAS.get(posto)
        if not col_map:
            return jsonify({'erro': f'Posto {posto} não reconhecido'}), 400

        # Processar arquivos da Cielo
        # matriz[dia][turno][coluna] = valor
        from collections import defaultdict
        matriz = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

        def processar_arquivo(file, tipo):
            if not file:
                return
            import openpyxl as ox
            wb2 = ox.load_workbook(file, data_only=True)
            ws2 = wb2.active
            rows = [[cell.value for cell in row] for row in ws2.iter_rows()]
            h = encontra_header(rows)
            if not h:
                return
            for row in rows[h['idx']+1:]:
                if not row[h['hora']]:
                    continue
                status = str(row[h['status']] or '').strip().lower()
                if status and status != 'aprovada':
                    continue
                min_val = parse_hora(row[h['hora']])
                if min_val is None:
                    continue
                valor = float(str(row[h['valor']]).replace(',', '.') or 0)
                if valor <= 0:
                    continue
                # Descobrir o dia desta transação
                dia_tx = parse_data(row[h['data']]) if h['data'] >= 0 else None
                if dia_tx is None:
                    continue
                # Encontrar config do turno para este dia
                cfg_dia = next((d for d in dias_config if d['dia'] == dia_tx), None)
                if not cfg_dia:
                    continue
                t = cfg_dia['turnos']
                turno = classifica_turno(
                    min_val,
                    to_min(t['t3af']), to_min(t['t1i']), to_min(t['t1f']),
                    to_min(t['t2i']), to_min(t['t2f']), to_min(t['t3i'])
                )
                if not turno:
                    continue
                if tipo == 'pix':
                    coluna = 'PIX'
                else:
                    coluna = map_bandeira(str(row[h['band']] or ''), str(row[h['forma']] or ''))
                if coluna and coluna in col_map:
                    matriz[dia_tx][turno][coluna] += valor

        processar_arquivo(cartoes, 'cartoes')
        processar_arquivo(pix, 'pix')

        # Preencher planilha
        wb = openpyxl.load_workbook(planilha)
        if aba not in wb.sheetnames:
            return jsonify({'erro': f'Aba {aba} não encontrada'}), 400
        ws = wb[aba]

        for dia_tx, turnos in matriz.items():
            base = (dia_tx - 1) * LINHAS_POR_DIA
            for turno, colunas in turnos.items():
                if turno not in TURNO_OFFSET:
                    continue
                row = base + TURNO_OFFSET[turno]
                for col_name, valor in colunas.items():
                    if valor > 0:
                        ws.cell(row=row, column=col_map[col_name]).value = round(valor, 2)

        # Atualizar taxas em todos os dias
        for d in range(1, 32):
            base_d = (d - 1) * LINHAS_POR_DIA
            row_liq = base_d + 8
            row_bruto = base_d + 7
            for col_idx, taxa in taxas.items():
                bruto_ref = ws.cell(row=row_bruto, column=col_idx).coordinate
                ws.cell(row=row_liq, column=col_idx).value = f'={bruto_ref}-({bruto_ref}*{taxa})'

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        dias_str = '_'.join(str(d['dia']) for d in dias_config)
        nome_arquivo = f"{posto}_{aba}_DIAS{dias_str}.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nome_arquivo
        )

    except Exception as e:
        import traceback
        return jsonify({'erro': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(debug=True)
