from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
import io
import zipfile
from openpyxl.utils import get_column_letter
import json
import re
from datetime import date, timedelta

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

POSTO_CONFIGS = {
    'VILAS': {
        'col_map': {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'ELO CREDITO':7,'VISA DEBITO':8,'ELO DEBITO':9,'PIX':11},
        'taxas': {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0142,8:0.0078,9:0.0082,11:0.0074},
        'turnos': ['T3 DIA ANTERIOR','TURNO 1','TURNO 2','TURNO 3'],
        'turno_offset': {'T3 DIA ANTERIOR':3,'TURNO 1':4,'TURNO 2':5,'TURNO 3':6},
        'linhas_por_dia': 11,
        'row_bruto': 7, 'row_liq': 8,
    },
    'CENTRO': {
        'col_map': {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'VISA DEBITO':7,'ELO CREDITO':8,'ELO DEBITO':9,'PIX':11},
        'taxas': {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0078,8:0.0142,9:0.0082,11:0.0074},
        'turnos': ['T3 DIA ANTERIOR','TURNO 1','TURNO 2','TURNO 3'],
        'turno_offset': {'T3 DIA ANTERIOR':3,'TURNO 1':4,'TURNO 2':5,'TURNO 3':6},
        'linhas_por_dia': 11,
        'row_bruto': 7, 'row_liq': 8,
    },
    'BURAQUINHO': {
        'col_map': {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'ELO CREDITO':7,'VISA DEBITO':8,'ELO DEBITO':9,'PIX':11},
        'taxas': {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0142,8:0.0078,9:0.0082,11:0.0074},
        'turnos': ['T3 DIA ANTERIOR','TURNO 1','TURNO 2','TURNO 3'],
        'turno_offset': {'T3 DIA ANTERIOR':3,'TURNO 1':4,'TURNO 2':5,'TURNO 3':6},
        'linhas_por_dia': 11,
        'row_bruto': 7, 'row_liq': 8,
    },
    'ALADAH': {
        'col_map': {'AMEX':2,'MASTER CREDITO':3,'MASTER DEBITO':4,'VISA CREDITO':5,'VISA DEBITO':6,'ELO CREDITO':7,'ELO DEBITO':8,'PIX':10},
        'taxas': {2:0.0210,3:0.0185,4:0.0078,5:0.0173,6:0.0078,7:0.0142,8:0.0082,10:0.0074},
        'turnos': ['TURNO 1','TURNO 2'],
        'turno_offset': {'TURNO 1':3,'TURNO 2':4},
        'linhas_por_dia': 9,
        'row_bruto': 5, 'row_liq': 6,
    },
    'PHAB': {
        'col_map': {'AMEX':2,'HIPER':3,'MASTER CREDITO':4,'MASTER DEBITO':5,'VISA CREDITO':6,'ELO CREDITO':7,'VISA DEBITO':8,'ELO DEBITO':9,'PIX':12},
        'taxas': {2:0.0210,3:0.0249,4:0.0185,5:0.0078,6:0.0173,7:0.0142,8:0.0078,9:0.0082,12:0.0074},
        'turnos': ['TURNO 1','TURNO 2'],
        'turno_offset': {'TURNO 1':2,'TURNO 2':3},
        'linhas_por_dia': 9,
        'row_bruto': 5, 'row_liq': 6,
    },
}

def to_min(s):
    h, m = map(int, str(s).strip().split(':')[:2])
    return h * 60 + m

NOMES_CAMPOS = {
    't3af': 'T3 anterior — fechou às', 't1i': 'Turno 1 (início)', 't1f': 'Turno 1 (até)',
    't2i': 'Turno 2 (início)', 't2f': 'Turno 2 (até)', 't3i': 'Turno 3 — abre às',
}

def valida_horarios(dias_config, turnos):
    campos = ['t3af', 't1i', 't1f', 't2i', 't2f', 't3i'] if 'T3 DIA ANTERIOR' in turnos else ['t1i', 't1f', 't2i', 't2f']
    for d in dias_config:
        if not isinstance(d.get('dia'), int) or not 1 <= d['dia'] <= 31:
            return 'Preencha o "Dia do mês" (1 a 31).'
        t = d.get('turnos', {})
        for c in campos:
            v = str(t.get(c) or '').strip()
            if not re.match(r'^\d{1,2}:\d{2}', v):
                return f'Dia {d["dia"]}: o horário "{NOMES_CAMPOS[c]}" está vazio ou incompleto. Preencha hora E minuto.'
        m = {c: to_min(t[c]) for c in campos}
        pares = [('t3af', 't1f'), ('t1i', 't1f'), ('t1f', 't2f'), ('t2i', 't2f'), ('t2f', 't3i')] if 'T3 DIA ANTERIOR' in turnos else [('t1i', 't1f'), ('t1f', 't2f'), ('t2i', 't2f')]
        for a, b in pares:
            if m[a] > m[b]:
                return f'Dia {d["dia"]}: "{NOMES_CAMPOS[a]}" ({t[a]}) está depois de "{NOMES_CAMPOS[b]}" ({t[b]}). Confira os horários.'
    return None

def parse_hora(val):
    s = str(val).strip()
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

def parse_data_dia(val):
    s = str(val).strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return int(m.group(1))
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return int(m.group(3))
    try:
        f = float(s)
        if f > 40000:
            d = date(1899, 12, 30) + timedelta(days=int(f))
            return d.day
    except:
        pass
    return None

def classifica_turno(min_val, turnos_config, t3af=None, t1i=None, t1f=None, t2i=None, t2f=None, t3i=None):
    # Postos com 4 turnos: cada turno vai do FECHAMENTO do anterior até o seu fechamento.
    # Assim nenhuma venda que caia no intervalo entre fechar um turno e abrir o próximo se perde
    # (ela entra no turno que está abrindo).
    if 'T3 DIA ANTERIOR' in turnos_config:
        if min_val <= t3af:
            return 'T3 DIA ANTERIOR'
        if min_val <= t1f:
            return 'TURNO 1'
        if min_val <= t2f:
            return 'TURNO 2'
        return 'TURNO 3'
    # Postos com 2 turnos (PHAB / ALADAH)
    if t1i <= min_val <= t1f:
        return 'TURNO 1'
    if t1f < min_val <= t2f:
        return 'TURNO 2'
    return None

def fmt_min(m):
    return f'{m // 60:02d}:{m % 60:02d}'

def map_bandeira(band, forma):
    b = band.lower()
    f = (forma or '').lower()
    is_d = 'débit' in f or 'debito' in f or 'débito' in f or 'debit' in f
    if 'visa' in b: return 'VISA DEBITO' if is_d else 'VISA CREDITO'
    if 'master' in b: return 'MASTER DEBITO' if is_d else 'MASTER CREDITO'
    if 'elo' in b: return 'ELO DEBITO' if is_d else 'ELO CREDITO'
    if 'amex' in b or 'american' in b: return 'AMEX'
    if 'hiper' in b: return 'HIPER'
    return None

def encontra_header(rows):
    for i, row in enumerate(rows):
        r = [str(c).strip().lower() for c in row]
        hi = next((j for j, c in enumerate(r) if c == 'hora da venda'), -1)
        if hi != -1:
            return {
                'idx': i, 'hora': hi,
                'data': next((j for j, c in enumerate(r) if c == 'data da venda'), -1),
                'valor': next((j for j, c in enumerate(r) if c == 'valor bruto'), -1),
                'band': next((j for j, c in enumerate(r) if c == 'bandeira'), -1),
                'status': next((j for j, c in enumerate(r) if c == 'status da venda'), -1),
                'forma': next((j for j, c in enumerate(r) if c == 'forma de pagamento'), -1),
            }
    return None

# ---------------------------------------------------------------------------
# Gravação direta no XML da aba (rápida): só a aba do mês é alterada, o resto do
# arquivo é copiado byte a byte. Evita abrir/regravar a planilha inteira, que com
# muitas abas estoura o tempo limite do servidor grátis.
# ---------------------------------------------------------------------------
def _caminho_aba(zf, aba):
    wb_xml = zf.read('xl/workbook.xml').decode('utf-8')
    rels = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = None
    for m in re.finditer(r'<sheet\b[^>]*>', wb_xml):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag)
        if nm and _xml_unescape(nm.group(1)) == aba:
            rid = re.search(r'r:id="([^"]*)"', tag).group(1)
            break
    if not rid:
        return None
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        tag = m.group(0)
        if f'Id="{rid}"' in tag:
            alvo = re.search(r'Target="([^"]*)"', tag).group(1)
            alvo = alvo.lstrip('/')
            return alvo if alvo.startswith('xl/') else 'xl/' + alvo
    return None

def _xml_unescape(t):
    return t.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')

def _col_num(ref):
    letras = re.match(r'[A-Z]+', ref).group(0)
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n

_RE_CELL = r'<c r="{ref}"(?:\s[^>]*?)?(?:/>|>.*?</c>)'

def _set_cell(xml, ref, valor=None, formula=None):
    """Grava número (valor) ou fórmula na célula, mantendo o estilo (s=)."""
    m = re.search(_RE_CELL.format(ref=ref), xml, flags=re.S)
    estilo = ''
    if m:
        old = m.group(0)
        if formula is not None:
            fm = re.search(r'<f(?:\s[^>]*)?>(.*?)</f>', old, flags=re.S)
            if fm and 't="shared"' not in old and _xml_unescape(fm.group(1)) == formula:
                return xml  # fórmula já está certa
            if 't="shared"' in old:
                return xml  # não mexe em fórmula compartilhada
        sm = re.search(r'\ss="(\d+)"', old[:old.find('>') + 1])
        if sm:
            estilo = f' s="{sm.group(1)}"'
    if formula is not None:
        f_esc = formula.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        novo = f'<c r="{ref}"{estilo}><f>{f_esc}</f></c>'
    else:
        novo = f'<c r="{ref}"{estilo}><v>{repr(float(valor)) if valor % 1 else int(valor)}</v></c>'
    if m:
        return xml[:m.start()] + novo + xml[m.end():]
    # célula não existe: insere na linha, na posição da coluna
    row_num = re.match(r'[A-Z]+(\d+)', ref).group(1)
    rm = re.search(rf'<row r="{row_num}"(?:\s[^>]*?)?(?:/>|>.*?</row>)', xml, flags=re.S)
    if not rm:
        raise ValueError(f'Linha {row_num} não encontrada na aba')
    row_xml = rm.group(0)
    if row_xml.endswith('/>'):
        row_novo = row_xml[:-2] + '>' + novo + '</row>'
    else:
        alvo_col = _col_num(ref)
        pos = len(row_xml) - len('</row>')
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"', row_xml):
            if _col_num(cm.group(1)) > alvo_col:
                pos = cm.start()
                break
        row_novo = row_xml[:pos] + novo + row_xml[pos:]
    return xml[:rm.start()] + row_novo + xml[rm.end():]

def grava_xlsx_rapido(arquivo_bytes, aba, valores, formulas):
    """valores: {ref: numero}; formulas: {ref: 'B7-(B7*0.021)'}"""
    zin = zipfile.ZipFile(io.BytesIO(arquivo_bytes))
    caminho = _caminho_aba(zin, aba)
    if not caminho:
        return None
    xml = zin.read(caminho).decode('utf-8')
    for ref, v in valores.items():
        xml = _set_cell(xml, ref, valor=v)
    for ref, f in formulas.items():
        xml = _set_cell(xml, ref, formula=f)
    # apaga os resultados guardados das fórmulas desta aba -> Excel/LibreOffice recalculam ao abrir
    xml = re.sub(r'(</f>|<f\b[^>]*/>)<v>[^<]*</v>', r'\1', xml)
    xml = re.sub(r'(<c\b[^>]*?)\st="(?:str|e|b)"([^>]*>(?:<f\b[^>]*/>|<f\b[^>]*>.*?</f>)</c>)', r'\1\2', xml)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            nome = item.filename
            if nome == 'xl/calcChain.xml':
                continue  # o Excel recria sozinho
            dados = zin.read(nome)
            if nome == caminho:
                dados = xml.encode('utf-8')
            elif nome == 'xl/workbook.xml':
                w = dados.decode('utf-8')
                if '<calcPr' in w:
                    w = re.sub(r'<calcPr\b([^>]*?)/?>', lambda m: '<calcPr' + re.sub(r'\sfullCalcOnLoad="[^"]*"', '', m.group(1)).rstrip('/') + ' fullCalcOnLoad="1"/>', w, count=1)
                else:
                    w = w.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
                dados = w.encode('utf-8')
            elif nome == 'xl/_rels/workbook.xml.rels':
                dados = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', '', dados.decode('utf-8')).encode('utf-8')
            elif nome == '[Content_Types].xml':
                dados = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', '', dados.decode('utf-8')).encode('utf-8')
            zout.writestr(item, dados)
    out.seek(0)
    return out

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
            return jsonify({'erro': 'Arquivo ou dados nao enviados'}), 400

        dados = json.loads(dados_json)
        posto = dados['posto']
        aba = dados['aba']
        dias_config = dados['dias']

        cfg = POSTO_CONFIGS.get(posto)
        if not cfg:
            return jsonify({'erro': f'Posto {posto} nao reconhecido'}), 400

        col_map = cfg['col_map']
        taxas = cfg['taxas']
        turnos = cfg['turnos']
        turno_offset = cfg['turno_offset']
        linhas_por_dia = cfg['linhas_por_dia']
        row_bruto = cfg['row_bruto']
        row_liq = cfg['row_liq']

        erro_h = valida_horarios(dias_config, turnos)
        if erro_h:
            return jsonify({'erro': erro_h}), 400

        avisos = []

        from collections import defaultdict
        matriz = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))

        def processar_arquivo(file, tipo):
            if not file:
                return
            wb2 = openpyxl.load_workbook(file, data_only=True)
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
                valor = 0
                try:
                    valor = float(str(row[h['valor']]).replace(',', '.'))
                except:
                    pass
                if valor <= 0:
                    continue
                dia_tx = parse_data_dia(row[h['data']]) if h['data'] >= 0 else None
                if dia_tx is None:
                    continue
                cfg_dia = next((d for d in dias_config if d['dia'] == dia_tx), None)
                if not cfg_dia:
                    continue
                t = cfg_dia['turnos']
                tm = {k: to_min(v) for k, v in t.items() if v}
                turno = classifica_turno(min_val, turnos, **tm)
                if not turno:
                    avisos.append(f'Dia {dia_tx} {fmt_min(min_val)}: venda de R$ {valor:.2f} fora dos turnos (não lançada)'.replace('.', ','))
                    continue
                if 'T3 DIA ANTERIOR' in turnos:
                    ini = {'TURNO 1': tm.get('t1i'), 'TURNO 2': tm.get('t2i'), 'TURNO 3': tm.get('t3i')}.get(turno)
                    if ini is not None and min_val < ini:
                        avisos.append(f'Dia {dia_tx} {fmt_min(min_val)} — R$ {valor:.2f} — sem caixa aberto → lançada no {turno}. Confira se foi registrada manualmente.'.replace('.', ',').replace(', Confira', '. Confira').replace('manualmente,', 'manualmente.'))
                coluna = 'PIX' if tipo == 'pix' else map_bandeira(str(row[h['band']] or ''), str(row[h['forma']] or ''))
                if coluna and coluna in col_map:
                    matriz[dia_tx][turno][coluna] += valor

        processar_arquivo(cartoes, 'cartoes')
        processar_arquivo(pix, 'pix')

        # Monta as células a gravar
        valores = {}
        for dia_tx, turnos_dia in matriz.items():
            base = (dia_tx - 1) * linhas_por_dia
            # zera TODOS os turnos do dia (evita sobrar valor de uma rodada anterior)
            for turno in turnos:
                row = base + turno_offset[turno]
                for col_idx in col_map.values():
                    valores[f'{get_column_letter(col_idx)}{row}'] = 0
            for turno, colunas in turnos_dia.items():
                if turno not in turno_offset:
                    continue
                row = base + turno_offset[turno]
                for col_name, valor in colunas.items():
                    if valor > 0:
                        valores[f'{get_column_letter(col_map[col_name])}{row}'] = round(valor, 2)

        # Fórmulas de líquido: só nos dias que foram processados agora.
        # Dias anteriores ficam com a taxa que já tinham (se a Cielo mudar a taxa, o histórico não muda).
        formulas = {}
        for d in sorted(matriz.keys()):
            base_d = (d - 1) * linhas_por_dia
            r_liq = base_d + row_liq
            r_bruto = base_d + row_bruto
            for col_idx, taxa in taxas.items():
                L = get_column_letter(col_idx)
                formulas[f'{L}{r_liq}'] = f'{L}{r_bruto}-({L}{r_bruto}*{taxa})'

        planilha_bytes = planilha.read()
        try:
            output = grava_xlsx_rapido(planilha_bytes, aba, valores, formulas)
        except zipfile.BadZipFile:
            return jsonify({'erro': 'A planilha mensal não é um arquivo .xlsx válido.'}), 400
        if output is None:
            nomes = re.findall(r'<sheet\b[^>]*name="([^"]*)"', zipfile.ZipFile(io.BytesIO(planilha_bytes)).read('xl/workbook.xml').decode('utf-8'))
            return jsonify({'erro': f'Aba "{aba}" nao encontrada. Abas disponiveis: {", ".join(_xml_unescape(n) for n in nomes)}'}), 400

        dias_str = '_'.join(str(d['dia']) for d in dias_config)
        nome_arquivo = f"{posto}_{aba}_DIAS{dias_str}.xlsx"

        resp = send_file(output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True, download_name=nome_arquivo)
        if avisos:
            from urllib.parse import quote
            resp.headers['X-Avisos'] = quote(json.dumps(avisos[:30], ensure_ascii=False))
            resp.headers['Access-Control-Expose-Headers'] = 'X-Avisos'
        return resp

    except Exception as e:
        import traceback
        return jsonify({'erro': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(debug=True)

