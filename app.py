# HOYA Spare Parts Consolidation - Flask App
import os, io, tempfile
from pathlib import Path
from flask import Flask, request, jsonify, send_file
import msoffcrypto, openpyxl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

SAMPLE_SHEET_MAP = {
    'AO':'AO','AR':'AR COAT','AS':'Assembling','FI':'Final',
    'GK':'GK','HC':'HARD COAT','HR':'HOLV ADMIN COMMON','IT':'IT',
    'Material WH':'Material WH','MF':'Mixing Filling','MF-174':'Mixing Filling 174',
    'MF-PNX':'Mixing Filling PNX','Mold':'Mold Preparation',
    'PRO PLAN':'PRO PLAN','Pro-WH':'PRO WH','QA':'QA',
    'Sensity':'SUNTECH','Hóa chất-Technical':'TECHNICAL','Facility':'TPM',
    'Visual':'Visual','Export':'EXPORT',
}
MONTHS = ['Mar','Apr','May','Jun']

import re
def norm(s): return re.sub(r'[^a-z0-9]','',str(s).lower().strip())

def decrypt(fb, pw='sp'):
    try:
        wb = openpyxl.load_workbook(io.BytesIO(fb), data_only=True)
        wb.close(); return fb
    except: pass
    with tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False) as t:
        t.write(fb); p = t.name
    try:
        with open(p,'rb') as f:
            o = msoffcrypto.OfficeFile(f); o.load_key(password=pw)
            b = io.BytesIO(); o.decrypt(b); b.seek(0); return b.read()
    finally: os.unlink(p)

def read_sample(fb):
    wb = openpyxl.load_workbook(io.BytesIO(fb), data_only=True)
    r = {}
    for sn in wb.sheetnames:
        if sn=='Check': continue
        ws = wb[sn]; sec = SAMPLE_SHEET_MAP.get(sn,sn); rows = []
        for row in range(7, ws.max_row+1):
            q = ws.cell(row=row, column=14).value
            if q is None or str(q).strip() in ('','0',0): continue
            try: q=float(q)
            except: continue
            d=str(ws.cell(row=row,column=5).value or '').strip()
            p=str(ws.cell(row=row,column=6).value or '').strip()
            if d.upper()=='TOTAL' or p.upper()=='TOTAL': continue
            rows.append({'desc':d,'po':p,'qty':q})
        if rows: r[sec]=rows
    wb.close(); return r

def match_score(s, m):
    sc=0
    ps,nm=norm(s.get('po','')),norm(m.get('p',''))
    if ps and nm:
        if ps==nm: sc+=10
        elif ps in nm or nm in ps: sc+=5
    ds,dm=norm(s.get('desc','')),norm(m.get('d',''))
    if ds and dm:
        if ds==dm: sc+=8
        elif ds in dm or dm in ds: sc+=4
    return sc

def consolidate(mb, sd, month):
    wb=openpyxl.load_workbook(io.BytesIO(mb), data_only=False)
    wd=openpyxl.load_workbook(io.BytesIO(mb), data_only=True)
    s153,f152=f"153-{month}'26",f"152-{month}'26"
    m153=m152=0; um=[]
    for sec,rows in sd.items():
        if s153 in wb.sheetnames:
            ws,w=wb[s153],wd[s153]; ms=[]
            for r in range(12,w.max_row+1):
                if str(w.cell(r,2).value or '').strip().upper()!=sec.upper(): continue
                ms.append({'r':r,'d':str(w.cell(r,1).value or ''),'p':str(w.cell(r,7).value or '')})
            u=set()
            for row in rows:
                best=None;bsc=0
                for mr in ms:
                    if mr['r'] in u: continue
                    s2=match_score(row,mr)
                    if s2>bsc: best,bsc=mr,s2
                if best and bsc>=4:
                    ws.cell(best['r'],23).value=row['qty']; u.add(best['r']); m153+=1
                else: um.append({'section':sec,'po':row.get('po',''),'desc':row.get('desc','')[:60],'qty':row.get('qty',0)})
        if f152 in wb.sheetnames:
            ws,w=wb[f152],wd[f152]; ms=[]
            for r in range(12,w.max_row+1):
                if str(w.cell(r,2).value or '').strip().upper()!=sec.upper(): continue
                ms.append({'r':r,'d':str(w.cell(r,1).value or ''),'p':str(w.cell(r,7).value or '')})
            u=set()
            for row in rows:
                best=None;bsc=0
                for mr in ms:
                    if mr['r'] in u: continue
                    s2=match_score(row,mr)
                    if s2>bsc: best,bsc=mr,s2
                if best and bsc>=4:
                    ws.cell(best['r'],43).value=row['qty']; u.add(best['r']); m152+=1
    b=io.BytesIO(); wb.save(b); b.seek(0); wb.close(); wd.close()
    return b.getvalue(),{'matched153':m153,'matched152':m152,'unmatched':um,'total':sum(len(v) for v in sd.values())}

# --- READ HTML ---
_dir = os.path.dirname(os.path.abspath(__file__))
html_path = os.path.join(_dir, 'index_flask.html')
if os.path.exists(html_path):
    HTML = open(html_path, encoding='utf-8').read()
else:
    HTML = open(os.path.join(_dir, 'index.html'), encoding='utf-8').read()

@app.route('/')
def index(): return HTML

@app.route('/api/consolidate', methods=['POST'])
def api():
    try:
        mf=request.files.get('master'); sfs=request.files.getlist('samples')
        mo=request.form.get('month','May'); pw=request.form.get('password','') or 'sp'
        if not mf: return jsonify({'error':'Thiếu master'}),400
        if not sfs: return jsonify({'error':'Thiếu sample'}),400
        mb=decrypt(mf.read(),pw)
        ad={}
        for sf in sfs:
            d=read_sample(sf.read())
            for s,rows in d.items(): ad.setdefault(s,[]).extend(rows)
        if not ad: return jsonify({'error':'Không có dữ liệu output!'}),400
        rb,st=consolidate(mb,ad,mo)
        on=f"Detail_Spare_parts_{mo}26_CONSOLIDATED.xlsx"
        op=os.path.join(app.config['UPLOAD_FOLDER'],on)
        with open(op,'wb') as f: f.write(rb)
        return jsonify({'success':True,'filename':on,'stats':st,'download_url':f'/api/download/{on}'})
    except Exception as e:
        import traceback
        return jsonify({'error':str(e),'trace':traceback.format_exc()}),500

@app.route('/api/download/<fn>')
def download(fn):
    p=os.path.join(app.config['UPLOAD_FOLDER'],fn)
    if not os.path.exists(p): return jsonify({'error':'Hết hạn, upload lại'}),404
    return send_file(p,as_attachment=True,download_name=fn)

if __name__=='__main__':
    port=int(os.environ.get('PORT',8787))
    app.run(host='0.0.0.0',port=port)
