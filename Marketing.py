#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Makropetrol Marketing: edición autónoma para Streamlit, sin claves.
Ejecutar: python -m streamlit run Marketing.py
Subir a GitHub: Marketing.py y requirements.txt.
Catálogo, apertura, interfaz, lógica y exportadores incluidos en este archivo.
Los movimientos se guardan en datos_marketing/marketing.sqlite3 en el SERVIDOR.
Streamlit Community Cloud NO garantiza la conservación de ese disco.
Descarga respaldos y restáuralos desde la app cuando sea necesario.
No utiliza Supabase, claves administrativas ni almacenamiento externo.
Esta edición no migra ni sustituye automáticamente una base PostgreSQL existente.
La publicación no tiene autenticación: cualquier visitante puede modificar datos.
"""
from __future__ import annotations
import base64, calendar, collections, csv, hashlib, html, io, json, math
import os, re, secrets, shutil, socket, sqlite3, sys, tempfile, threading, traceback
import unicodedata, uuid, webbrowser, zipfile, zlib
from contextlib import contextmanager
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
import logging, time, types
from pathlib import Path
from urllib.parse import urlparse, parse_qs

VERSION = '1.3.0-autonomo-sin-claves'
OPENING_DAY = '2026-10-30'
SOURCE_NAME = 'INVENTARIOS MARCAS30-10.xlsx'
SEED_B64 = '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'
DATA_DIR = Path(os.environ.get('MARKETING_DATA_DIR', str(Path(__file__).resolve().parent / 'datos_marketing')))
DB_PATH = DATA_DIR / 'marketing.sqlite3'
# Streamlit vuelve a ejecutar el script por sesión. Compartir el mismo bloqueo
# por ruta evita que una restauración compita con una entrega de otra pestaña.
_runtime = sys.modules.setdefault('_makropetrol_marketing_locks', types.ModuleType('_makropetrol_marketing_locks'))
if not hasattr(_runtime, 'locks'):
    _runtime.locks = {}
LOCK = _runtime.locks.setdefault(str(DB_PATH.resolve()), threading.RLock())
TOKEN = secrets.token_hex(32)
MAX_UPLOAD = 12 * 1024 * 1024
MAX_QTY = 100_000_000
KINDS = {'APERTURA':'Apertura', 'ENTRADA':'Recepción de proveedor', 'ENTREGA':'Entrega a asesor',
         'EVENTO':'Salida para evento', 'PRESTAMO':'Préstamo de exhibidor',
         'LIQUIDACION':'Regreso / cierre de evento', 'AJUSTE':'Conteo físico / ajuste', 'ANULACION':'Anulación'}
PREFIX = {'APERTURA':'AP','ENTRADA':'EN','ENTREGA':'SA','EVENTO':'EV','PRESTAMO':'PR','LIQUIDACION':'RE','AJUSTE':'CO','ANULACION':'AN'}

class AppError(Exception):
    pass

def now_iso():
    return datetime.now(timezone(timedelta(hours=-4))).isoformat(timespec='seconds')

def today():
    return datetime.now(timezone(timedelta(hours=-4))).date()

def text(v, max_len=200):
    s = '' if v is None else str(v).strip()
    if len(s) > max_len: raise AppError(f'El texto supera {max_len} caracteres.')
    if any(ord(c) < 32 and c not in '\n\t' for c in s): raise AppError('El texto contiene caracteres de control.')
    return s

def norm(v):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFKD', text(v,1000)).upper() if not unicodedata.combining(c)).split())

def integer(v, *, zero=False, signed=False):
    # No expandir exponentes ni convertir números enormes en cadenas.
    if v is None or isinstance(v, bool): raise AppError('Escribe una cantidad entera; no dejes celdas vacías.')
    s=str(v).strip()
    if not s or len(s)>35: raise AppError('Cantidad vacía o demasiado larga.')
    if ',' in s: raise AppError('Usa cantidades enteras, sin puntos ni comas de miles.')
    try: d=Decimal(s)
    except InvalidOperation: raise AppError(f'Cantidad inválida: {s[:25]}.')
    if not d.is_finite() or d.copy_abs()>MAX_QTY or d!=d.to_integral_value():
        raise AppError('La cantidad debe ser un entero entre 0 y 100.000.000.')
    n=int(d)
    if not signed and n<0: raise AppError('No se admiten cantidades negativas.')
    if not zero and n==0: raise AppError('La cantidad debe ser mayor que cero.')
    return n

def connect():
    c=sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA foreign_keys=ON')
    c.execute('PRAGMA busy_timeout=30000')
    return c

@contextmanager
def transaction():
    with LOCK:
        c=connect()
        try:
            c.execute('BEGIN IMMEDIATE')
            yield c
            c.commit()
        except BaseException:
            c.rollback()
            raise
        finally: c.close()

@contextmanager
def reading():
    with LOCK:
        c=connect()
        try: yield c
        finally: c.close()

def rows(c, sql, args=()): return [dict(r) for r in c.execute(sql,args).fetchall()]
def meta(c,key,default=''):
    r=c.execute('SELECT value FROM meta WHERE key=?',(key,)).fetchone()
    return r['value'] if r else default

def put_meta(c,key,value):
    c.execute('INSERT INTO meta(key,value) VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',(key,str(value)))

def bump(c): put_meta(c,'revision',int(meta(c,'revision','0'))+1)

def qty(c,pid):
    return c.execute('SELECT COALESCE(SUM(delta),0) FROM lines WHERE product_id=?',(pid,)).fetchone()[0]

def init_db():
    DATA_DIR.mkdir(parents=True,exist_ok=True)
    with transaction() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS brands(id INTEGER PRIMARY KEY,name TEXT NOT NULL,normalized TEXT UNIQUE NOT NULL,logo BLOB);
        CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY,code TEXT UNIQUE NOT NULL,brand_id INTEGER NOT NULL REFERENCES brands(id),
          name TEXT NOT NULL,normalized TEXT NOT NULL,source TEXT NOT NULL DEFAULT '',note TEXT NOT NULL DEFAULT '',created_day TEXT NOT NULL,
          UNIQUE(brand_id,normalized));
        CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY,number TEXT UNIQUE,kind TEXT NOT NULL,day TEXT NOT NULL,
          recipient TEXT NOT NULL DEFAULT '',operator TEXT NOT NULL,note TEXT NOT NULL DEFAULT '',reference TEXT NOT NULL DEFAULT '',
          request_key TEXT NOT NULL UNIQUE,payload_hash TEXT NOT NULL,input_hash TEXT,related_id INTEGER REFERENCES documents(id),
          reversed_by INTEGER REFERENCES documents(id),created_at TEXT NOT NULL);
        CREATE UNIQUE INDEX IF NOT EXISTS import_once ON documents(input_hash) WHERE input_hash IS NOT NULL;
        CREATE TABLE IF NOT EXISTS lines(id INTEGER PRIMARY KEY,document_id INTEGER NOT NULL REFERENCES documents(id),
          product_id INTEGER NOT NULL REFERENCES products(id),quantity INTEGER NOT NULL,delta INTEGER NOT NULL,
          disposition TEXT NOT NULL DEFAULT '');
        CREATE INDEX IF NOT EXISTS line_product ON lines(product_id);
        CREATE INDEX IF NOT EXISTS doc_day ON documents(day,id);
        CREATE TABLE IF NOT EXISTS attachments(id INTEGER PRIMARY KEY,document_id INTEGER NOT NULL REFERENCES documents(id),
          name TEXT NOT NULL,mime TEXT NOT NULL,sha256 TEXT NOT NULL,data BLOB NOT NULL,created_at TEXT NOT NULL,
          UNIQUE(document_id,sha256));
        CREATE TABLE IF NOT EXISTS counts(id INTEGER PRIMARY KEY,document_id INTEGER NOT NULL REFERENCES documents(id),
          day TEXT NOT NULL,data_json TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS closures(month TEXT PRIMARY KEY,data_json TEXT NOT NULL,created_at TEXT NOT NULL,operator TEXT NOT NULL);
        ''')
        # executescript puede cerrar la transacción; iniciar otra para la carga atómica.
        if not c.in_transaction: c.execute('BEGIN IMMEDIATE')
        schema=meta(c,'schema')
        if schema and schema!='1': raise AppError('Esta base pertenece a otra versión. No se modificó.')
        if not meta(c,'seeded'):
            if c.execute('SELECT COUNT(*) FROM products').fetchone()[0]:
                raise AppError('Hay un catálogo sin marcador de apertura. Recupera el respaldo; no se recargó inventario.')
            seed=json.loads(zlib.decompress(base64.b64decode(SEED_B64)))
            doc=c.execute('''INSERT INTO documents(kind,day,operator,note,request_key,payload_hash,created_at)
                VALUES ('APERTURA',?,'Inventario físico inicial',?,'apertura-v1','seed',?)''',
                (OPENING_DAY,'Conteo confirmado por Angeluis. Blancos y guiones = agotado. No se importan movimientos históricos.',now_iso())).lastrowid
            c.execute('UPDATE documents SET number=? WHERE id=?',(f'AP-{doc:06}',doc))
            counted=[]
            for p in seed:
                c.execute('INSERT OR IGNORE INTO brands(name,normalized) VALUES (?,?)',(p['brand'],norm(p['brand'])))
                bid=c.execute('SELECT id FROM brands WHERE normalized=?',(norm(p['brand']),)).fetchone()[0]
                pid=c.execute('''INSERT INTO products(code,brand_id,name,normalized,source,note,created_day)
                   VALUES (?,?,?,?,?,?,?)''',(p['code'],bid,p['name'],norm(p['name']),p['source'],p['note'],OPENING_DAY)).lastrowid
                c.execute('INSERT INTO lines(document_id,product_id,quantity,delta) VALUES (?,?,?,?)',(doc,pid,p['qty'],p['qty']))
                counted.append({'id':pid,'code':p['code'],'brand':p['brand'],'name':p['name'],'expected':p['qty'],'counted':p['qty'],'difference':0})
            c.execute('INSERT INTO counts(document_id,day,data_json) VALUES (?,?,?)',(doc,OPENING_DAY,json.dumps(counted,ensure_ascii=False)))
            for k,v in [('seeded','1'),('schema','1'),('revision','1'),('opening_day',OPENING_DAY),('app_version',VERSION),('source',SOURCE_NAME)]: put_meta(c,k,v)
        if not meta(c, 'storage_epoch'):
            put_meta(c, 'storage_epoch', uuid.uuid4().hex)
        put_meta(c, 'app_version', VERSION)
    daily_backup()

def daily_backup():
    dest=DATA_DIR/'respaldos'
    dest.mkdir(exist_ok=True)
    p=dest/f'marketing_{today().isoformat()}.sqlite3'
    if not p.exists():
        with LOCK:
            tmp=dest/f'.respaldo-{uuid.uuid4().hex}.tmp'
            src=connect(); out=sqlite3.connect(tmp)
            try: src.backup(out)
            finally: out.close();src.close()
            os.replace(tmp,p)
        older=sorted(dest.glob('marketing_*.sqlite3'))[:-14]
        for old in older: old.unlink(missing_ok=True)

def check_day(c,value):
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',str(value)): raise AppError('Usa una fecha con formato AAAA-MM-DD.')
    try: d=date.fromisoformat(str(value))
    except (ValueError,TypeError): raise AppError('Fecha inválida.')
    if d<date.fromisoformat(meta(c,'opening_day')): raise AppError('La fecha no puede ser anterior al inventario inicial.')
    if d>today(): raise AppError('No puedes registrar una salida o entrada con fecha futura.')
    latest=c.execute('SELECT MAX(day) FROM documents').fetchone()[0]
    if latest and value<latest:
        raise AppError(f'Ya hay movimientos hasta {latest}. Registra la corrección hoy, sin alterar fechas anteriores.')
    if c.execute('SELECT 1 FROM closures WHERE month=?',(value[:7],)).fetchone():
        raise AppError('Este mes está cerrado. Registra el movimiento en un período abierto.')
    return value

def fingerprint(data):
    return hashlib.sha256(json.dumps(data,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()

def validate_key(value):
    key=text(value,120)
    if not key or len(key)<12: raise AppError('Falta el identificador de la operación. Recarga la pantalla.')
    return key

def already(c,key,payload):
    r=c.execute('SELECT id,payload_hash FROM documents WHERE request_key=?',(key,)).fetchone()
    if r:
        if r['payload_hash']!=payload: raise AppError('La operación ya fue guardada con otros datos. Abre una ficha nueva.')
        return r['id']
    return None

def insert_document(c,kind,day,operator,recipient,note,reference,key,payload,line_items,related=None,input_hash=None):
    if input_hash and c.execute('SELECT number FROM documents WHERE input_hash=?',(input_hash,)).fetchone():
        r=c.execute('SELECT number FROM documents WHERE input_hash=?',(input_hash,)).fetchone()
        raise AppError(f'Este archivo ya fue aplicado en {r[0]}. No se descontará ni sumará otra vez.')
    aggregates=collections.Counter()
    for l in line_items: aggregates[l['product_id']]+=l['delta']
    for pid,change in aggregates.items():
        p=c.execute('SELECT code,name,created_day FROM products WHERE id=?',(pid,)).fetchone()
        if not p: raise AppError('Un artículo ya no existe en el catálogo.')
        if day<p['created_day']: raise AppError(f"{p['code']} fue creado el {p['created_day']}. Registra su movimiento desde esa fecha.")
        available=qty(c,pid)
        if available+change<0: raise AppError(f"Existencia insuficiente en {p['code']} · {p['name']}: hay {available}.")
    did=c.execute('''INSERT INTO documents(kind,day,recipient,operator,note,reference,request_key,payload_hash,input_hash,related_id,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                (kind,day,recipient,operator,note,reference,key,payload,input_hash,related,now_iso())).lastrowid
    c.execute('UPDATE documents SET number=? WHERE id=?',(f'{PREFIX[kind]}-{did:06}',did))
    for l in line_items:
        c.execute('INSERT INTO lines(document_id,product_id,quantity,delta,disposition) VALUES (?,?,?,?,?)',
            (did,l['product_id'],l['quantity'],l['delta'],l.get('disposition','')))
    bump(c)
    return did

def submit_movement(data):
    key=validate_key(data.get('key')); payload=fingerprint(data)
    with transaction() as c:
        previous=already(c,key,payload)
        if previous: return {'id':previous,'reused':True}
        kind=text(data.get('kind'))
        if kind not in ('ENTRADA','ENTREGA','EVENTO','PRESTAMO'): raise AppError('Tipo de movimiento no permitido.')
        day=check_day(c,text(data.get('day')))
        operator=text(data.get('operator')); recipient=text(data.get('recipient'))
        if not operator: raise AppError('Escribe quién registra el movimiento.')
        if kind!='ENTRADA' and not recipient: raise AppError('Escribe el asesor o el nombre del evento.')
        note=text(data.get('note'),2000);reference=text(data.get('reference'))
        raw=data.get('items',[])
        if not isinstance(raw,list) or not 1<=len(raw)<=500: raise AppError('Agrega entre 1 y 500 artículos.')
        totals=collections.Counter()
        for item in raw:
            pid=integer(item.get('product_id'));totals[pid]+=integer(item.get('quantity'))
        for n in totals.values(): integer(n)
        line_items=[{'product_id':pid,'quantity':n,'delta':n if kind=='ENTRADA' else -n} for pid,n in totals.items()]
        ih=text(data.get('import_hash')) or None
        if ih:
            if not re.fullmatch('[0-9a-f]{64}',ih): raise AppError('Identificador del archivo inválido.')
            ih=kind+':'+ih
        did=insert_document(c,kind,day,operator,recipient,note,reference,key,payload,line_items,input_hash=ih)
    return {'id':did,'reused':False}

def product_list(c):
    return rows(c,'''SELECT p.id,p.code,p.name,b.id AS brand_id,b.name AS brand,p.source,p.note,p.created_day,
       COALESCE(SUM(l.delta),0) AS stock FROM products p JOIN brands b ON b.id=p.brand_id
       LEFT JOIN lines l ON l.product_id=p.id GROUP BY p.id ORDER BY b.name,p.name''')

def event_lines(c,eid,until=None):
    e=c.execute('SELECT * FROM documents WHERE id=?',(eid,)).fetchone()
    if not e or e['kind'] not in ('EVENTO','PRESTAMO') or e['reversed_by']:
        raise AppError('La salida de evento no está disponible.')
    result=rows(c,'''SELECT p.id AS product_id,p.code,p.name,b.name AS brand,SUM(l.quantity) AS dispatched
      FROM lines l JOIN products p ON p.id=l.product_id JOIN brands b ON b.id=p.brand_id
      WHERE l.document_id=? GROUP BY p.id ORDER BY p.code''',(eid,))
    for item in result:
        a=c.execute('''SELECT COALESCE(SUM(l.quantity),0),COALESCE(SUM(l.delta),0)
          FROM lines l JOIN documents d ON d.id=l.document_id
          WHERE d.kind='LIQUIDACION' AND d.related_id=? AND d.reversed_by IS NULL
          AND l.product_id=?''',(eid,item['product_id'])).fetchone()
        item['resolved']=a[0];item['returned']=a[1];item['pending']=item['dispatched']-a[0]
    return dict(e),result

def submit_resolution(data):
    key=validate_key(data.get('key')); payload=fingerprint(data)
    with transaction() as c:
        previous=already(c,key,payload)
        if previous: return {'id':previous,'reused':True}
        eid=integer(data.get('event_id')); event,els=event_lines(c,eid)
        pending={r['product_id']:r['pending'] for r in els}
        day=check_day(c,text(data.get('day')));operator=text(data.get('operator'));note=text(data.get('note'),2000)
        if not operator: raise AppError('Escribe quién recibe o confirma la mercancía.')
        items=[];seen=set()
        for r in data.get('items',[]):
            pid=integer(r.get('product_id'))
            if pid in seen or pid not in pending: raise AppError('Artículo repetido o ajeno a la salida.')
            seen.add(pid)
            returned=integer(r.get('returned',0),zero=True); delivered=integer(r.get('delivered',0),zero=True)
            if returned+delivered>pending[pid]: raise AppError('No se puede devolver o confirmar más de lo pendiente.')
            if delivered and event['kind']=='PRESTAMO':
                raise AppError('En un préstamo solo se permite regresar material. Una pérdida necesita revisión y ajuste separado.')
            if returned: items.append({'product_id':pid,'quantity':returned,'delta':returned,'disposition':'REGRESO'})
            if delivered: items.append({'product_id':pid,'quantity':delivered,'delta':0,'disposition':'ENTREGADO'})
        if not items: raise AppError('Introduce al menos una cantidad para regresar o confirmar.')
        did=insert_document(c,'LIQUIDACION',day,operator,event['recipient'],note,event['number'],key,payload,items,related=eid)
    return {'id':did,'reused':False}

def cancel_document(data):
    key=validate_key(data.get('key'));payload=fingerprint(data)
    with transaction() as c:
        previous=already(c,key,payload)
        if previous:return {'id':previous,'reused':True}
        original=c.execute('SELECT * FROM documents WHERE id=?',(integer(data.get('id')),)).fetchone()
        if not original or original['kind'] in ('APERTURA','AJUSTE','ANULACION') or original['reversed_by']:
            raise AppError('Este documento no puede anularse. Para diferencias de conteo usa Conteo físico.')
        if c.execute('SELECT 1 FROM closures WHERE month=?',(original['day'][:7],)).fetchone():
            raise AppError('El documento pertenece a un mes cerrado. No se altera; registra un ajuste mediante conteo en el mes abierto.')
        if c.execute("SELECT 1 FROM documents WHERE related_id=? AND reversed_by IS NULL AND kind!='ANULACION'",(original['id'],)).fetchone():
            raise AppError('Primero anula los regresos o cierres asociados a esta salida.')
        day=check_day(c,today().isoformat());note=text(data.get('note'),2000);operator=text(data.get('operator'))
        if len(note)<8 or not operator: raise AppError('Indica responsable y un motivo claro para anular.')
        ls=rows(c,'SELECT product_id,quantity,delta,disposition FROM lines WHERE document_id=?',(original['id'],))
        for l in ls:l['delta']=-l['delta'];l['quantity']=-l['quantity']
        did=insert_document(c,'ANULACION',day,operator,original['recipient'],note,original['number'],key,payload,ls,original['id'])
        c.execute('UPDATE documents SET reversed_by=? WHERE id=?',(did,original['id']))
    return {'id':did,'reused':False}

def add_product(data):
    brand=text(data.get('brand'));name=text(data.get('name'))
    if not brand or not name: raise AppError('Escribe marca/proveedor y artículo.')
    with transaction() as c:
        c.execute('INSERT OR IGNORE INTO brands(name,normalized) VALUES (?,?)',(brand.upper(),norm(brand)))
        bid=c.execute('SELECT id FROM brands WHERE normalized=?',(norm(brand),)).fetchone()[0]
        old=c.execute('SELECT code FROM products WHERE brand_id=? AND normalized=?',(bid,norm(name))).fetchone()
        if old: raise AppError(f'El artículo ya existe con código {old[0]}. Registra una entrada; no lo crees otra vez.')
        seq=max([int(r[0].split('-')[1]) for r in c.execute('SELECT code FROM products')]+[0])+1
        code=f'MKT-{seq:04}'
        pid=c.execute('INSERT INTO products(code,brand_id,name,normalized,created_day,note) VALUES (?,?,?,?,?,?)',
            (code,bid,name.upper(),norm(name),today().isoformat(),'Alta manual sin existencia inicial.')).lastrowid
        bump(c)
    return {'id':pid,'code':code}

def document_data(did):
    with reading() as c:
        doc=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone()
        if not doc:raise AppError('Documento no encontrado.')
        d=dict(doc)
        d['items']=rows(c,'''SELECT l.*,p.code,p.name,b.name AS brand,b.id AS brand_id
             FROM lines l JOIN products p ON p.id=l.product_id JOIN brands b ON b.id=p.brand_id
             WHERE l.document_id=? ORDER BY l.id''',(did,))
        d['attachments']=rows(c,'SELECT id,name,mime,created_at FROM attachments WHERE document_id=?',(did,))
        return d

def snapshot():
    with reading() as c:
        products=product_list(c)
        brands=rows(c,'SELECT id,name FROM brands ORDER BY name')
        docs=rows(c,'''SELECT d.*,COALESCE(SUM(l.quantity),0) AS total,COUNT(l.id) AS item_count,
              (SELECT COUNT(*) FROM attachments a WHERE a.document_id=d.id) AS attachments
              FROM documents d LEFT JOIN lines l ON l.document_id=d.id GROUP BY d.id ORDER BY d.id DESC LIMIT 400''')
        events=[]
        for e in rows(c,"SELECT id FROM documents WHERE kind IN ('EVENTO','PRESTAMO') AND reversed_by IS NULL ORDER BY id DESC"):
            h,ls=event_lines(c,e['id'])
            h['items']=ls;h['pending']=sum(x['pending'] for x in ls)
            if h['pending']>0:events.append(h)
        return {'version':VERSION,'today':today().isoformat(),'opening_day':meta(c,'opening_day'),
            'revision':int(meta(c,'revision')),'source':SOURCE_NAME,'products':products,'brands':brands,'documents':docs,'events':events,
            'closures':rows(c,'SELECT month,operator,created_at FROM closures ORDER BY month DESC'),
            'operators':[r[0] for r in c.execute("SELECT DISTINCT operator FROM documents WHERE kind!='APERTURA' ORDER BY operator")],
            'recipients':[r[0] for r in c.execute("SELECT DISTINCT recipient FROM documents WHERE recipient!='' ORDER BY recipient")],
            'data_path':'SQLite en el servidor · no es almacenamiento externo garantizado','storage_epoch':meta(c,'storage_epoch'),'last_count':dict(c.execute('SELECT document_id,day FROM counts ORDER BY id DESC LIMIT 1').fetchone())}

def save_count(data):
    key=validate_key(data.get('key'));payload=fingerprint(data)
    with transaction() as c:
        previous=already(c,key,payload)
        if previous:return {'id':previous,'reused':True}
        if integer(data.get('revision'),zero=True)!=int(meta(c,'revision')):
            raise AppError('Hubo cambios desde que abriste el conteo. Descarga tu borrador y vuelve a abrir el conteo para comparar con el saldo actualizado.')
        day=check_day(c,text(data.get('day')));operator=text(data.get('operator'));note=text(data.get('note'),2000)
        if not operator:raise AppError('Escribe quién realizó el conteo.')
        plist=product_list(c);actual={p['id']:p for p in plist};found={};saved=[];adjust=[]
        for r in data.get('items',[]):
            pid=integer(r.get('product_id'))
            if pid in found:raise AppError('Hay artículos repetidos en el conteo.')
            found[pid]=integer(r.get('counted'),zero=True)
        if set(actual)!=set(found):raise AppError('El conteo debe incluir todos los artículos, también los agotados. Usa 0 cuando hayas contado cero.')
        for pid,p in actual.items():
            n=found[pid];dif=n-p['stock']
            saved.append({'id':pid,'code':p['code'],'brand':p['brand'],'name':p['name'],'expected':p['stock'],'counted':n,'difference':dif})
            if dif:adjust.append({'product_id':pid,'quantity':abs(dif),'delta':dif,'disposition':'AJUSTE_CONTEO'})
        if adjust and len(note)<8:raise AppError('Hay diferencias: explica el motivo antes de confirmar el ajuste.')
        did=insert_document(c,'AJUSTE',day,operator,'Conteo físico',note,'Conteo completo',key,payload,adjust)
        c.execute('INSERT INTO counts(document_id,day,data_json) VALUES (?,?,?)',(did,day,json.dumps(saved,ensure_ascii=False)))
    return {'id':did,'reused':False,'differences':len(adjust)}

def month_bounds(month):
    if not re.fullmatch(r'20\d{2}-(0[1-9]|1[0-2])',str(month)):raise AppError('Mes inválido.')
    y,m=map(int,month.split('-'))
    return date(y,m,1).isoformat(),date(y,m,calendar.monthrange(y,m)[1]).isoformat()

def report_data(c,month):
    cached=c.execute('SELECT data_json FROM closures WHERE month=?',(month,)).fetchone()
    if cached:
        r=json.loads(cached[0]);r['status']='CERRADO';return r
    start,end=month_bounds(month)
    if month<meta(c,'opening_day')[:7] or start>today().isoformat():raise AppError('No hay inventario inicial para ese período.')
    cutoff=min(end,today().isoformat())
    plist=rows(c,'''SELECT p.id,p.code,p.name,b.name AS brand,b.id AS brand_id
                  FROM products p JOIN brands b ON b.id=p.brand_id WHERE p.created_day<=? ORDER BY b.name,p.name''',(cutoff,))
    lookup={p['id']:p for p in plist}
    for p in plist:
        for k in ['initial','incoming','deliveries','event_out','returns','adjustments','final','event_delivered','pending']:p[k]=0
    ledger=rows(c,'''SELECT l.*,d.number,d.kind,d.day,d.operator,d.recipient,d.note,d.reference,d.related_id,d.reversed_by,
                      p.code,p.name,b.name AS brand,b.id AS brand_id
             FROM lines l JOIN documents d ON d.id=l.document_id JOIN products p ON p.id=l.product_id
             JOIN brands b ON b.id=p.brand_id WHERE d.day<=? ORDER BY d.day,d.id,l.id''',(cutoff,))
    movements=[]
    for l in ledger:
        p=lookup.get(l['product_id'])
        if not p:continue
        if l['day']<start or l['kind']=='APERTURA':p['initial']+=l['delta']
        elif l['kind']=='ENTRADA':p['incoming']+=l['delta']
        elif l['kind']=='ENTREGA':p['deliveries']-=l['delta']
        elif l['kind'] in ('EVENTO','PRESTAMO'):p['event_out']-=l['delta']
        elif l['kind']=='LIQUIDACION':
            p['returns']+=l['delta']
            if l['disposition']=='ENTREGADO':p['event_delivered']+=l['quantity']
        elif l['kind'] in ('AJUSTE','ANULACION'):
            p['adjustments']+=l['delta']
            if l['kind']=='ANULACION' and l['disposition']=='ENTREGADO':p['event_delivered']+=l['quantity']
        if start<=l['day']<=cutoff:movements.append(l)
    # Pendiente real a la fecha del corte, incluyendo anulación fechada después del corte.
    pending_events=[]
    for e in rows(c,"SELECT * FROM documents WHERE kind IN ('EVENTO','PRESTAMO') AND day<=?",(cutoff,)):
        if e['reversed_by']:
            rev=c.execute('SELECT day FROM documents WHERE id=?',(e['reversed_by'],)).fetchone()
            if rev and rev[0]<=cutoff:continue
        for l in rows(c,'SELECT product_id,SUM(quantity) AS n FROM lines WHERE document_id=? GROUP BY product_id',(e['id'],)):
            settled=c.execute('''SELECT COALESCE(SUM(l.quantity),0) FROM lines l JOIN documents d ON d.id=l.document_id
               LEFT JOIN documents r ON r.id=d.reversed_by WHERE d.kind='LIQUIDACION' AND d.related_id=?
               AND d.day<=? AND (r.id IS NULL OR r.day>?) AND l.product_id=?''',(e['id'],cutoff,cutoff,l['product_id'])).fetchone()[0]
            pending=l['n']-settled
            if pending and l['product_id'] in lookup:
                p=lookup[l['product_id']];p['pending']+=pending
                pending_events.append({'number':e['number'],'recipient':e['recipient'],'kind':e['kind'],'code':p['code'],'name':p['name'],'brand':p['brand'],'brand_id':p['brand_id'],'pending':pending})
    for p in plist:
        p['final']=p['initial']+p['incoming']-p['deliveries']-p['event_out']+p['returns']+p['adjustments']
    counts=rows(c,'SELECT day,document_id FROM counts WHERE day<=? ORDER BY id DESC LIMIT 1',(cutoff,))
    return {'month':month,'start':start,'end':cutoff,'calendar_end':end,'opening_day':meta(c,'opening_day'),
            'status':'EN CURSO' if today().isoformat()<end else 'SIN CERRAR','generated_at':now_iso(),
            'products':plist,'movements':movements,'events':pending_events,'count':counts[0] if counts else None,
            'note':'Las salidas a eventos descuentan almacén; confirmar lo entregado en el evento no vuelve a descontar. Los regresos no son entradas del proveedor. Las anulaciones se muestran como ajustes.'}

def filtered_report(report,brand_id=None):
    r=json.loads(json.dumps(report))
    r['scope']='Todas las marcas / gerencia'
    if brand_id:
        bid=integer(brand_id)
        r['products']=[p for p in r['products'] if p['brand_id']==bid]
        r['movements']=[p for p in r['movements'] if p['brand_id']==bid]
        r['events']=[p for p in r['events'] if p['brand_id']==bid]
        if not r['products']:raise AppError('No hay artículos de esa marca en el período.')
        r['scope']=r['products'][0]['brand']
    r['brand_id']=int(brand_id) if brand_id else None
    keys=['initial','incoming','deliveries','event_out','returns','adjustments','final','event_delivered','pending']
    r['totals']={k:sum(p[k] for p in r['products']) for k in keys}
    return r

def close_month(data):
    month=text(data.get('month'));start,end=month_bounds(month);operator=text(data.get('operator'))
    if not operator:raise AppError('Escribe el responsable del cierre.')
    if today().isoformat()<end:raise AppError('El mes todavía está en curso. Puedes descargar el reporte sin cerrarlo.')
    with transaction() as c:
        if c.execute('SELECT 1 FROM closures WHERE month=?',(month,)).fetchone():return {'month':month,'reused':True}
        r=report_data(c,month)
        cnt=c.execute('SELECT * FROM counts WHERE day BETWEEN ? AND ? ORDER BY id DESC LIMIT 1',(start,end)).fetchone()
        if not cnt:raise AppError('Falta un conteo físico completo de este mes.')
        changed=c.execute('''SELECT 1 FROM documents d JOIN lines l ON l.document_id=d.id
                   WHERE d.id>? AND d.day<=? AND l.delta!=0 LIMIT 1''',(cnt['document_id'],end)).fetchone()
        if changed:raise AppError('Hay movimientos posteriores al conteo. Realiza un nuevo conteo antes de cerrar.')
        counted={x['id'] for x in json.loads(cnt['data_json'])}
        if counted!={p['id'] for p in r['products']}:raise AppError('Hay artículos nuevos pendientes de incluir en el conteo.')
        r['status']='CERRADO';r['closed_by']=operator;r['closed_at']=now_iso()
        c.execute('INSERT INTO closures(month,data_json,created_at,operator) VALUES (?,?,?,?)',
                  (month,json.dumps(r,ensure_ascii=False),now_iso(),operator));bump(c)
    return {'month':month,'reused':False}

def decode_upload(data):
    name=text(data.get('name'),180)
    try:b=base64.b64decode(data.get('data',''),validate=True)
    except Exception:raise AppError('No se pudo leer el archivo.')
    if not b or len(b)>MAX_UPLOAD:raise AppError('El archivo está vacío o supera 12 MB.')
    return name,b

def import_rows(data):
    name,b=decode_upload(data);ext=Path(name).suffix.lower()
    if ext=='.csv':
        try:s=b.decode('utf-8-sig')
        except UnicodeDecodeError:s=b.decode('cp1252')
        try:delim=csv.Sniffer().sniff(s[:8192],delimiters=',;\t').delimiter
        except csv.Error:delim=';'
        rr=list(csv.reader(io.StringIO(s),delimiter=delim))
    elif ext=='.xlsx':
        try:
            with zipfile.ZipFile(io.BytesIO(b)) as z:
                if sum(i.file_size for i in z.infolist())>60*1024*1024:raise AppError('El Excel descomprimido es demasiado grande.')
            from openpyxl import load_workbook
            wb=load_workbook(io.BytesIO(b),read_only=True,data_only=False,keep_links=False)
            sh=wb.worksheets[0]
            if sh.max_row and sh.max_row>10002:raise AppError('La plantilla admite hasta 10.000 filas.')
            rr=[]
            for n,row in enumerate(sh.iter_rows(values_only=True)):
                if n>10001:raise AppError('Demasiadas filas.')
                rr.append(list(row[:30]))
            wb.close()
        except AppError:raise
        except ImportError:raise AppError('Instala openpyxl para importar Excel.')
        except Exception:raise AppError('No se pudo abrir ese Excel. Usa la plantilla descargable de la app.')
    else:raise AppError('Para cargar cantidades usa Excel .xlsx o CSV; una foto firmada es solo un respaldo.')
    rr=[r for r in rr if any(v is not None and str(v).strip() for v in r)]
    if len(rr)<2 or len(rr)>10001:raise AppError('No hay filas de datos o el archivo supera 10.000 filas.')
    headers=[norm(x or '') for x in rr[0]]
    def idx(options):return next((i for i,h in enumerate(headers) if h in options),None)
    ci=idx({'CODIGO','CODIGO INTERNO'});qi=idx({'CANTIDAD','CONTADO','CONTEO'});bi=idx({'MARCA','PROVEEDOR','MARCA / PROVEEDOR'});ni=idx({'ARTICULO','PRODUCTO','DESCRIPCION'})
    if qi is None or (ci is None and (bi is None or ni is None)):
        raise AppError('Se necesitan columnas Código y Cantidad, o Marca, Artículo y Cantidad. Descarga la plantilla.')
    get=lambda row,i: row[i] if i is not None and i<len(row) else None
    is_count=data.get('mode')=='count'
    with reading() as c:plist=product_list(c)
    bycode={p['code']:p for p in plist};byname={(norm(p['brand']),norm(p['name'])):p for p in plist}
    amounts=collections.Counter();errors=[];seen=set()
    for num,row in enumerate(rr[1:],2):
        q=get(row,qi)
        if q is None or str(q).strip()=='':continue
        code=text(get(row,ci)).upper() if ci is not None else ''
        p=bycode.get(code) if code else byname.get((norm(get(row,bi) or ''),norm(get(row,ni) or '')))
        if not p:errors.append(f'Fila {num}: artículo no identificado. Crea el artículo o revisa su código.');continue
        # Si la fila trae marca y nombre, no ignorar una contradicción con el código.
        bval=get(row,bi);nval=get(row,ni)
        if bval and norm(bval)!=norm(p['brand']):errors.append(f'Fila {num}: el código no corresponde a esa marca.');continue
        if nval and norm(nval)!=norm(p['name']):errors.append(f'Fila {num}: el código no corresponde a ese artículo.');continue
        try:n=integer(q,zero=is_count)
        except AppError as ex:errors.append(f'Fila {num}: {ex}');continue
        if is_count and p['id'] in seen:errors.append(f'Fila {num}: artículo repetido en conteo.');continue
        seen.add(p['id']);amounts[p['id']]+=n
    if errors:raise AppError('\n'.join(errors[:15])+ ('\nHay más errores; no se importó ninguna fila.' if len(errors)>15 else '\nNo se importó ninguna fila.'))
    if not amounts:raise AppError('No se encontraron cantidades para importar.')
    mapped=[{'product_id':p['id'],'code':p['code'],'name':p['name'],'brand':p['brand'],'quantity':amounts[p['id']]} for p in plist if p['id'] in amounts]
    return {'items':mapped,'hash':hashlib.sha256(b).hexdigest(),'name':name,'rows':len(mapped)}

def attach_file(data):
    name,b=decode_upload(data);did=integer(data.get('document_id'));ext=Path(name).suffix.lower()
    if ext=='.pdf' and b.startswith(b'%PDF-'):mime='application/pdf'
    elif ext in ('.jpg','.jpeg','.png'):
        try:
            from PIL import Image
            image=Image.open(io.BytesIO(b))
            if image.width*image.height>16_000_000: raise ValueError('Imagen demasiado grande')
            image.verify()
        except Exception:raise AppError('La imagen no es válida.')
        mime='image/png' if ext=='.png' else 'image/jpeg'
    else:raise AppError('Adjunta una planilla PDF, JPG o PNG válida.')
    sha=hashlib.sha256(b).hexdigest()
    with transaction() as c:
        if not c.execute('SELECT 1 FROM documents WHERE id=?',(did,)).fetchone():raise AppError('Documento no encontrado.')
        c.execute('INSERT OR IGNORE INTO attachments(document_id,name,mime,sha256,data,created_at) VALUES (?,?,?,?,?,?)',
                  (did,Path(name).name,mime,sha,b,now_iso()))
    return {'saved':True,'stock_changed':False}

def save_logo(data):
    name,b=decode_upload(data);bid=integer(data.get('brand_id'),zero=True)
    try:
        from PIL import Image
        im=Image.open(io.BytesIO(b))
        if im.width*im.height>16_000_000:raise AppError('El logo es demasiado grande.')
        im.load()
        im=im.convert('RGBA');im.thumbnail((800,400));out=io.BytesIO();im.save(out,format='PNG');clean=out.getvalue()
    except AppError:raise
    except Exception:raise AppError('Sube el logo como PNG o JPG.')
    with transaction() as c:
        if bid==0:put_meta(c,'company_logo',base64.b64encode(clean).decode())
        else:
            if not c.execute('SELECT 1 FROM brands WHERE id=?',(bid,)).fetchone():raise AppError('Marca no encontrada.')
            c.execute('UPDATE brands SET logo=? WHERE id=?',(clean,bid))
    return {'saved':True}

def get_logo(brand_id=None):
    with reading() as c:
        if brand_id:
            r=c.execute('SELECT logo FROM brands WHERE id=?',(brand_id,)).fetchone()
            if r and r[0]:return r[0]
            return None  # No sustituir un logo de proveedor faltante por otro.
        b=meta(c,'company_logo')
        return base64.b64decode(b) if b else None

# ---------------------------------------------------------------------------
# Exportadores. Las dependencias se cargan solo cuando se solicitan archivos.
# ---------------------------------------------------------------------------
def xlsx_report(r):
    try:import xlsxwriter
    except ImportError:raise AppError('Instala XlsxWriter para generar Excel.')
    out=io.BytesIO();w=xlsxwriter.Workbook(out,{'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    title=w.add_format({'bold':True,'font_size':20,'font_color':'#32234E'})
    sub=w.add_format({'font_color':'#756783','font_size':10})
    head=w.add_format({'bg_color':'#6650A4','font_color':'#FFFFFF','bold':True,'text_wrap':True,'valign':'vcenter'})
    num=w.add_format({'num_format':'#,##0','valign':'top'})
    wrap=w.add_format({'text_wrap':True,'valign':'top'})
    muted=w.add_format({'bg_color':'#F4F0FA','font_color':'#58456F','text_wrap':True,'valign':'vcenter'})
    big=w.add_format({'bold':True,'font_size':21,'font_color':'#6650A4','num_format':'#,##0'})
    sumfmt=w.add_format({'bold':True,'num_format':'#,##0','bg_color':'#EDE5F8'})
    names=['Código','Marca / proveedor','Artículo','Inicial / apertura','Entradas proveedor','Entregas asesores',
           'Salidas eventos / préstamos','Regresos','Ajustes / anulaciones','Saldo en almacén','Fuera pendiente','Entregado en eventos']
    detail=w.add_worksheet('Inventario mensual');detail.freeze_panes(5,3);detail.hide_gridlines(2)
    detail.merge_range('A1:L1',f"INVENTARIO MENSUAL · {r['scope']}",title)
    detail.merge_range('A2:L2',f"{r['month']} · {r['status']} · Corte: {r['end']}",sub)
    detail.merge_range('A3:L3','Inicial incluye la apertura cuando corresponde. Regresos no son entradas del proveedor. No sumar entregados en eventos como otra salida.',muted)
    detail.set_row(2,32);detail.set_row(4,36)
    detail.set_column('A:A',14);detail.set_column('B:B',19);detail.set_column('C:C',45,wrap);detail.set_column('D:L',15,num)
    detail.write_row(4,0,names,head)
    for i,p in enumerate(r['products'],5):
        detail.write_row(i,0,[p['code'],p['brand'],p['name']],wrap)
        vals=[p[k] for k in ['initial','incoming','deliveries','event_out','returns','adjustments']]
        detail.write_row(i,3,vals,num)
        er=i+1;detail.write_formula(i,9,f'=D{er}+E{er}-F{er}-G{er}+H{er}+I{er}',num,p['final'])
        detail.write_number(i,10,p['pending'],num);detail.write_number(i,11,p['event_delivered'],num)
        detail.set_row(i,30 if len(p['name'])>42 else 21)
    last=5+len(r['products'])
    detail.autofilter(4,0,max(4,last-1),11)
    detail.write(last,2,'TOTAL',sumfmt)
    for col,key in enumerate(['initial','incoming','deliveries','event_out','returns','adjustments','final','pending','event_delivered'],3):
        letter=chr(65+col)
        detail.write_formula(last,col,f'=SUM({letter}6:{letter}{last})',sumfmt,r['totals'][key])
    detail.conditional_format(5,9,max(5,last-1),9,{'type':'cell','criteria':'==','value':0,'format':w.add_format({'font_color':'#A3607A','bg_color':'#FCEEF3'})})
    detail.set_landscape();detail.fit_to_pages(1,0);detail.repeat_rows(4);detail.set_paper(9)
    summary=w.add_worksheet('Resumen');summary.hide_gridlines(2);summary.set_column('A:J',13)
    summary.merge_range('A1:J2','MARKETING · CONTROL MENSUAL',title)
    summary.merge_range('A3:J3',f"{r['scope']} | {r['month']} | {r['status']}",sub)
    totals=r['totals'];blocks=[('A5:C5','A6:C7','Existencia final',totals['final']),('D5:F5','D6:F7','Entradas proveedor',totals['incoming']),('G5:J5','G6:J7','Salidas del almacén',totals['deliveries']+totals['event_out'])]
    for a,b,label,value in blocks:summary.merge_range(a,label,muted);summary.merge_range(b,value,big)
    summary.merge_range('A9:J10',f"Regresaron: {totals['returns']:,} · Fuera pendientes: {totals['pending']:,} · Ajustes netos: {totals['adjustments']:,}".replace(',','.'),muted)
    summary.merge_range('A12:J14',r['note']+f" Apertura: {r['opening_day']}. No reconstruye movimientos anteriores a esa fecha.",wrap)
    summary.merge_range('A16:J16','EXISTENCIAS POR MARCA' if not r['brand_id'] else 'ARTÍCULOS CON MAYOR EXISTENCIA',head)
    grouped=collections.Counter()
    for p in r['products']:grouped[p['name'] if r['brand_id'] else p['brand']]+=p['final']
    chartdata=w.add_worksheet('Datos gráfica');chartdata.write_row(0,0,['Concepto','Existencia'],head);chartdata.set_column('A:A',45);chartdata.set_column('B:B',16,num)
    for i,(label,value) in enumerate(grouped.most_common(10),1):chartdata.write_row(i,0,[label,value])
    if grouped:
        chart=w.add_chart({'type':'bar'})
        chart.add_series({'name':'Unidades en almacén','categories':['Datos gráfica',1,0,min(10,len(grouped)),0],
                          'values':['Datos gráfica',1,1,min(10,len(grouped)),1], 'fill':{'color':'#8D73C7'},'border':{'none':True},'data_labels':{'value':True}})
        chart.set_legend({'none':True});chart.set_x_axis({'major_gridlines':{'visible':True,'line':{'color':'#EDE8F4'}}});chart.set_y_axis({'reverse':True})
        chart.set_chartarea({'border':{'none':True}});chart.set_plotarea({'border':{'none':True}});chart.set_size({'width':780,'height':380})
        summary.insert_chart('A18',chart)
    summary.set_landscape();summary.set_paper(9);summary.fit_to_pages(1,1);summary.print_area('A1:J38')
    moves=w.add_worksheet('Movimientos');moves.freeze_panes(1,0);moves.hide_gridlines(2)
    headers=['Fecha','Ficha','Tipo','Marca / proveedor','Código','Artículo','Cantidad','Impacto en almacén','Destino / asesor','Registrado por','Disposición','Referencia','Observaciones']
    moves.write_row(0,0,headers,head);moves.set_row(0,34)
    moves.set_column('A:B',15);moves.set_column('C:D',23);moves.set_column('E:E',14);moves.set_column('F:F',43,wrap);moves.set_column('G:H',17,num);moves.set_column('I:J',23);moves.set_column('K:M',30,wrap)
    for i,m in enumerate(r['movements'],1):
        moves.write_row(i,0,[m['day'],m['number'],KINDS[m['kind']],m['brand'],m['code'],m['name'],m['quantity'],m['delta'],m['recipient'],m['operator'],m['disposition'],m['reference'],m['note']])
    moves.autofilter(0,0,max(1,len(r['movements'])),12)
    events=w.add_worksheet('Fuera del almacén');events.hide_gridlines(2);events.freeze_panes(1,0)
    events.write_row(0,0,['Ficha','Evento / destino','Tipo','Marca','Código','Artículo','Pendiente'],head)
    events.set_column('A:A',15);events.set_column('B:D',26);events.set_column('E:E',14);events.set_column('F:F',45,wrap);events.set_column('G:G',15,num)
    for i,e in enumerate(r['events'],1):events.write_row(i,0,[e['number'],e['recipient'],KINDS[e['kind']],e['brand'],e['code'],e['name'],e['pending']])
    audit=w.add_worksheet('Lectura y alcance');audit.set_column('A:A',26);audit.set_column('B:B',105,wrap)
    notes=[['Versión',VERSION],['Origen apertura',SOURCE_NAME],['Fecha de apertura',r['opening_day']],['Regla confirmada','Blancos y guiones del inventario inicial = cero; artículos conservados.'],['Estado del mes',r['status']],['Corte',r['end']],['Conteo físico',json.dumps(r.get('count'),ensure_ascii=False)],['Anulaciones','El original conserva su evidencia. La contrapartida se identifica como ajuste en el mes de anulación.'],['Límite','No contiene historial anterior al inventario inicial.'],['Eventos','Salida resta una vez; confirmar entrega no vuelve a restar; regreso suma únicamente al proveedor original.']]
    for i,row in enumerate(notes):audit.write_row(i,0,row);audit.set_row(i,35)
    summary.activate();w.close();return out.getvalue()

def excel_template(mode='movement'):
    import xlsxwriter
    with reading() as c:plist=product_list(c)
    out=io.BytesIO();w=xlsxwriter.Workbook(out,{'in_memory':True,'strings_to_formulas':False,'strings_to_urls':False})
    s=w.add_worksheet('Carga');h=w.add_format({'bold':True,'bg_color':'#6650A4','font_color':'white'});q=w.add_format({'bg_color':'#FFF6DB','num_format':'0'})
    s.write_row(0,0,['Codigo','Marca','Articulo','Contado' if mode=='count' else 'Cantidad'],h)
    s.set_column('A:A',15);s.set_column('B:B',20);s.set_column('C:C',60);s.set_column('D:D',15,q);s.freeze_panes(1,3);s.autofilter(0,0,len(plist),3)
    for i,p in enumerate(plist,1):s.write_row(i,0,[p['code'],p['brand'],p['name']]);s.write_blank(i,3,None,q)
    s.data_validation(1,3,len(plist),3,{'validate':'integer','criteria':'between','minimum':0 if mode=='count' else 1,'maximum':MAX_QTY,'input_title':'Cantidad contada' if mode=='count' else 'Cantidad del movimiento','input_message':'Completa solo cantidades enteras.','error_message':'Usa un número entero válido.'})
    n=w.add_worksheet('Instrucciones');n.set_column('A:A',110)
    instructions=['PLANTILLA DE CONTEO COMPLETO' if mode=='count' else 'PLANTILLA DE ENTRADA O SALIDA',
       'No cambies códigos ni marca. Los productos nuevos se crean primero en la app.',
       'En conteo: completa todos los artículos; cero es cero, vacío es pendiente.' if mode=='count' else 'En movimientos: llena solo los artículos que entran/salen; deja los demás en blanco.',
       'Las cantidades se suman o descuentan únicamente después de revisar y confirmar en la app.',
       'Esta plantilla no interpreta automáticamente planillas fotografiadas ni el Excel comparativo original.',
       'Los encabezados deben estar en la primera fila. No agregues filas de totales.']
    for i,t in enumerate(instructions):n.write(i,0,t)
    w.close();return out.getvalue()

def pdf_setup():
    try:
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, Image
        from reportlab.lib.pagesizes import A4, landscape
    except ImportError:raise AppError('Instala reportlab y Pillow para generar PDF.')
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name='MPTitle',fontName='Helvetica-Bold',fontSize=21,leading=25,textColor=colors.HexColor('#32234E'),spaceAfter=12))
    styles.add(ParagraphStyle(name='MPSub',fontName='Helvetica',fontSize=9,leading=13,textColor=colors.HexColor('#72647F'),spaceAfter=10))
    styles.add(ParagraphStyle(name='MPCell',fontName='Helvetica',fontSize=8,leading=11,textColor=colors.HexColor('#30233F')))
    styles.add(ParagraphStyle(name='MPHead',fontName='Helvetica-Bold',fontSize=8,leading=10,textColor=colors.white))
    styles.add(ParagraphStyle(name='MPNote',fontName='Helvetica',fontSize=9,leading=13,textColor=colors.HexColor('#58456F')))
    return locals()

def pdf_table_style(colors):
    return [('BACKGROUND',(0,0),(-1,0),colors.HexColor('#6650A4')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F7F4FB')]),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,1),(-1,-1),'Helvetica'),
            ('FONTSIZE',(0,0),(-1,-1),8),('VALIGN',(0,0),(-1,-1),'TOP'),
            ('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),
            ('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),
            ('LINEBELOW',(0,0),(-1,0),.5,colors.HexColor('#D8CCE9'))]

def footer(canvas,doc):
    canvas.saveState();canvas.setFont('Helvetica',8);canvas.setFillColorRGB(.45,.39,.5)
    canvas.drawString(34,21,'MAKROPETROL · Almacén de marketing · '+VERSION)
    canvas.drawRightString(doc.pagesize[0]-34,21,f'Página {doc.page}');canvas.restoreState()

def pdf_logo(logo,Image):
    if not logo:return []
    from reportlab.lib.utils import ImageReader
    w,h=ImageReader(io.BytesIO(logo)).getSize();scale=min(130/w,45/h)
    return [Image(io.BytesIO(logo),width=w*scale,height=h*scale,hAlign='LEFT')]

def receipt_pdf(d):
    e=pdf_setup();P=e['Paragraph'];T=e['Table'];S=e['Spacer'];styles=e['styles'];C=e['colors']
    out=io.BytesIO();doc=e['SimpleDocTemplate'](out,pagesize=e['A4'],leftMargin=34,rightMargin=34,topMargin=32,bottomMargin=40,
                      title=f"Ficha {d['number']}",author='Makropetrol · Marketing')
    story=pdf_logo(get_logo(),e['Image'])+[S(1,10),P(html.escape(KINDS[d['kind']]).upper(),styles['MPTitle']),
          P(f"<b>{html.escape(d['number'])}</b> · {html.escape(d['day'])} · {'ANULADA' if d['reversed_by'] else 'REGISTRADA'}",styles['MPSub'])]
    info=f"<b>Destino / asesor:</b> {html.escape(d['recipient'] or 'Almacén de marketing')}<br/><b>Registrado por:</b> {html.escape(d['operator'])}<br/><b>Referencia:</b> {html.escape(d['reference'] or 'Sin referencia externa')}"
    story.extend([P(info,styles['MPNote']),S(1,14)])
    head=['Código','Marca','Artículo','Cantidad','Movimiento']
    data=[[P(x,styles['MPHead']) for x in head]]
    for p in d['items']:
        disposition=p['disposition'] or ('Ingreso' if p['delta']>0 else 'Salida' if p['delta']<0 else 'Sin cambio')
        data.append([P(html.escape(str(x)),styles['MPCell']) for x in [p['code'],p['brand'],p['name'],p['quantity'],disposition]])
    if len(data)==1:data.append(['','','Sin diferencias en el conteo','0','Sin cambio'])
    table=T(data,colWidths=[63,70,218,54,122],repeatRows=1);table.setStyle(e['TableStyle'](pdf_table_style(C)))
    story.extend([table,S(1,14)])
    if d['note']:story.extend([P('<b>Observaciones:</b> '+html.escape(d['note']).replace('\n','<br/>'),styles['MPNote']),S(1,16)])
    note='La ficha ya corresponde a un movimiento registrado. Imprimir, descargar o adjuntar su firma no vuelve a modificar inventario.'
    if d['kind']=='PRESTAMO':note+=' Material pendiente de regresar al almacén.'
    signatures=T([[P('_______________________________<br/>Entregó / marketing',styles['MPNote']),P('_______________________________<br/>Recibió / asesor / responsable',styles['MPNote'])]],colWidths=[263,264])
    story.append(e['KeepTogether']([S(1,35),signatures,S(1,22),P(note,styles['MPSub'])]))
    doc.build(story,onFirstPage=footer,onLaterPages=footer);return out.getvalue()

def monthly_pdf(r):
    e=pdf_setup();P=e['Paragraph'];T=e['Table'];S=e['Spacer'];styles=e['styles'];C=e['colors']
    out=io.BytesIO();doc=e['SimpleDocTemplate'](out,pagesize=e['landscape'](e['A4']),leftMargin=34,rightMargin=34,topMargin=26,bottomMargin=38,
                         title=f"Marketing {r['scope']} {r['month']}",author='Makropetrol · Marketing')
    story=pdf_logo(get_logo(r['brand_id']),e['Image'])+[S(1,8),P('REPORTE MENSUAL DE MARKETING',styles['MPTitle']),
             P(f"{html.escape(r['scope'])} · {r['month']} · {r['status']} · Corte {r['end']}",styles['MPSub'])]
    t=r['totals']
    cards=[('EN ALMACÉN',t['final']),('RECIBIDO',t['incoming']),('SALIDAS',t['deliveries']+t['event_out']),('REGRESOS',t['returns']),('FUERA PENDIENTE',t['pending'])]
    ctable=T([[P(f'<font size="8">{a}</font><br/><br/><font size="24"><b>{b:,}</b></font>'.replace(',','.'),styles['MPNote']) for a,b in cards]],colWidths=[154]*5)
    ctable.setStyle(e['TableStyle']([('BACKGROUND',(0,0),(-1,-1),C.HexColor('#F0E9F8')),('TOPPADDING',(0,0),(-1,-1),12),('BOTTOMPADDING',(0,0),(-1,-1),13),('LEFTPADDING',(0,0),(-1,-1),12)]))
    story.extend([ctable,S(1,16)])
    formula=f"Inicial / apertura: <b>{t['initial']:,}</b> + entradas: <b>{t['incoming']:,}</b> - salidas: <b>{t['deliveries']+t['event_out']:,}</b> + regresos: <b>{t['returns']:,}</b> + ajustes: <b>{t['adjustments']:,}</b> = saldo: <b>{t['final']:,}</b>."
    story.extend([P(formula.replace(',','.'),styles['MPNote']),S(1,12)])
    from reportlab.graphics.shapes import Drawing, Rect, String
    grouped=collections.Counter()
    for p in r['products']:grouped[p['name'] if r['brand_id'] else p['brand']]+=p['final']
    top=grouped.most_common(6);drawing=Drawing(770,143)
    maximum=max([n for _,n in top]+[1])
    for i,(label,value) in enumerate(top):
        y=127-i*22;short=label if len(label)<=41 else label[:38]+'...'
        drawing.add(String(0,y+3,short,fontName='Helvetica',fontSize=8,fillColor=C.HexColor('#58456F')))
        drawing.add(Rect(245,y,450*value/maximum,14,fillColor=C.HexColor('#9277C9'),strokeColor=None))
        drawing.add(String(705,y+3,str(value),fontName='Helvetica-Bold',fontSize=9,fillColor=C.HexColor('#58456F')))
    story.extend([P('Mayor existencia por artículo' if r['brand_id'] else 'Mayor existencia por marca',styles['MPNote']),S(1,5),drawing])
    story.extend([P(r['note'],styles['MPSub']),P(f"Apertura del inventario: {r['opening_day']}. No reconstruye entregas anteriores. El reporte incluye únicamente la marca elegida cuando se filtra por proveedor.",styles['MPSub']),e['PageBreak']()])
    story.extend([P('DETALLE DE EXISTENCIAS',styles['MPTitle']),P(html.escape(r['scope'])+' · '+r['month'],styles['MPSub'])])
    heads=['Código / marca','Artículo','Inicial','Entró','Asesor','Evento / préstamo','Regresó','Ajuste','Saldo','Fuera']
    data=[[P(x,styles['MPHead']) for x in heads]]
    for p in r['products']:
        values=[p['code']+' / '+p['brand'],p['name'],p['initial'],p['incoming'],p['deliveries'],p['event_out'],p['returns'],p['adjustments'],p['final'],p['pending']]
        data.append([P(html.escape(str(v)),styles['MPCell']) for v in values])
    totals=['TOTAL','',t['initial'],t['incoming'],t['deliveries'],t['event_out'],t['returns'],t['adjustments'],t['final'],t['pending']]
    data.append([P(str(v),styles['MPCell']) for v in totals])
    table=T(data,colWidths=[91,235,48,48,48,65,50,50,50,85],repeatRows=1)
    cmds=pdf_table_style(C)+[('BACKGROUND',(0,-1),(-1,-1),C.HexColor('#E8DFF4'))]
    table.setStyle(e['TableStyle'](cmds));story.append(table)
    if r['events']:
        story.extend([e['PageBreak'](),P('MATERIAL FUERA DEL ALMACÉN',styles['MPTitle'])])
        edata=[[P(x,styles['MPHead']) for x in ['Ficha','Evento / responsable','Artículo','Marca','Pendiente']]]
        for ev in r['events']:edata.append([P(html.escape(str(x)),styles['MPCell']) for x in [ev['number'],ev['recipient'],ev['name'],ev['brand'],ev['pending']]])
        et=T(edata,colWidths=[85,195,290,100,100],repeatRows=1);et.setStyle(e['TableStyle'](pdf_table_style(C)));story.append(et)
    doc.build(story,onFirstPage=footer,onLaterPages=footer);return out.getvalue()

def all_reports_zip(month):
    with reading() as c:r=report_data(c,month)
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        general=filtered_report(r)
        z.writestr(f'GERENCIA/Marketing_{month}.xlsx',xlsx_report(general));z.writestr(f'GERENCIA/Marketing_{month}.pdf',monthly_pdf(general))
        bids=sorted({p['brand_id'] for p in r['products']})
        for bid in bids:
            part=filtered_report(r,bid);safe=re.sub('[^A-Za-z0-9_-]','_',part['scope'])
            z.writestr(f'PROVEEDORES/{safe}_{month}.xlsx',xlsx_report(part));z.writestr(f'PROVEEDORES/{safe}_{month}.pdf',monthly_pdf(part))
        z.writestr('LEEME.txt','El ZIP es para uso interno. Envía a cada proveedor únicamente sus dos archivos de PROVEEDORES.\nNo envíes el ZIP completo a un proveedor.\n')
    return out.getvalue()


# ---------------------------------------------------------------------------
# Respaldo y recuperación. SQLite compartido entre las sesiones del servidor.
# No se utiliza ninguna cuenta, API key, contraseña ni Secrets de Streamlit.
# ---------------------------------------------------------------------------
MAX_BACKUP_BYTES = 150 * 1024 * 1024
BACKUP_TABLES = ('meta', 'brands', 'products', 'documents', 'lines', 'attachments', 'counts', 'closures')
REQUIRED_COLUMNS = {
    'meta': ('key', 'value'),
    'brands': ('id','name','normalized','logo'),
    'products': ('id','code','brand_id','name','normalized','source','note','created_day'),
    'documents': ('id','number','kind','day','recipient','operator','note','reference','request_key',
                  'payload_hash','input_hash','related_id','reversed_by','created_at'),
    'lines': ('id','document_id','product_id','quantity','delta','disposition'),
    'attachments': ('id','document_id','name','mime','sha256','data','created_at'),
    'counts': ('id','document_id','day','data_json'),
    'closures': ('month','data_json','created_at','operator'),
}


def storage_status():
    """Revisión y generación para impedir escrituras desde una pantalla obsoleta."""
    with reading() as c:
        return {
            'epoch': meta(c, 'storage_epoch'),
            'revision': int(meta(c, 'revision', '0')),
            'documents': c.execute('SELECT COUNT(*) FROM documents').fetchone()[0],
            'references': c.execute('SELECT COUNT(*) FROM products').fetchone()[0],
            'units': c.execute('SELECT COALESCE(SUM(delta),0) FROM lines').fetchone()[0],
            'last_restore': meta(c, 'last_restore'),
        }


def _backup_database_bytes(blob):
    """Acepta respaldos ZIP propios o SQLite. Nunca extrae rutas del archivo."""
    if not isinstance(blob, (bytes, bytearray)) or not 0 < len(blob) <= MAX_BACKUP_BYTES:
        raise AppError('El respaldo está vacío o supera 150 MB.')
    raw = bytes(blob)
    if raw.startswith(b'PK'):
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                entries = z.infolist()
                if len(entries) > 20 or sum(x.file_size for x in entries) > MAX_BACKUP_BYTES:
                    raise AppError('El respaldo descomprimido supera el límite permitido.')
                databases = [x for x in entries if x.filename == 'marketing.sqlite3' and not x.is_dir()]
                if len(databases) != 1 or databases[0].flag_bits & 1:
                    raise AppError('El ZIP debe contener una única base marketing.sqlite3 sin contraseña.')
                raw = z.read(databases[0])
        except AppError:
            raise
        except (zipfile.BadZipFile, RuntimeError, ValueError):
            raise AppError('No se pudo leer el ZIP. Usa un respaldo completo descargado desde esta app.') from None
    if len(raw) > MAX_BACKUP_BYTES or not raw.startswith(b'SQLite format 3\x00'):
        raise AppError('No es un respaldo SQLite del almacén. Un reporte Excel no es un respaldo.')
    return raw


@contextmanager
def validated_backup(blob):
    """Abre una copia de solo lectura y verifica su estructura antes de restaurar."""
    raw = _backup_database_bytes(blob)
    fd, path = tempfile.mkstemp(prefix='marketing_validacion_', suffix='.sqlite3')
    source = None
    try:
        with os.fdopen(fd, 'wb') as f:
            f.write(raw)
        source = sqlite3.connect(Path(path).as_uri()+'?mode=ro&immutable=1', uri=True)
        source.row_factory = sqlite3.Row
        source.execute('PRAGMA trusted_schema=OFF')
        source.execute('PRAGMA query_only=ON')
        deadline = time.monotonic()+20
        source.set_progress_handler(lambda: int(time.monotonic() > deadline), 5000)
        objects = source.execute("SELECT name,type FROM sqlite_master WHERE name NOT LIKE 'sqlite_%'").fetchall()
        if any(x['type'] in ('trigger', 'view') for x in objects):
            raise AppError('El respaldo incluye vistas o disparadores ajenos al formato del almacén.')
        if {x['name'] for x in objects if x['type']=='table'} != set(BACKUP_TABLES):
            raise AppError('La estructura del respaldo no corresponde a esta versión del almacén.')
        for table, columns in REQUIRED_COLUMNS.items():
            actual = tuple(x['name'] for x in source.execute(f'PRAGMA table_info("{table}")'))
            if actual != columns:
                raise AppError(f'El respaldo tiene una estructura incompatible en {table}.')
        if source.execute('PRAGMA quick_check').fetchone()[0] != 'ok':
            raise AppError('El respaldo está dañado. No se modificó el inventario.')
        if source.execute('PRAGMA foreign_key_check').fetchone():
            raise AppError('El respaldo tiene referencias internas incompletas.')
        if meta(source, 'schema') != '1' or meta(source, 'seeded') != '1':
            raise AppError('El respaldo no tiene una apertura válida del almacén.')
        if meta(source, 'source') != SOURCE_NAME:
            raise AppError('El respaldo no corresponde al inventario de apertura de esta aplicación.')
        try:
            date.fromisoformat(meta(source, 'opening_day'))
            int(meta(source, 'revision'))
        except (ValueError, TypeError):
            raise AppError('La fecha o revisión del respaldo es inválida.') from None
        if source.execute("SELECT COUNT(*) FROM documents WHERE kind='APERTURA'").fetchone()[0] != 1:
            raise AppError('El respaldo no tiene exactamente una apertura.')
        if source.execute('SELECT 1 FROM lines GROUP BY product_id HAVING SUM(delta)<0 LIMIT 1').fetchone():
            raise AppError('El respaldo contiene existencias negativas; requiere revisión.')
        for row in source.execute('SELECT day,data_json FROM counts'):
            value = json.loads(row['data_json'])
            if not isinstance(value, list): raise AppError('Un conteo del respaldo no es válido.')
        for row in source.execute('SELECT month,data_json FROM closures'):
            value = json.loads(row['data_json'])
            if not isinstance(value, dict): raise AppError('Un cierre del respaldo no es válido.')
        summary = {
            'references': source.execute('SELECT COUNT(*) FROM products').fetchone()[0],
            'units': source.execute('SELECT COALESCE(SUM(delta),0) FROM lines').fetchone()[0],
            'documents': source.execute('SELECT COUNT(*) FROM documents').fetchone()[0],
            'attachments': source.execute('SELECT COUNT(*) FROM attachments').fetchone()[0],
            'last_movement': source.execute('SELECT MAX(day) FROM documents').fetchone()[0],
            'opening_day': meta(source,'opening_day'),
            'sha256': hashlib.sha256(raw).hexdigest(),
        }
        source.set_progress_handler(None, 0)
        yield source, summary
    except AppError:
        raise
    except (sqlite3.Error, ValueError, TypeError, KeyError):
        raise AppError('El respaldo es incompatible o está dañado. No se modificó el inventario.') from None
    finally:
        if source is not None: source.close()
        Path(path).unlink(missing_ok=True)


def inspect_backup(blob):
    with validated_backup(blob) as (_, summary):
        return summary


def restore_backup(blob, expected_revision, expected_epoch):
    """Sustituye el estado completo, no suma saldos. Requiere confirmación en la UI.

    La copia previa se conserva en el servidor como protección adicional, NO
    como sustituto de descargar una copia fuera de Streamlit.
    """
    with validated_backup(blob) as (source, summary):
        with LOCK:
            current = storage_status()
            if current['revision'] != expected_revision or current['epoch'] != expected_epoch:
                raise AppError('El almacén cambió desde la vista previa. Revísala de nuevo antes de restaurar.')
            previous = backup_zip()
            folder = DATA_DIR/'respaldos'
            folder.mkdir(parents=True, exist_ok=True)
            filename = 'antes_restaurar_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:8]+'.zip'
            (folder/filename).write_bytes(previous)
            dest = connect()
            try:
                source.backup(dest, pages=128)
            finally:
                dest.close()
            with transaction() as c:
                put_meta(c, 'storage_epoch', uuid.uuid4().hex)
                put_meta(c, 'last_restore', now_iso())
                put_meta(c, 'last_restore_hash', summary['sha256'])
                put_meta(c, 'revision', max(current['revision'], int(meta(c, 'revision', '0')))+1)
                put_meta(c, 'app_version', VERSION)
            return {**summary, 'previous_backup': filename}


def backup_zip():
    """Copia coherente del inventario, fichas, firmas, logos y cierres."""
    with LOCK:
        fd, path = tempfile.mkstemp(prefix='marketing_respaldo_', suffix='.sqlite3')
        os.close(fd)
        src=connect(); dest=sqlite3.connect(path)
        try:
            src.backup(dest)
        finally:
            dest.close(); src.close()
        try:
            out = io.BytesIO()
            with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
                z.write(path, 'marketing.sqlite3')
                z.writestr('RESTAURAR.txt',
                    'MARKETING MAKROPETROL - RESPALDO COMPLETO\n'
                    'En la app abre Respaldos y recuperación. Sube este ZIP, revisa la vista previa\n'
                    'y confirma Restaurar. Reemplaza el estado completo; NO suma inventarios.\n'
                    'Antes de restaurar descarga una copia del estado actual.\n'
                    'Incluye movimientos, catálogo, cantidades, firmas adjuntas, logos y cierres.\n'
                    'Un Excel o PDF de reporte NO sustituye este respaldo.\n'
                    'En Streamlit el disco local puede perderse; conserva este ZIP fuera de la app.\n')
                z.writestr('manifest.json', json.dumps({
                    'application':'Makropetrol Marketing', 'version':VERSION,
                    'created_at':now_iso(), 'storage':'sqlite-autonomo',
                    **storage_status()}, ensure_ascii=False, indent=2))
            return out.getvalue()
        finally:
            Path(path).unlink(missing_ok=True)

# MENSAJES DEL COMPONENTE
logger=logging.getLogger('marketing')
POST_ROUTES={
    '/api/movement':submit_movement,'/api/resolve':submit_resolution,
    '/api/cancel':cancel_document,'/api/product':add_product,
    '/api/count':save_count,'/api/close-month':close_month,
    '/api/import':import_rows,'/api/attach':attach_file,'/api/logo':save_logo,
}


def dispatch(request):
    rid=str(request.get('id',''))
    if not re.fullmatch(r'[a-zA-Z0-9-]{12,80}',rid):
        return {'id':rid,'status':400,'body':{'error':'Identificador de solicitud inválido.'}}
    def result(body,status=200,mime='application/json; charset=utf-8',filename=None):
        headers={'Content-Type':mime,'Cache-Control':'no-store'}
        if filename:
            safe=re.sub(r'[^A-Za-z0-9_.-]','_',filename)
            headers['Content-Disposition']=f'attachment; filename="{safe}"'
        if isinstance(body,(bytes,bytearray,memoryview)):
            return {'id':rid,'status':status,'headers':headers,'encoding':'base64','body':base64.b64encode(body).decode()}
        return {'id':rid,'status':status,'headers':headers,'body':body}
    try:
        raw_path=str(request.get('path',''))
        if len(raw_path)>2048: raise AppError('Ruta demasiado larga.')
        url=urlparse(raw_path)
        if url.scheme or url.netloc: raise AppError('Ruta no permitida.')
        path=url.path
        q={k:v[0] for k,v in parse_qs(url.query).items()}
        method=request.get('method','GET')
        if method=='POST':
            data=request.get('data')
            if path not in POST_ROUTES: return result({'error':'Acción no disponible.'},404)
            if not isinstance(data,dict): raise AppError('Solicitud inválida.')
            if len(json.dumps(data))>MAX_UPLOAD*2: raise AppError('El archivo supera el tamaño permitido.')
            with LOCK:
                with reading() as c:
                    if data.get('_storage_epoch') != meta(c,'storage_epoch'):
                        raise AppError('El almacén se restauró o reinició. Pulsa Actualizar y revisa el inventario antes de registrar otra ficha.')
                return result(POST_ROUTES[path](data))
        if method!='GET': raise AppError('Método no permitido.')
        if path=='/api/state': return result(snapshot())
        if path=='/api/document': return result(document_data(integer(q.get('id'))))
        if path=='/api/receipt':
            doc=document_data(integer(q.get('id')))
            return result(receipt_pdf(doc),mime='application/pdf',filename=doc['number']+'.pdf')
        if path in ('/api/report','/api/report-file'):
            with reading() as c:r=filtered_report(report_data(c,q.get('month','')),q.get('brand'))
            if path=='/api/report': return result(r)
            name=re.sub('[^A-Za-z0-9_-]','_',r['scope'])+'_'+r['month']
            if q.get('format')=='pdf': return result(monthly_pdf(r),mime='application/pdf',filename=name+'.pdf')
            return result(xlsx_report(r),mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename=name+'.xlsx')
        if path=='/api/all-reports':
            return result(all_reports_zip(q.get('month','')),mime='application/zip',filename='Reportes_marketing_'+q.get('month','')+'.zip')
        if path=='/api/template':
            return result(excel_template(q.get('mode','movement')),mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename='Plantilla_marketing.xlsx')
        if path=='/api/backup': return result(backup_zip(),mime='application/zip',filename='Respaldo_marketing_'+today().isoformat()+'.zip')
        if path=='/api/attachment':
            with reading() as c:
                r=c.execute('SELECT name,mime,data FROM attachments WHERE id=?',(integer(q.get('id')),)).fetchone()
            if not r: raise AppError('Adjunto no encontrado.')
            return result(r['data'],mime=r['mime'],filename=r['name'])
        return result({'error':'Ruta no encontrada.'},404)
    except AppError as exc:
        return result({'error':str(exc)},400)
    except ImportError:
        return result({'error':'Falta una dependencia. Revisa requirements.txt y los registros de Streamlit.'},500)
    except Exception as exc:
        # No devolver datos de conexión, credenciales, SQL ni fichas al navegador/log.
        logger.error('Acción fallida (%s); solicitud %s',type(exc).__name__,rid)
        return result({'error':'No se pudo completar la operación. Consulta Movimientos antes de repetir una entrega y revisa los registros de Streamlit. No borres datos_marketing.'},503)


FRONTEND_HTML = r"""
<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Marketing · Makropetrol</title>
<style>
:root{--ink:#302340;--muted:#81728f;--purple:#7755b7;--plum:#543479;--lav:#eee6fa;--pink:#f4dfeb;--bg:#f7f5fa;--line:#e9e1ef;--good:#317968;--red:#ae4b69;--shadow:0 12px 38px #5e3e8710;--radius:22px}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Segoe UI,Arial,sans-serif;font-size:14px}button,input,select,textarea{font:inherit}button{cursor:pointer;transition:transform .16s,box-shadow .16s,background .16s}button:hover{transform:translateY(-1px)}button:active{transform:translateY(0)}button:disabled{opacity:.55;cursor:wait;transform:none}button:focus-visible,input:focus,select:focus,textarea:focus{outline:3px solid #d3b7f0;outline-offset:2px}a{color:var(--purple)}
.side{position:fixed;inset:0 auto 0 0;width:232px;background:#fff;border-right:1px solid var(--line);padding:30px 20px;display:flex;flex-direction:column;z-index:10}.brand{display:flex;align-items:center;gap:12px;margin-bottom:44px}.brand-icon{background:linear-gradient(135deg,#9570c9,#d7a5c3);height:44px;width:44px;border-radius:15px;display:grid;place-items:center;color:white;font-weight:800;letter-spacing:-1px}.brand strong{font-size:19px}.eyebrow{font-size:10px;font-weight:700;letter-spacing:1.6px;color:var(--muted);text-transform:uppercase}.navgroup{display:grid;gap:9px}.navbtn{display:flex;align-items:center;gap:12px;text-align:left;border:0;background:transparent;padding:14px 15px;border-radius:14px;color:var(--muted);font-weight:600}.navbtn.active{background:var(--lav);color:var(--plum)}.navbtn:hover{background:#f7f3fc}.sidefoot{margin-top:auto;border:1px solid var(--line);border-radius:18px;padding:16px;background:linear-gradient(135deg,#faf7fd,#fff4f7);line-height:1.7}.small{font-size:12px;color:var(--muted)}.ico{width:20px;height:20px;display:inline-block;vertical-align:middle;flex-shrink:0}
main{margin-left:232px;padding:28px 40px 60px;max-width:1700px}.topbar{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-bottom:28px}.crumb{font-size:12px;color:var(--muted)}.pill{display:inline-flex;align-items:center;gap:7px;background:white;border:1px solid var(--line);border-radius:30px;padding:9px 13px;font-size:12px}.dot{width:7px;height:7px;border-radius:50%;background:#60a48d}.headrow{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:25px}.headrow h1{font-size:29px;letter-spacing:-.8px;margin:0 0 7px}.headrow p{color:var(--muted);margin:0;line-height:1.6}.actions{display:flex;gap:9px;flex-wrap:wrap}.btn{border:1px solid var(--line);border-radius:12px;background:#fff;padding:11px 17px;display:inline-flex;align-items:center;justify-content:center;gap:8px;color:var(--ink);font-weight:600}.btn:hover{box-shadow:0 5px 14px #68469319}.primary{background:var(--purple);color:#fff;border-color:var(--purple)}.soft{background:var(--lav);border-color:var(--lav);color:var(--plum)}.danger{color:var(--red);border-color:#f0d7e1}.btn.sm{padding:7px 10px;font-size:12px;border-radius:9px}.hero{background:linear-gradient(110deg,#e8dbf6 0%,#f1e2f0 58%,#f9e9e8 100%);border-radius:27px;padding:30px 34px;display:flex;align-items:center;justify-content:space-between;overflow:hidden;position:relative;margin-bottom:24px}.hero:after{content:'';width:245px;height:245px;border:35px solid #ffffff50;border-radius:50%;position:absolute;right:-50px;top:-80px;pointer-events:none}.hero h1{font-size:34px;letter-spacing:-1.1px;margin:10px 0 12px;line-height:1.15}.hero p{color:#786183;max-width:500px;line-height:1.6;margin-bottom:22px}.hero-art{display:grid;grid-template-columns:repeat(2,82px);gap:12px;transform:rotate(-7deg);margin-right:45px;z-index:1}.hero-art div{height:77px;display:grid;place-items:center;background:#ffffffb8;border:1px solid #fff;border-radius:23px;box-shadow:0 10px 20px #643f8010;color:#8b6cb5}.hero-art .ico{height:32px;width:32px}.hero-art div:nth-child(2){transform:translateY(15px)}.hero-art div:nth-child(4){transform:translateY(15px);background:#b79acb;color:#fff}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}.stat{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:21px 22px;box-shadow:var(--shadow)}.stat .label{color:var(--muted);font-size:12px;display:flex;align-items:center;justify-content:space-between}.stat strong{font-size:31px;display:block;font-weight:650;letter-spacing:-1px;margin:13px 0 6px}.stat .hint{font-size:11px;color:var(--muted)}.card{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:24px;box-shadow:var(--shadow);margin-bottom:22px}.card h2{font-size:17px;margin:0 0 7px;letter-spacing:-.3px}.twocol{display:grid;grid-template-columns:1.15fr 1fr;gap:22px}.quickgrid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-top:19px}.quick{border:1px solid var(--line);background:#fff;padding:19px;text-align:left;border-radius:17px}.quick>span{display:inline-flex;align-items:center;justify-content:center;background:#f2ebf9;border-radius:12px;width:38px;height:38px;color:var(--purple);margin-bottom:12px}.quick b{display:block;margin-bottom:6px}.quick p{font-size:12px;color:var(--muted);line-height:1.5;margin:0}.quick:hover{border-color:#c6b2df;background:#fdfbff}.bars{display:grid;gap:16px;margin-top:23px}.barlabel{display:flex;justify-content:space-between;gap:12px;font-size:12px;margin-bottom:7px}.bartrack{height:8px;background:#f4effa;border-radius:7px;overflow:hidden}.barfill{height:100%;background:linear-gradient(90deg,#b59cd4,#8060b0);border-radius:7px}.filters{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin:18px 0}.filters input{flex:1;min-width:200px}.field{display:flex;flex-direction:column;gap:7px}.field label{font-size:12px;color:#74617d;font-weight:600}.field small{font-size:11px;color:var(--muted)}input,select,textarea{border:1px solid #ded4e7;padding:11px 12px;border-radius:10px;background:#fff;color:var(--ink);min-width:0}textarea{resize:vertical;min-height:70px}.formgrid{display:grid;grid-template-columns:repeat(2,1fr);gap:17px}.full{grid-column:1/-1}.badge{display:inline-flex;padding:5px 9px;border-radius:30px;background:#f1ebf8;font-size:11px;color:#795b9e;white-space:nowrap}.badge.good{background:#eaf5f0;color:var(--good)}.badge.zero{background:#fbebf1;color:var(--red)}.badge.warn{background:#fff2da;color:#9d752e}.tablewrap{overflow:auto;max-height:590px;border:1px solid var(--line);border-radius:14px}table{border-collapse:collapse;width:100%;font-size:12px}th{font-size:10px;letter-spacing:.5px;text-transform:uppercase;background:#f8f5fb;color:#82728f;text-align:left;padding:13px 12px;position:sticky;top:0;z-index:1;white-space:nowrap}td{border-top:1px solid #f1ebf5;padding:12px;vertical-align:middle}tr:hover td{background:#fdfbff}.num{text-align:right;font-variant-numeric:tabular-nums}.code{font-size:11px;color:#8e7aa4;white-space:nowrap}.cellname{min-width:180px;max-width:390px}.empty{text-align:center;padding:45px;color:var(--muted);line-height:1.8}.notice{padding:13px 16px;background:#f4eefb;color:#79648a;border-radius:13px;line-height:1.6;font-size:12px;margin:15px 0;border:1px solid #ece2f5}.notice.warn{background:#fff8e9;color:#927044;border-color:#f7ecd6}.eventcard{border:1px solid var(--line);border-radius:17px;padding:20px;display:flex;justify-content:space-between;align-items:center;gap:12px;margin-top:14px}.eventcard h3{margin:8px 0}.divider{height:1px;background:var(--line);margin:22px 0}.path{font-family:Consolas,monospace;font-size:12px;word-break:break-all;background:#f8f4fb;padding:15px;border-radius:12px}.checkline{display:flex;gap:9px;align-items:flex-start;font-size:12px;color:#776283;line-height:1.6}.checkline input{margin-top:3px}
.overlay{position:fixed;inset:0;background:#291c3c66;backdrop-filter:blur(7px);z-index:30;display:flex;align-items:center;justify-content:center;padding:25px}.modal{background:#fff;border-radius:25px;width:min(900px,100%);max-height:94vh;display:flex;flex-direction:column;box-shadow:0 24px 90px #24172c45;animation:enter .18s ease-out}.modal.wide{width:min(1180px,100%)}.modalhead{padding:23px 27px 17px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:flex-start;gap:16px}.modalhead h2{font-size:22px;margin:0 0 6px}.modalbody{padding:23px 27px;overflow:auto}.modalfoot{padding:17px 27px;border-top:1px solid var(--line);display:flex;justify-content:flex-end;gap:10px;background:#fdfbff;border-radius:0 0 25px 25px;flex-wrap:wrap}.close{border:0;background:#f4eef9;border-radius:50%;width:31px;height:31px;color:#8e75a7;font-size:21px}.addline{display:grid;grid-template-columns:1fr 120px 100px;gap:10px;align-items:end}.addline input{width:100%}.dropzone{border:1.5px dashed #c8b2df;border-radius:15px;padding:17px;text-align:center;background:#fcfaff;color:#836698;font-size:12px}.dropzone.drag{background:#eee3fa}.dropzone input{padding:0;border:0;font-size:12px;max-width:100%;margin-top:8px}.qtybox{width:90px;text-align:right;padding:7px}.error{background:#fbeaf0;color:#9e3957;border:1px solid #f2cbda;padding:12px;border-radius:12px;font-size:12px;white-space:pre-wrap;margin:14px 0}.success{background:#eaf5ee;color:#2f755a;padding:14px;border-radius:12px;line-height:1.6;margin:14px 0}.toast{position:fixed;right:30px;bottom:25px;max-width:480px;padding:17px 23px;border-radius:16px;background:#39254f;color:#fff;box-shadow:0 12px 40px #0003;z-index:80;font-size:13px;line-height:1.5}.toast.err{background:#a94c67}.loading{display:grid;place-items:center;min-height:70vh;color:var(--muted)}.loader{width:35px;height:35px;border:3px solid #e5d8f4;border-top-color:#9872c7;border-radius:50%;animation:spin .7s linear infinite}.sectionlabel{color:var(--purple);font-size:11px;font-weight:bold;letter-spacing:1px;text-transform:uppercase;margin-bottom:15px}.roundicon{padding:8px;border-radius:11px;background:#f4eef9;color:var(--purple)}.right{text-align:right}.checkfooter{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-top:16px}.countstats{font-size:12px;color:#7b658e}.attachment{display:flex;justify-content:space-between;align-items:center;border:1px solid var(--line);padding:10px;border-radius:10px;margin-top:7px}
@keyframes spin{to{transform:rotate(360deg)}}@keyframes enter{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@media(max-width:1100px){.side{width:195px;padding:25px 14px}main{margin-left:195px;padding:25px}.stats{gap:10px}.stat{padding:17px}.hero-art{margin-right:0}.twocol{grid-template-columns:1fr}.hero h1{font-size:30px}}
@media(max-width:760px){.side{position:static;width:100%;padding:15px}.brand{margin-bottom:15px}.navgroup{display:flex;overflow-x:auto}.navbtn{padding:10px;white-space:nowrap}.navbtn .ico,.sidefoot{display:none}main{margin:0;padding:20px 15px}.topbar{margin-bottom:18px}.stats{grid-template-columns:repeat(2,1fr)}.hero{padding:25px}.hero-art{display:none}.headrow{align-items:flex-start;flex-direction:column}.formgrid{grid-template-columns:1fr}.full{grid-column:auto}.overlay{padding:9px}.modalbody,.modalhead{padding:18px}.addline{grid-template-columns:1fr 80px}.addline .btn{grid-column:1/-1}.hero h1{font-size:29px}.pill.time{display:none}.eventcard{flex-direction:column;align-items:flex-start}.modalfoot{padding:14px}.topbar .btn{padding:8px}}
@media(prefers-reduced-motion:reduce){*,*:before,*:after{transition:none!important;animation:none!important}}
html,body{height:100%;overflow-y:auto} main{min-height:100vh} @media(min-width:761px){.side{height:100vh}} </style></head><body><noscript><p>Activa JavaScript para utilizar el almacén.</p></noscript>
<aside class="side"><div class="brand"><div class="brand-icon">MP</div><div><strong>Marketing</strong><div class="eyebrow" style="margin-top:5px">Makropetrol</div></div></div><div class="navgroup" id="nav"></div><div class="sidefoot"><div class="eyebrow">Un solo almacén</div><b>Matriz Makropetrol</b><div class="small">Datos guardados en el servidor. Descarga respaldo diario y antes de actualizar; Streamlit puede eliminar este disco.</div><button class="btn sm soft" style="margin-top:12px;width:100%" onclick="download('/api/backup')">Descargar respaldo</button></div></aside>
<main><div class="topbar"><div class="crumb">Marketing <span style="padding:0 8px">/</span> <b id="crumb">Inicio</b></div><div class="actions"><span class="pill time" id="datepill"></span><span class="pill"><span class="dot"></span>Modo autónomo</span><button class="btn sm" onclick="refresh()" title="Actualizar datos">↻ Actualizar</button></div></div><div id="content"><div class="loading"><div><div class="loader"></div><p>Abriendo tu almacén…</p></div></div></div></main>
<div id="modalroot"></div><div id="toastroot"></div>
<script src="bridge.js"></script>
<script>
'use strict';
const TOKEN='streamlit-component';let S=null,page='home',cart=[],moveKind='',moveKey='',importHash=null,countDraft={},countRevision=0,countProducts=[],countKey='',report=null;
const $=id=>document.getElementById(id);const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=n=>Number(n||0).toLocaleString('es-VE');const clean=s=>String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();
const key=()=>crypto.randomUUID();const savedOperator=()=>{try{return localStorage.getItem('mp_operator')||''}catch(_){return ''}};const rememberOperator=v=>{try{localStorage.setItem('mp_operator',v)}catch(_){}};const datefmt=s=>s?new Date(s+'T12:00:00').toLocaleDateString('es-VE',{day:'2-digit',month:'short',year:'numeric'}):'';
const iconPaths={home:'<path d="m3 10 9-7 9 7v10H3z"/><path d="M9 20v-8h6v8"/>',box:'<path d="m3 7 9-4 9 4v11l-9 4-9-4zM3 7l9 5 9-5M12 12v10M7 5l10 5"/>',arrow:'<path d="M4 12h16m-6-6 6 6-6 6"/>',in:'<path d="M12 3v12m-5-5 5 5 5-5M4 15v6h16v-6"/>',out:'<path d="M12 16V4m-5 5 5-5 5 5M4 15v6h16v-6"/>',calendar:'<rect x="3" y="5" width="18" height="16" rx="3"/><path d="M7 2v6m10-6v6M3 11h18m-13 5h4"/>',chart:'<path d="M4 3v18h17M8 16v-5m5 5V6m5 10V9"/>',settings:'<circle cx="12" cy="12" r="4"/><path d="M12 2v3m0 14v3M2 12h3m14 0h3M5 5l2 2m10 10 2 2M5 19l2-2M17 7l2-2"/>',check:'<path d="m4 12 5 5L20 6"/>',gift:'<rect x="3" y="9" width="18" height="4" rx="1"/><path d="M5 13v8h14v-8M12 9v12M12 9C3 9 4 0 9 4l3 5zm0 0c9 0 8-9 3-5l-3 5z"/>',file:'<path d="M14 2H5v20h14V7zM14 2v5h5M8 12h8m-8 4h8"/>',plus:'<path d="M12 4v16M4 12h16"/>'};
const ico=n=>`<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${iconPaths[n]||iconPaths.box}</svg>`;
const labels={APERTURA:'Apertura',ENTRADA:'Recepción',ENTREGA:'Entrega a asesor',EVENTO:'Salida para evento',PRESTAMO:'Préstamo',LIQUIDACION:'Regreso / cierre',AJUSTE:'Conteo / ajuste',ANULACION:'Anulación'};
const navItems=[['home','Inicio','home'],['inventory','Inventario','box'],['history','Movimientos','file'],['events','Eventos y préstamos','calendar'],['reports','Cierre y reportes','chart'],['settings','Ayuda y ajustes','settings']];
async function api(url,data){if(data!==undefined)data={...data,_storage_epoch:S?.storage_epoch||''};const res=await window.marketingBridgeFetch(url,{method:data===undefined?'GET':'POST',headers:{'X-App-Token':TOKEN,...(data===undefined?{}:{'Content-Type':'application/json'})},...(data===undefined?{}:{body:JSON.stringify(data)})});let v;try{v=await res.json()}catch{throw Error('La app no respondió. Comprueba tu conexión y recarga la página si es necesario.')}if(!res.ok||v.error)throw Error(v.error||'No se pudo completar la operación.');return v}
function toast(t,err=false){$('toastroot').innerHTML=`<div class="toast ${err?'err':''}" role="status">${esc(t)}</div>`;setTimeout(()=>{$('toastroot').innerHTML=''},6500)}
async function download(url){try{toast('Preparando el archivo…');const res=await window.marketingBridgeFetch(url,{headers:{'X-App-Token':TOKEN}});if(!res.ok){const x=await res.json();throw Error(x.error)}const blob=await res.blob();const a=document.createElement('a');a.href=URL.createObjectURL(blob);const name=(res.headers.get('Content-Disposition')||'').match(/filename="([^"]+)"/);a.download=name?name[1]:'marketing.zip';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),30000);toast('Archivo preparado. Revisa tus descargas.')}catch(e){toast(e.message,true)}}
async function refresh(){try{S=await api('/api/state');$('datepill').textContent=datefmt(S.today);render()}catch(e){$('content').innerHTML=`<div class="error">${esc(e.message)}</div>`}}
function go(p){page=p;render()}
function render(){$('nav').innerHTML=navItems.map(([id,t,i])=>`<button class="navbtn ${id===page?'active':''}" onclick="go('${id}')">${ico(i)}${t}</button>`).join('');$('crumb').textContent=navItems.find(x=>x[0]===page)?.[1]||'';({home:home,inventory:inventory,history:history,events:events,reports:reports,settings:settings}[page]||home)()}
function stat(label,value,hint,ic='box'){return `<div class="stat"><div class="label">${label}<span class="roundicon">${ico(ic)}</span></div><strong>${fmt(value)}</strong><span class="hint">${hint}</span></div>`}
function header(title,subtitle,actions=''){return `<div class="headrow"><div><h1>${title}</h1><p>${subtitle}</p></div><div class="actions">${actions}</div></div>`}
function brandOptions(all=true){return (all?'<option value="">Todas las marcas</option>':'')+S.brands.map(b=>`<option value="${b.id}">${esc(b.name)}</option>`).join('')}
function home(){const total=S.products.reduce((a,p)=>a+p.stock,0),zero=S.products.filter(p=>!p.stock).length,pending=S.events.reduce((a,e)=>a+e.pending,0);const grouped={};for(const p of S.products)grouped[p.brand]=(grouped[p.brand]||0)+p.stock;const top=Object.entries(grouped).sort((a,b)=>b[1]-a[1]).slice(0,6);const max=Math.max(...top.map(x=>x[1]),1);
$('content').innerHTML=`<section class="hero"><div><div class="eyebrow">Tu espacio de trabajo</div><h1>Todo el almacén,<br>en un solo lugar.</h1><p>Recibe, entrega y consulta la mercancía de cada marca.<br>Menos planillas sueltas. Más claridad para tu equipo.</p><div class="actions"><button class="btn primary" onclick="movement('ENTREGA')">${ico('out')} Registrar entrega</button><button class="btn" onclick="movement('ENTRADA')">${ico('in')} Recibir mercancía</button></div></div><div class="hero-art"><div>${ico('gift')}</div><div>${ico('file')}</div><div>${ico('chart')}</div><div>${ico('check')}</div></div></section>
<div class="stats">${stat('Unidades en almacén',total,'Saldo disponible de todas las marcas')}${stat('Referencias en catálogo',S.products.length,`${S.brands.length} marcas / grupos`,'gift')}${stat('Artículos agotados',zero,'Permanecen disponibles para recargar','box')}${stat('Fuera del almacén',pending,'Eventos y préstamos pendientes','calendar')}</div>
<div class="twocol"><div class="card"><h2>¿Qué necesitas hacer?</h2><div class="small">Cada acción actualiza un mismo inventario.</div><div class="quickgrid"><button class="quick" onclick="movement('ENTRADA')"><span>${ico('in')}</span><b>Recibir mercancía</b><p>Agregar lo que entrega el proveedor.</p></button><button class="quick" onclick="movement('ENTREGA')"><span>${ico('out')}</span><b>Entregar a un asesor</b><p>Descontar y generar su ficha para firma.</p></button><button class="quick" onclick="movement('EVENTO')"><span>${ico('calendar')}</span><b>Preparar un evento</b><p>Controlar lo que sale y los sobrantes.</p></button><button class="quick" onclick="go('reports')"><span>${ico('chart')}</span><b>Preparar reportes</b><p>Gerencia o una marca por separado.</p></button></div></div>
<div class="card"><h2>Existencias por marca</h2><div class="small">Las seis marcas con mayor cantidad disponible.</div><div class="bars">${top.map(([b,n])=>`<div><div class="barlabel"><span>${esc(b)}</span><b>${fmt(n)}</b></div><div class="bartrack"><div class="barfill" style="width:${100*n/max}%"></div></div></div>`).join('')}</div></div></div>
<div class="notice">Inventario inicial: <b>${datefmt(S.opening_day)}</b> · 237 referencias incorporadas desde tu conteo. Los artículos agotados no se eliminan; una recepción vuelve a darles existencia.</div>
<div class="card"><div class="headrow"><div><h2>Actividad reciente</h2><div class="small">Las fichas guardadas permanecen en el historial.</div></div><button class="btn sm" onclick="go('history')">Ver movimientos →</button></div>${docsTable(S.documents.filter(d=>d.kind!=='APERTURA').slice(0,5))}</div>`}
function inventory(){$('content').innerHTML=header('Tu inventario','Una referencia por artículo y marca. Las cantidades cambian con los movimientos.',`<button class="btn" onclick="newProduct()">${ico('plus')} Nuevo artículo</button><button class="btn primary" onclick="movement('ENTRADA')">${ico('in')} Recibir mercancía</button>`)+`<div class="card"><div class="filters"><input id="invSearch" placeholder="Buscar gorra, llavero, código…" oninput="inventoryRows()"><select id="invBrand" onchange="inventoryRows()">${brandOptions()}</select><select id="invStatus" onchange="inventoryRows()"><option value="all">Todos los estados</option><option value="stock">Con existencia</option><option value="zero">Agotados</option></select></div><div id="invRows"></div></div>`;inventoryRows()}
function inventoryRows(){const q=clean($('invSearch').value),b=$('invBrand').value,status=$('invStatus').value;const list=S.products.filter(p=>(!b||p.brand_id===+b)&&(!q||clean(p.code+' '+p.name+' '+p.brand).includes(q))&&(status==='all'||(status==='zero'?p.stock===0:p.stock>0)));$('invRows').innerHTML=`<div class="small" style="margin-bottom:12px">${list.length} artículos · ${fmt(list.reduce((a,p)=>a+p.stock,0))} unidades</div><div class="tablewrap"><table><thead><tr><th>Código</th><th>Artículo</th><th>Marca / proveedor</th><th class="num">Existencia</th><th>Estado</th></tr></thead><tbody>${list.map(p=>`<tr><td class="code">${p.code}</td><td class="cellname">${esc(p.name)}</td><td>${esc(p.brand)}</td><td class="num"><b>${fmt(p.stock)}</b></td><td><span class="badge ${p.stock?'good':'zero'}">${p.stock?'Disponible':'Agotado'}</span></td></tr>`).join('')||'<tr><td colspan="5" class="empty">No hay artículos con esos filtros.</td></tr>'}</tbody></table></div>`}
function docsTable(ds){if(!ds.length)return '<div class="empty">Todavía no hay movimientos.<br>Registra una recepción o una entrega para comenzar.</div>';return `<div class="tablewrap"><table><thead><tr><th>Ficha</th><th>Fecha</th><th>Movimiento</th><th>Destino / asesor</th><th>Responsable</th><th>Estado</th><th></th></tr></thead><tbody>${ds.map(d=>`<tr><td class="code">${d.number}</td><td>${d.day}</td><td>${labels[d.kind]}</td><td>${esc(d.recipient||'Almacén')}</td><td>${esc(d.operator)}</td><td><span class="badge ${d.reversed_by?'zero':'good'}">${d.reversed_by?'Anulada':'Guardada'}</span></td><td><button class="btn sm" onclick="showDocument(${d.id})">Ver ficha</button></td></tr>`).join('')}</tbody></table></div>`}
function history(){$('content').innerHTML=header('Movimientos y fichas','Reimprime comprobantes, adjunta firmas y consulta quién registró cada entrega.',`<button class="btn primary" onclick="movement('ENTREGA')">${ico('plus')} Nueva entrega</button>`)+`<div class="card"><div class="filters"><input id="histQ" placeholder="Buscar ficha, asesor o responsable…" oninput="historyRows()"><select id="histKind" onchange="historyRows()"><option value="">Todos los movimientos</option>${Object.entries(labels).map(([k,l])=>`<option value="${k}">${l}</option>`).join('')}</select></div><div id="histRows"></div><p class="small">Últimos 400 documentos. El Excel mensual incluye todos los movimientos de su período.</p></div>`;historyRows()}
function historyRows(){const q=clean($('histQ').value),k=$('histKind').value;$('histRows').innerHTML=docsTable(S.documents.filter(d=>(!k||d.kind===k)&&(!q||clean([d.number,d.recipient,d.operator,d.note].join(' ')).includes(q))))}
function events(){$('content').innerHTML=header('Eventos y préstamos','Controla lo que está fuera y reincorpora únicamente lo que realmente regresa.',`<button class="btn" onclick="movement('PRESTAMO')">Préstamo de exhibidor</button><button class="btn primary" onclick="movement('EVENTO')">${ico('plus')} Salida para evento</button>`)+`<div class="card"><h2>Salidas pendientes de resolver</h2><p class="small">Confirmar una entrega durante el evento no vuelve a descontar inventario. El regreso sí aumenta la existencia.</p>${S.events.map(e=>`<div class="eventcard"><div><span class="badge">${labels[e.kind]} · ${e.number}</span><h3>${esc(e.recipient)}</h3><div class="small">${datefmt(e.day)} · ${esc(e.operator)} · ${e.items.length} artículos</div></div><div class="right"><b>${fmt(e.pending)} pendientes</b><div style="margin-top:12px"><button class="btn soft" onclick="resolveEvent(${e.id})">Registrar regreso / cierre →</button></div></div></div>`).join('')||'<div class="empty">No hay mercancía pendiente fuera del almacén.</div>'}</div><div class="notice">Para stands, ruletas, mesas o exhibidores que deban regresar, elige <b>Préstamo de exhibidor</b>. No se impone esa clasificación a todos los productos.</div>`}
function modal(title,subtitle,body,foot='',wide=false){$('modalroot').innerHTML=`<div class="overlay" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="modal ${wide?'wide':''}"><div class="modalhead"><div><h2>${title}</h2><div class="small">${subtitle}</div></div><button class="close" aria-label="Cerrar" onclick="closeModal()">×</button></div><div class="modalbody">${body}<div id="modalError"></div></div>${foot?`<div class="modalfoot">${foot}</div>`:''}</div></div>`;document.body.style.overflow='hidden'}
function closeModal(){$('modalroot').innerHTML='';document.body.style.overflow=''}
function modalError(e){if($('modalError')){$('modalError').innerHTML=`<div class="error" role="alert">${esc(e.message||e)}</div>`;$('modalError').scrollIntoView({block:'nearest'})}else toast(e.message||e,true)}
function peopleList(){return `<datalist id="operators">${S.operators.map(s=>`<option value="${esc(s)}">`).join('')}</datalist><datalist id="recipients">${S.recipients.map(s=>`<option value="${esc(s)}">`).join('')}</datalist>`}
function movement(kind){moveKind=kind;moveKey=key();cart=[];importHash=null;const isIn=kind==='ENTRADA';const title=isIn?'Recibir mercancía':kind==='ENTREGA'?'Entregar a un asesor':kind==='EVENTO'?'Preparar un evento':'Registrar préstamo';
modal(title,'El inventario se actualiza únicamente al confirmar.',`${peopleList()}<form id="moveForm"><div class="formgrid"><div class="field"><label>Fecha</label><input type="date" id="mDay" value="${S.today}" min="${S.opening_day}" max="${S.today}" required></div><div class="field"><label>Registrado por</label><input id="mOperator" list="operators" autocomplete="off" value="${esc(savedOperator())}" maxlength="200" placeholder="Nombre de la encargada" required></div><div class="field"><label>${isIn?'Referencia de recepción (opcional)':kind==='ENTREGA'?'Nombre del asesor':'Nombre del evento / responsable'}</label><input id="mRecipient" list="recipients" maxlength="200" ${isIn?'':'required'} placeholder="${isIn?'Ejemplo: recepción de material promocional':'Indica quién recibe o el destino'}"></div><div class="field"><label>Documento / referencia (opcional)</label><input id="mReference" maxlength="200" placeholder="Número de guía o referencia interna"></div></div>
<div class="divider"></div><div class="sectionlabel">Artículos del movimiento</div><div class="field" style="margin-bottom:12px"><label>Filtrar por marca</label><select id="mBrand" onchange="productOptions()">${brandOptions()}</select></div><div class="addline"><div class="field"><label>Buscar artículo</label><input id="mProduct" list="productOptions" placeholder="Escribe el nombre o código…" onchange="selectedProduct()" autocomplete="off"><datalist id="productOptions"></datalist></div><div class="field"><label>Cantidad</label><input id="mQty" type="number" min="1" step="1" max="100000000" value="1"></div><button type="button" class="btn soft" onclick="addCart()">+ Agregar</button></div><p id="mAvailable" class="small"></p><div id="cartArea"></div>
<details style="margin-top:16px"><summary style="cursor:pointer;color:var(--purple);font-size:12px">También puedes arrastrar una plantilla Excel / CSV</summary><div class="dropzone" id="moveDrop" style="margin-top:12px">Arrastra aquí el archivo con códigos y cantidades.<br><input type="file" id="moveUpload" accept=".xlsx,.csv" onchange="importMovement(this.files[0])"></div><button type="button" class="btn sm" style="margin-top:9px" onclick="download('/api/template?mode=movement')">Descargar plantilla con catálogo</button><p class="small">La importación reemplaza la lista de esta ficha; no mueve inventario hasta confirmar.</p></details>
<div class="field" style="margin-top:18px"><label>Observaciones (opcional)</label><textarea id="mNote" maxlength="2000" placeholder="Actividad, destino o detalle que deba quedar registrado"></textarea></div><label class="checkline" style="margin-top:15px"><input type="checkbox" id="mConfirm" required><span>Confirmo que ${isIn?'esta mercancía ya fue recibida':'esta mercancía sale físicamente del almacén'}. Guardar la ficha ${isIn?'sumará':'descontará'} las cantidades indicadas.</span></label></form>`, `<button class="btn" onclick="closeModal()">Cancelar</button><button class="btn primary" type="submit" form="moveForm" id="moveSave">${ico('check')} ${isIn?'Confirmar recepción':'Confirmar salida y generar ficha'}</button>`,true);
productOptions();renderCart();$('moveForm').onsubmit=submitMove;setDrop('moveDrop',importMovement)}
function productOptions(){const bid=$('mBrand').value;$('productOptions').innerHTML=S.products.filter(p=>!bid||p.brand_id===+bid).map(p=>`<option value="${esc(p.code+' | '+p.brand+' | '+p.name)}">${p.stock} disponibles</option>`).join('')}
function chosenProduct(){const v=$('mProduct').value.trim(),code=v.split('|')[0].trim().toUpperCase();return S.products.find(p=>p.code===code)||S.products.find(p=>clean(p.name)===clean(v)&&(!$('mBrand').value||p.brand_id===+$('mBrand').value))}
function selectedProduct(){const p=chosenProduct();$('mAvailable').textContent=p?`${p.brand} · ${fmt(p.stock)} unidades disponibles.`:'Selecciona un artículo del catálogo.'}
function addCart(){const p=chosenProduct(),q=Number($('mQty').value);if(!p){modalError('Selecciona un artículo del catálogo.');return}if(!Number.isInteger(q)||q<1){modalError('Escribe una cantidad entera mayor que cero.');return}const old=cart.find(r=>r.product_id===p.id);if(old)old.quantity+=q;else cart.push({product_id:p.id,quantity:q});$('mProduct').value='';$('mQty').value=1;$('mAvailable').textContent='';renderCart()}
function renderCart(){let sum=0;const isIn=moveKind==='ENTRADA';$('cartArea').innerHTML=cart.length?`<div class="tablewrap" style="max-height:300px;margin-top:12px"><table><thead><tr><th>Artículo / marca</th><th class="num">Disponible</th><th class="num">Cantidad</th><th class="num">Quedará</th><th></th></tr></thead><tbody>${cart.map((r,i)=>{const p=S.products.find(p=>p.id===r.product_id);sum+=r.quantity;const after=p.stock+(isIn?r.quantity:-r.quantity);return `<tr><td>${esc(p.name)}<div class="small">${p.code} · ${esc(p.brand)}</div></td><td class="num">${fmt(p.stock)}</td><td class="num"><input class="qtybox" type="number" min="1" max="100000000" value="${r.quantity}" onchange="editCart(${i},this.value)"></td><td class="num" style="${after<0?'color:var(--red)':''}"><b>${fmt(after)}</b></td><td><button class="btn sm danger" type="button" onclick="cart.splice(${i},1);renderCart()" aria-label="Quitar artículo">×</button></td></tr>`}).join('')}</tbody></table></div><p class="small">${cart.length} artículos · <b>${fmt(sum)} unidades</b>${importHash?' · Archivo importado y pendiente de confirmar':''}</p>`:'<div class="notice">Agrega los artículos que entran o salen. Puedes combinar marcas en una misma ficha.</div>'}
function editCart(i,v){const n=Number(v);if(Number.isInteger(n)&&n>0)cart[i].quantity=n;else toast('Usa una cantidad entera positiva.',true);renderCart()}
function readFile(file){return new Promise((resolve,reject)=>{if(!file){reject(Error('Selecciona un archivo.'));return}if(file.size>12*1024*1024){reject(Error('El archivo supera 12 MB.'));return}const r=new FileReader();r.onload=()=>resolve({name:file.name,data:String(r.result).split(',')[1]});r.onerror=()=>reject(Error('No se pudo leer el archivo.'));r.readAsDataURL(file)})}
function setDrop(id,fn){const el=$(id);el.ondragover=e=>{e.preventDefault();el.classList.add('drag')};el.ondragleave=()=>el.classList.remove('drag');el.ondrop=e=>{e.preventDefault();el.classList.remove('drag');fn(e.dataTransfer.files[0])}}
async function importMovement(file){try{const v=await api('/api/import',{...await readFile(file),mode:'movement'});cart=v.items.map(x=>({product_id:x.product_id,quantity:x.quantity}));importHash=v.hash;renderCart();toast(`${v.rows} artículos preparados. Revisa y confirma la ficha.`)}catch(e){modalError(e)}}
async function submitMove(ev){ev.preventDefault();if(!cart.length){modalError('Agrega al menos un artículo.');return}const btn=$('moveSave');btn.disabled=true;try{const operator=$('mOperator').value;const r=await api('/api/movement',{key:moveKey,kind:moveKind,day:$('mDay').value,operator,recipient:$('mRecipient').value,reference:$('mReference').value,note:$('mNote').value,items:cart,import_hash:importHash});rememberOperator(operator);closeModal();await refresh();await showDocument(r.id);toast('Movimiento guardado. El inventario ya está actualizado.')}catch(e){modalError(e)}finally{if($('moveSave'))$('moveSave').disabled=false}}
async function showDocument(id){try{const d=await api('/api/document?id='+id);modal(`Ficha ${d.number}`,`${labels[d.kind]} · ${datefmt(d.day)}`,`<div class="${d.reversed_by?'notice warn':'success'}">${d.reversed_by?'Documento anulado; su contrapartida conserva el historial.':'Movimiento guardado. Descargar, imprimir o adjuntar la firma no cambia de nuevo el inventario.'}</div><div class="formgrid"><div><div class="small">Destino / asesor</div><b>${esc(d.recipient||'Almacén')}</b></div><div><div class="small">Registrado por</div><b>${esc(d.operator)}</b></div></div><div class="divider"></div><div class="tablewrap" style="max-height:310px"><table><thead><tr><th>Artículo</th><th>Marca</th><th class="num">Cantidad</th><th>Disposición</th></tr></thead><tbody>${d.items.map(x=>`<tr><td>${esc(x.name)}<div class="code">${x.code}</div></td><td>${esc(x.brand)}</td><td class="num">${fmt(x.quantity)}</td><td>${esc(x.disposition||(x.delta>0?'Ingreso':'Salida'))}</td></tr>`).join('')||'<tr><td colspan="4">Conteo guardado sin diferencias.</td></tr>'}</tbody></table></div>${d.note?`<p class="small"><b>Observaciones:</b> ${esc(d.note)}</p>`:''}<div class="divider"></div><div class="sectionlabel">Respaldo firmado</div><p class="small">Imprime el PDF, pide la firma y guarda el papel. También puedes adjuntar su foto o PDF aquí.</p>${d.attachments.map(a=>`<div class="attachment"><span>${esc(a.name)}</span><button class="btn sm" onclick="download('/api/attachment?id=${a.id}')">Descargar</button></div>`).join('')}<div class="dropzone" style="margin-top:12px">Adjuntar foto o PDF de esta ficha; no descuenta existencias.<br><input type="file" accept=".pdf,.png,.jpg,.jpeg" onchange="attachDocument(${d.id},this.files[0])"></div>${!d.reversed_by&&!['APERTURA','AJUSTE','ANULACION'].includes(d.kind)?`<details style="margin-top:20px"><summary class="small" style="cursor:pointer">Corregir un movimiento equivocado</summary><p class="small">Anular genera una contrapartida y conserva el original. No borra evidencia.</p><button class="btn sm danger" onclick="cancelDoc(${d.id})">Anular con motivo</button></details>`:''}`,`<button class="btn" onclick="closeModal()">Cerrar</button><button class="btn primary" onclick="download('/api/receipt?id=${d.id}')">${ico('file')} Descargar ficha PDF</button>`,true)}catch(e){toast(e.message,true)}}
async function attachDocument(id,file){try{await api('/api/attach',{...await readFile(file),document_id:id});await showDocument(id);toast('Firma adjuntada. El inventario no se modificó.')}catch(e){modalError(e)}}
async function cancelDoc(id){const note=prompt('Motivo de la anulación (mínimo 8 caracteres):');if(!note)return;const operator=prompt('Nombre de quien realiza la corrección:',savedOperator());if(!operator)return;if(!confirm('Se anulará la ficha mediante una contrapartida, sin borrar su historial. ¿Continuar?'))return;try{const r=await api('/api/cancel',{key:key(),id,note,operator});closeModal();await refresh();await showDocument(r.id)}catch(e){modalError(e)}}
function newProduct(){modal('Crear artículo','Se creará con saldo cero. Después registra su recepción.',`<form id="newForm"><div class="formgrid"><div class="field full"><label>Marca / proveedor</label><input id="newBrand" list="brandNames" required maxlength="200"><datalist id="brandNames">${S.brands.map(b=>`<option value="${esc(b.name)}">`).join('')}</datalist></div><div class="field full"><label>Nombre del artículo</label><input id="newName" maxlength="200" required placeholder="Ejemplo: GORRAS"></div></div><div class="notice">El código se asigna automáticamente y no cambia cuando vuelves a recibir el mismo artículo.</div></form>`,`<button class="btn" onclick="closeModal()">Cancelar</button><button class="btn primary" type="submit" form="newForm">Crear artículo</button>`);$('newForm').onsubmit=async e=>{e.preventDefault();try{const r=await api('/api/product',{brand:$('newBrand').value,name:$('newName').value});closeModal();await refresh();toast(`Artículo ${r.code} creado en cero. Ya puedes registrar su entrada.`)}catch(er){modalError(er)}}}
function resolveEvent(id){const ev=S.events.find(e=>e.id===id);if(!ev)return;const k=key();const pending=ev.items.filter(x=>x.pending>0);modal('Regreso y cierre de evento',`${ev.number} · ${esc(ev.recipient)}`,`${peopleList()}<form id="resolveForm"><div class="formgrid"><div class="field"><label>Fecha</label><input id="rDay" type="date" value="${S.today}" max="${S.today}" min="${ev.day}" required></div><div class="field"><label>Registrado por</label><input id="rOperator" list="operators" value="${esc(savedOperator())}" required></div></div><div class="notice">${ev.kind==='PRESTAMO'?'Registra únicamente material que regresó físicamente. Lo restante seguirá pendiente.':'Lo que regresa vuelve al almacén. Lo entregado durante el evento se confirma, pero no se descuenta otra vez. Puedes dejar cantidades pendientes.'}</div><div class="tablewrap"><table><thead><tr><th>Artículo / marca</th><th class="num">Pendiente</th><th>Regresa</th>${ev.kind==='EVENTO'?'<th>Entregado en evento</th>':''}</tr></thead><tbody>${pending.map(p=>`<tr><td>${esc(p.name)}<div class="small">${p.code} · ${esc(p.brand)}</div></td><td class="num">${fmt(p.pending)}</td><td><input class="qtybox" id="return_${p.product_id}" type="number" min="0" max="${p.pending}" value="0"></td>${ev.kind==='EVENTO'?`<td><input class="qtybox" id="deliver_${p.product_id}" type="number" min="0" max="${p.pending}" value="0"></td>`:''}</tr>`).join('')}</tbody></table></div><div class="field" style="margin-top:17px"><label>Observaciones</label><textarea id="rNote" maxlength="2000"></textarea></div></form>`,`<button class="btn" onclick="closeModal()">Cancelar</button><button class="btn primary" id="resolveSave" form="resolveForm" type="submit">Confirmar cantidades</button>`,true);$('resolveForm').onsubmit=async e=>{e.preventDefault();$('resolveSave').disabled=true;try{const r=await api('/api/resolve',{key:k,event_id:id,day:$('rDay').value,operator:$('rOperator').value,note:$('rNote').value,items:pending.map(p=>({product_id:p.product_id,returned:$('return_'+p.product_id).value,delivered:$('deliver_'+p.product_id)?.value||0}))});closeModal();await refresh();await showDocument(r.id)}catch(er){modalError(er)}finally{if($('resolveSave'))$('resolveSave').disabled=false}}}
function monthOptions(){const list=[];let d=new Date(S.opening_day.slice(0,7)+'-01T12:00:00');const end=S.today.slice(0,7);for(let i=0;i<1200;i++){const m=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');if(m>end)break;list.push(m);d.setMonth(d.getMonth()+1)}return list.reverse().map(m=>`<option value="${m}">${new Date(m+'-01T12:00:00').toLocaleDateString('es-VE',{month:'long',year:'numeric'})}</option>`).join('')}
function reports(){$('content').innerHTML=header('Cierre y reportes','Descarga un resumen para gerencia o un reporte independiente para cada marca.',`<button class="btn" onclick="startCount()">${ico('check')} Hacer conteo físico</button>`)+`<div class="card"><div class="filters"><div class="field"><label>Mes del reporte</label><select id="repMonth" onchange="loadReport()">${monthOptions()}</select></div><div class="field"><label>Destinatario</label><select id="repBrand" onchange="loadReport()"><option value="">Gerencia · todas las marcas</option>${brandOptions(false)}</select></div><div style="flex:1"></div><div class="actions"><button class="btn" onclick="exportReport('xlsx')">Excel detallado</button><button class="btn primary" onclick="exportReport('pdf')">${ico('file')} Reporte PDF</button></div></div><div id="repBody"><div class="empty">Preparando el resumen…</div></div></div><div class="card"><h2>Cerrar el mes</h2><p class="small">El cierre conserva una fotografía del reporte y bloquea movimientos con fecha de ese mes. Solo se permite después del conteo final y cuando llegue el último día del mes.</p><div class="actions" style="margin-top:18px"><button id="closeMonthBtn" class="btn soft" onclick="closeMonthAction()">Cerrar el mes seleccionado</button><button class="btn" onclick="allReports()">Descargar todos los reportes en ZIP</button></div><div class="notice">El ZIP completo es para uso interno. A cada proveedor se le entregan solo sus archivos. Las firmas y planillas no se adjuntan a los reportes de proveedores.</div></div>`;loadReport()}
let reportSeq=0;
async function loadReport(){const seq=++reportSeq;try{const r=await api('/api/report?month='+$('repMonth').value+'&brand='+$('repBrand').value);if(seq!==reportSeq||!$('repBody'))return;report=r;const t=r.totals;$('repBody').innerHTML=`<div style="display:flex;justify-content:space-between;gap:12px;align-items:center;margin:20px 0"><div><b>${esc(r.scope)}</b><div class="small" style="margin-top:5px">Corte: ${datefmt(r.end)} · Apertura: ${datefmt(r.opening_day)}</div></div><span class="badge ${r.status==='CERRADO'?'good':'warn'}">${r.status}</span></div><div class="stats">${stat('Existencia final',t.final,'Disponible al corte')}${stat('Entradas',t.incoming,'Recepciones de proveedores','in')}${stat('Salidas registradas',t.deliveries+t.event_out,'Asesores, eventos y préstamos','out')}${stat('Regresos',t.returns,`${fmt(t.pending)} unidades fuera pendientes`,'calendar')}</div><div class="notice"><b>Cómo cuadra:</b> ${fmt(t.initial)} inicial + ${fmt(t.incoming)} recibido − ${fmt(t.deliveries+t.event_out)} salido + ${fmt(t.returns)} regresado + (${fmt(t.adjustments)}) ajustes = <b>${fmt(t.final)} en almacén.</b><br>Las anulaciones aparecen como ajustes; consulta su detalle en el Excel.</div><div class="tablewrap" style="max-height:430px"><table><thead><tr><th>Marca / artículo</th><th class="num">Inicial</th><th class="num">Entradas</th><th class="num">Salidas</th><th class="num">Regresos</th><th class="num">Ajustes</th><th class="num">Final</th><th class="num">Fuera</th></tr></thead><tbody>${r.products.map(p=>`<tr><td>${esc(p.name)}<div class="small">${p.code} · ${esc(p.brand)}</div></td><td class="num">${fmt(p.initial)}</td><td class="num">${fmt(p.incoming)}</td><td class="num">${fmt(p.deliveries+p.event_out)}</td><td class="num">${fmt(p.returns)}</td><td class="num">${fmt(p.adjustments)}</td><td class="num"><b>${fmt(p.final)}</b></td><td class="num">${fmt(p.pending)}</td></tr>`).join('')}</tbody></table></div><p class="small">Los eventos confirmados como entregados se detallan sin descontar dos veces. El reporte del mes en curso sigue actualizándose.</p>`;$('closeMonthBtn').disabled=r.status==='CERRADO'||r.status==='EN CURSO';$('closeMonthBtn').textContent=r.status==='CERRADO'?'Mes ya cerrado':r.status==='EN CURSO'?'Mes en curso · cierre al finalizar':'Cerrar el mes seleccionado'}catch(e){if($('repBody'))$('repBody').innerHTML=`<div class="error">${esc(e.message)}</div>`}}
function exportReport(format){download('/api/report-file?month='+$('repMonth').value+'&brand='+$('repBrand').value+'&format='+format)}
function allReports(){if(confirm('Este ZIP contiene gerencia y todas las marcas. Es para uso interno, no para enviarlo completo a un proveedor. ¿Prepararlo?'))download('/api/all-reports?month='+$('repMonth').value)}
async function closeMonthAction(){const month=$('repMonth').value;if(!confirm(`Se cerrará ${month} para todas las marcas. No se podrán agregar movimientos con fecha de ese mes. ¿Continuar?`))return;const operator=prompt('Responsable del cierre:',savedOperator());if(!operator)return;try{await api('/api/close-month',{month,operator});toast('Mes cerrado. Se conservó el reporte.');await loadReport()}catch(e){toast(e.message,true)}}
async function startCount(){try{S=await api('/api/state');countProducts=S.products.map(p=>({...p}));countRevision=S.revision;countDraft={};countKey=key();modal('Conteo físico del almacén','Compara lo que hay físicamente con el saldo de la app. No registres movimientos mientras cuentas.',`${peopleList()}<form id="countForm"><div class="formgrid"><div class="field"><label>Fecha del conteo</label><input id="cDay" type="date" value="${S.today}" min="${S.opening_day}" max="${S.today}" required></div><div class="field"><label>Realizado por</label><input id="cOperator" list="operators" value="${esc(savedOperator())}" required></div></div><div class="notice">Una casilla vacía significa pendiente de contar. Escribe <b>0</b> si confirmaste que no hay unidades. El ajuste se aplica únicamente al confirmar el conteo completo.</div><div class="actions"><button type="button" class="btn sm" onclick="download('/api/template?mode=count')">Descargar plantilla de conteo</button><label class="btn sm" style="cursor:pointer">Importar conteo<input type="file" hidden accept=".xlsx,.csv" onchange="importCount(this.files[0])"></label><button type="button" class="btn sm" onclick="countZeros()">Confirmar agotados en cero</button><button type="button" class="btn sm" onclick="saveCountDraft()">Guardar borrador CSV</button></div><div class="filters"><input id="countSearch" placeholder="Buscar artículo…" oninput="countRows()"><select id="countBrand" onchange="countRows()">${brandOptions()}</select></div><div id="countProgress" class="countstats" style="margin:10px 0"></div><div id="countRows"></div><div class="field" style="margin-top:18px"><label>Motivo de las diferencias / observaciones</label><textarea id="cNote" maxlength="2000" placeholder="Explica las diferencias si existen; no se borrará el historial."></textarea></div><label class="checkline" style="margin-top:15px"><input type="checkbox" required><span>Confirmo que las cantidades fueron contadas físicamente. Al guardar, las diferencias quedarán como un ajuste identificado.</span></label></form>`,`<button class="btn" onclick="closeModal()">Cancelar</button><button class="btn primary" type="submit" form="countForm" id="countSave">Confirmar conteo y diferencias</button>`,true);countRows();$('countForm').onsubmit=submitCount}catch(e){toast(e.message,true)}}
function countRows(){const b=$('countBrand').value,q=clean($('countSearch').value);const ps=countProducts.filter(p=>(!b||p.brand_id===+b)&&(!q||clean(p.code+' '+p.name).includes(q)));$('countRows').innerHTML=`<div class="tablewrap" style="max-height:40vh"><table><thead><tr><th>Artículo / marca</th><th class="num">Saldo esperado</th><th class="num">Contado</th><th class="num">Diferencia</th></tr></thead><tbody>${ps.map(p=>`<tr><td>${esc(p.name)}<div class="small">${p.code} · ${esc(p.brand)}</div></td><td class="num">${fmt(p.stock)}</td><td class="num"><input class="qtybox" type="number" min="0" step="1" max="100000000" value="${countDraft[p.id]??''}" placeholder="—" oninput="countChange(${p.id},this.value)"></td><td class="num" id="dif_${p.id}">${countDraft[p.id]===undefined?'—':fmt(Number(countDraft[p.id])-p.stock)}</td></tr>`).join('')}</tbody></table></div>`;countProgress()}
function countChange(id,value){if(value==='')delete countDraft[id];else countDraft[id]=value;const p=countProducts.find(x=>x.id===id);if($('dif_'+id))$('dif_'+id).textContent=value===''?'—':fmt(Number(value)-p.stock);countProgress()}
function countProgress(){const n=Object.keys(countDraft).length,d=countProducts.filter(p=>countDraft[p.id]!==undefined&&Number(countDraft[p.id])!==p.stock).length;$('countProgress').textContent=`${n} de ${countProducts.length} artículos contados · ${d} referencias con diferencia · ${countProducts.length-n} pendientes`}
function countZeros(){if(!confirm('¿Verificaste físicamente que todos los artículos con saldo esperado cero están realmente agotados?'))return;for(const p of countProducts)if(p.stock===0&&countDraft[p.id]===undefined)countDraft[p.id]=0;countRows()}
async function importCount(file){try{const r=await api('/api/import',{...await readFile(file),mode:'count'});for(const x of r.items)countDraft[x.product_id]=x.quantity;countRows();toast('Cantidades cargadas para revisar. Todavía no se aplicaron ajustes.')}catch(e){modalError(e)}}
function saveCountDraft(){const cs=v=>'"'+String(v??'').replace(/"/g,'""')+'"';let content='\ufeffCodigo;Marca;Articulo;Contado\r\n'+countProducts.map(p=>[p.code,p.brand,p.name,countDraft[p.id]??''].map(cs).join(';')).join('\r\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type:'text/csv;charset=utf-8'}));a.download='Borrador_conteo.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
async function submitCount(e){e.preventDefault();$('countSave').disabled=true;try{const r=await api('/api/count',{key:countKey,revision:countRevision,day:$('cDay').value,operator:$('cOperator').value,note:$('cNote').value,items:countProducts.map(p=>({product_id:p.id,counted:countDraft[p.id]??null}))});closeModal();await refresh();await showDocument(r.id);toast('Conteo completo guardado con su historial de diferencias.')}catch(er){modalError(er)}finally{if($('countSave'))$('countSave').disabled=false}}
function settings(){$('content').innerHTML=header('Ayuda y ajustes','Lo esencial para usar la app sin complicaciones.')+`<div class="twocol"><div class="card"><h2>Cuatro pasos para empezar</h2><div class="notice"><b>1.</b> Consulta el inventario: el conteo inicial ya está cargado.<br><b>2.</b> Recibe mercancía para sumar nuevas existencias.<br><b>3.</b> Confirma entregas y descarga las fichas para firmar.<br><b>4.</b> Haz el conteo y descarga los reportes mensuales.</div><h2 style="margin-top:24px">Las planillas firmadas</h2><p class="small" style="line-height:1.8">La ficha incluida es un formato inicial con espacio para las firmas. El formato oficial pendiente se podrá adaptar después. Una foto o PDF firmado se adjunta como respaldo; no se interpreta como una nueva salida.</p><h2 style="margin-top:24px">Importaciones</h2><p class="small" style="line-height:1.8">Usa la plantilla con columnas Código, Marca, Artículo y Cantidad. Para un artículo nuevo, créalo primero. No cargues otra vez el inventario comparativo inicial como si fuera una recepción.</p><button class="btn soft" onclick="download('/api/template?mode=movement')">Descargar plantilla de movimientos</button></div>
<div><div class="card"><h2>Logos en los reportes</h2><p class="small" style="line-height:1.8">Elige la empresa o una marca y carga su logo. Solo se utilizará en sus fichas o reportes correspondientes.</p><div class="field"><label>Destino del logo</label><select id="logoBrand"><option value="0">Makropetrol · fichas y gerencia</option>${brandOptions(false)}</select></div><div class="dropzone" style="margin-top:15px">Logo en PNG o JPG<input type="file" accept=".png,.jpg,.jpeg" onchange="uploadLogo(this.files[0])"></div><div id="logoMessage"></div></div><div class="card"><h2>Datos y respaldo</h2><p class="small">La base, los logos, las firmas y los cierres están guardados aquí:</p><div class="path">${esc(S.data_path)}</div><p class="small" style="line-height:1.8">Los movimientos se guardan en un archivo SQLite del servidor, NO en GitHub. Streamlit puede eliminarlo al reiniciar, actualizar o recrear la app. Descarga un respaldo completo al terminar la jornada y antes de modificar el código. Para recuperarlo, abre Respaldos y recuperación, encima del espacio de trabajo. Las copias automáticas del mismo servidor también pueden perderse.</p><button class="btn primary" onclick="download('/api/backup')">Descargar respaldo completo</button></div></div></div><div class="card"><h2>Información de esta versión</h2><p class="small" style="line-height:1.8">Versión ${S.version} · Almacén único · Solo cantidades de artículos completos · Sin integración con A2 ni envío automático de correos.<br>Esta edición se ejecuta en Streamlit sin claves ni servicios externos. No requiere mantener encendida la PC, pero NO garantiza persistencia entre despliegues. No incluye login: cualquier visitante autorizado por el alojamiento puede ver y modificar el inventario.<br>Las dos referencias OILSTONE de 18 herramientas se conservan separadas. ULTRALUB permanece como referencia histórica en cero, sin copiar su antiguo saldo.</p></div>`}
async function uploadLogo(file){try{await api('/api/logo',{...await readFile(file),brand_id:$('logoBrand').value});$('logoMessage').innerHTML='<div class="success">Logo guardado para ese destinatario.</div>'}catch(e){$('logoMessage').innerHTML=`<div class="error">${esc(e.message)}</div>`}}
refresh();
</script></body></html>

"""

BRIDGE_JS = r"""/* Streamlit v1 component protocol. No database credentials or inventory here. */
(()=>{
  'use strict';
  const pending=new Map();
  const queue=[];
  let active=null, ready=false;
  let parentOrigin='*';
  try {if(document.referrer) parentOrigin=new URL(document.referrer).origin;} catch(_){}
  const send=(type,extra={})=>window.parent.postMessage({isStreamlitMessage:true,type,...extra},parentOrigin);
  const uid=()=>crypto.randomUUID();
  function next(){
    if(!ready || active || !queue.length)return;
    active=queue.shift();
    send('streamlit:setComponentValue',{value:active.request,dataType:'json'});
  }
  function receive(event){
    if(event.source!==window.parent || !event.data || event.data.type!=='streamlit:render')return;
    if(parentOrigin!=='*' && event.origin!==parentOrigin)return;
    ready=true;
    const args=event.data.args||{};
    send('streamlit:setFrameHeight',{height:Math.max(650,Math.min(1400,args.frame_height||900))});
    const result=args.response;
    if(result && pending.has(result.id)){
      const entry=pending.get(result.id);
      pending.delete(result.id);
      clearTimeout(entry.timer);
      if(active && active.request.id===result.id)active=null;
      const headers=new Headers(result.headers||{});
      let body;
      if(result.encoding==='base64'){
        const raw=atob(result.body||'');
        const bytes=new Uint8Array(raw.length);
        for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);
        body=bytes;
      }else{
        body=JSON.stringify(result.body??{});
      }
      entry.resolve(new Response(body,{status:result.status||200,headers}));
    }
    next();
  }
  window.marketingBridgeFetch=(url,options={})=>new Promise((resolve,reject)=>{
    let parsed;
    try{parsed=new URL(url,'https://marketing.invalid');}catch(e){reject(e);return;}
    if(parsed.origin!=='https://marketing.invalid' || !parsed.pathname.startsWith('/api/')){
      reject(new Error('Ruta no permitida.'));return;
    }
    let data=null;
    try{if(options.body)data=JSON.parse(options.body);}catch(e){reject(new Error('Solicitud inválida.'));return;}
    const request={id:uid(),path:parsed.pathname+parsed.search,method:options.method||'GET',data};
    const entry={request,resolve,reject};
    entry.timer=setTimeout(()=>{
      pending.delete(request.id);
      const idx=queue.indexOf(entry);if(idx>=0)queue.splice(idx,1);
      if(active===entry)active=null;
      reject(new Error('La respuesta está tardando o se perdió la conexión. Consulta Movimientos antes de crear otra ficha: el guardado puede haberse completado.'));
      next();
    },180000);
    pending.set(request.id,entry);queue.push(entry);next();
  });
  window.addEventListener('message',receive);
  send('streamlit:componentReady',{apiVersion:1});
})();
"""

# ---------------------------------------------------------------------------
# Entrada Streamlit. Todo el código y catálogo están en este mismo archivo.
# No pide contraseña; no lee Secrets; no crea servicios externos.
# ---------------------------------------------------------------------------
def run_streamlit():
    import streamlit as st
    import streamlit.components.v1 as components

    st.set_page_config(page_title='Marketing · Makropetrol', page_icon='📦',
                       layout='wide', initial_sidebar_state='collapsed')
    st.markdown('''<style>
        .stApp{background:#f7f5fa;color:#302340;}
        .block-container{padding:1rem .8rem .5rem;max-width:none;}
        [data-testid="stExpander"]{background:white;border-radius:16px;}
    </style>''', unsafe_allow_html=True)
    try:
        init_db()
        state = storage_status()
    except Exception as exc:
        logging.getLogger('marketing').error('No se pudo abrir el almacén: %s', type(exc).__name__)
        st.error('No se pudo abrir la base de esta app. No borres datos_marketing.')
        st.info('Revisa los registros de Streamlit y conserva cualquier respaldo descargado. '
                'No se ha sustituido una base dañada por una nueva.')
        st.stop()

    st.warning('Modo autónomo, sin claves. Los movimientos se guardan en el disco del servidor, '
               'NO en GitHub. Streamlit puede eliminar ese disco. Descarga un respaldo al terminar '
               'la jornada y ANTES de actualizar o reiniciar. Sin respaldo podrías perder los movimientos.')
    st.caption('Acceso directo: cualquier visitante con acceso a esta publicación puede ver y modificar '
               'el almacén. Versión '+VERSION)
    if state['documents'] == 1:
        st.info('Esta base contiene solo el inventario de apertura. Si ya habías registrado movimientos '
                'en otra ejecución o en tu PC, restaura tu respaldo antes de seguir. No se recuperan automáticamente.')
    if st.session_state.pop('marketing_restore_notice', False):
        st.success('Respaldo restaurado. Se recuperó su estado completo; no se sumó la apertura otra vez.')

    with st.expander('Respaldos y recuperación', expanded=False):
        st.write('Un respaldo ZIP contiene el catálogo, los movimientos, las firmas, los logos y los cierres. '
                 'Guárdalo en tu PC. Un reporte Excel o PDF no permite restaurar todo el almacén.')
        if st.button('Preparar respaldo actual', key='prepare_full_backup'):
            st.session_state['prepared_backup'] = backup_zip()
            st.session_state['prepared_backup_revision'] = state['revision']
            st.session_state['prepared_backup_time'] = now_iso()
        if st.session_state.get('prepared_backup'):
            if st.session_state.get('prepared_backup_revision') != state['revision']:
                st.warning('Hubo cambios desde que preparaste esta copia. Prepara otro respaldo para incluirlos.')
            st.download_button('Descargar respaldo completo', data=st.session_state['prepared_backup'],
                file_name='Respaldo_Marketing_'+today().isoformat()+'.zip', mime='application/zip',
                key='native_backup_download', on_click='ignore')
            st.caption('Copia preparada: '+st.session_state['prepared_backup_time'])
        uploaded=st.file_uploader('Recuperar un respaldo ZIP o SQLite de Marketing', type=['zip','sqlite3'],
                                 key='restore_file')
        if uploaded is not None:
            try:
                if uploaded.size > MAX_BACKUP_BYTES: raise AppError('El respaldo supera 150 MB.')
                content=uploaded.getvalue()
                details=inspect_backup(content)
                signature=details['sha256']
                st.write(f"El respaldo contiene {details['references']} artículos, {details['units']:,} unidades, "
                         f"{details['documents']} documentos y {details['attachments']} adjuntos. "
                         f"Último movimiento: {details['last_movement']}.")
                # La referencia se conserva durante la confirmación, no se actualiza
                # silenciosamente si otra usuaria registra una entrega entre ambas acciones.
                if st.session_state.get('restore_candidate') != signature:
                    st.session_state['restore_candidate'] = signature
                    st.session_state['restore_expected'] = (state['revision'], state['epoch'])
                confirmed=st.checkbox('Descargué mi copia actual y confirmo REEMPLAZAR todo el estado con este respaldo.',
                                      key='confirm_restore_'+signature[:16])
                if st.button('Restaurar respaldo', disabled=not confirmed, key='perform_restore'):
                    rev, epoch=st.session_state['restore_expected']
                    restore_backup(content, rev, epoch)
                    for key_ in ('cloud_seen_request','cloud_response','prepared_backup','restore_candidate','restore_expected'):
                        st.session_state.pop(key_, None)
                    st.session_state['marketing_restore_notice']=True
                    st.rerun()
            except AppError as exc:
                st.error(str(exc))
                if st.button('Actualizar vista previa', key='refresh_restore_preview'):
                    st.session_state.pop('restore_candidate',None)
                    st.rerun()

    # Los únicos archivos temporales del componente son HTML y JavaScript.
    # La base tiene su ruta propia; no se codifica en el iframe ni en session_state.
    @st.cache_resource(show_spinner=False)
    def component_files(asset_signature):
        root=Path(tempfile.gettempdir())/('marketing_autonomo_ui_'+asset_signature)
        root.mkdir(parents=True, exist_ok=True)
        for name, text_ in [('index.html', FRONTEND_HTML), ('bridge.js', BRIDGE_JS)]:
            temp=root/(name+'.'+uuid.uuid4().hex+'.tmp')
            temp.write_text(text_, encoding='utf-8')
            os.replace(temp, root/name)
        return str(root)

    signature=hashlib.sha256((FRONTEND_HTML+BRIDGE_JS).encode()).hexdigest()[:20]
    component=components.declare_component('marketing_autonomo',path=component_files(signature))
    request=component(response=st.session_state.get('cloud_response'), frame_height=950,
                      key='marketing_workspace_'+state['epoch'], default=None)
    if isinstance(request,dict) and request.get('id') != st.session_state.get('cloud_seen_request'):
        response=dispatch(request)
        st.session_state['cloud_seen_request']=request.get('id')
        st.session_state['cloud_response']=response
        st.rerun()


if __name__ == '__main__':
    run_streamlit()
