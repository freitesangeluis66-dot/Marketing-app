#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Entrada Streamlit sin contraseña de aplicación, solicitada por Angeluis.

Conserva Marketing.py como archivo principal. El motor y el catálogo permanecen
sin cambios en marketing_motor.py. No lee [app].password. Requiere únicamente
[postgres] en Secrets para conservar los movimientos en una base externa.
Si la publicación es pública, cualquier visitante puede usar y modificar el
inventario. La restricción de acceso, si se desea, corresponde al alojamiento.
No introduce SQLite temporal, no reinicia existencias ni modifica la apertura.
"""
from __future__ import annotations

import hashlib
import logging
import os
import tempfile
import uuid
from pathlib import Path
from typing import Mapping, Any

import marketing_motor as motor

VERSION = '1.2.1-sin-login'
motor.VERSION = VERSION


def database_settings(source: Mapping[str, Any]) -> dict[str, Any]:
    """Valida Secrets sin leer contraseñas de entrada ni mostrarlas en pantalla."""
    try:
        values = dict(source['postgres'])
    except Exception:
        raise ValueError('Falta conectar la base de datos del almacén.') from None
    if values.get('driver') not in (None, 'postgres', 'postgresql'):
        raise ValueError('Esta publicación necesita PostgreSQL; no utiliza almacenamiento temporal.')
    keys = ('host', 'dbname', 'user', 'password')
    if any(not isinstance(values.get(k), str) or not values[k].strip() for k in keys):
        raise ValueError('La conexión de la base todavía está incompleta.')
    example_markers = ('HOST_DEL', 'HOST_REAL', 'REFERENCIA_DEL', 'REFERENCIA_REAL',
                       'CONTRASENA_DE_LA_BASE', 'CONTRASENA_REAL', 'REEMPLAZA_', 'TU_CONTRASENA')
    if any(marker in values[k].upper() for k in keys for marker in example_markers):
        raise ValueError('La conexión todavía tiene textos de ejemplo, no los datos de una base real.')
    if '://' in values['host'] or '/' in values['host']:
        raise ValueError('El servidor debe ser un host, no una dirección web completa.')
    try:
        port = int(values.get('port', 5432))
    except (TypeError, ValueError):
        raise ValueError('El puerto de la base debe ser un número.') from None
    if not 1 <= port <= 65535:
        raise ValueError('El puerto de la base no es válido.')
    sslmode = str(values.get('sslmode', 'require'))
    if sslmode not in ('require', 'verify-ca', 'verify-full'):
        raise ValueError('La conexión requiere TLS: sslmode=require o verify-full.')
    return {**{key: values[key] for key in keys}, 'port': port, 'sslmode': sslmode}


def run_streamlit() -> None:
    import streamlit as st
    import streamlit.components.v1 as components

    st.set_page_config(page_title='Marketing · Makropetrol', page_icon='📦',
                       layout='wide', initial_sidebar_state='collapsed')
    st.markdown('''<style>
    .stApp {background:#f7f5fa;color:#302340;}
    .block-container {padding:1.1rem .8rem .5rem;max-width:none;}
    </style>''', unsafe_allow_html=True)

    try:
        settings = database_settings(st.secrets)
        motor.configure(settings)
    except (ValueError, TypeError, KeyError) as exc:
        st.title('Marketing · Makropetrol')
        st.success('Acceso directo: esta versión no pide contraseña para entrar.')
        st.warning(str(exc))
        st.subheader('Solo falta conectar dónde se guardará el inventario')
        st.write('Esta configuración se hace una sola vez en Settings → Secrets. '
                 'Las compañeras no tienen que escribir claves para trabajar.')
        st.info('La conexión debe pertenecer a un proyecto PostgreSQL real. '
                'HOST_REAL, REFERENCIA_REAL y CONTRASENA_REAL eran indicaciones '
                'de ejemplo, no claves que se puedan utilizar.')
        st.caption('La sección [app] ya no se utiliza; puede permanecer o eliminarse. '
                   'No compartas contraseñas en capturas ni en GitHub.')
        st.caption('Por seguridad de los datos, no se crean existencias temporales '
                   'ni se sustituye una base desconectada. Versión '+VERSION)
        st.stop()

    @st.cache_resource(show_spinner=False)
    def initialize_database(settings_signature: str, app_version: str) -> bool:
        motor.init_db()
        return True

    signature = hashlib.sha256(repr(sorted(settings.items())).encode()).hexdigest()
    try:
        with st.spinner('Abriendo el almacén…'):
            initialize_database(signature, VERSION)
    except Exception as exc:
        logging.getLogger('marketing').error(
            'Inicialización fallida: %s; SQLSTATE: %s',
            type(exc).__name__, getattr(exc, 'sqlstate', None))
        st.error('No se pudo conectar al almacén. No se registraron movimientos.')
        st.info(motor.connection_hint(exc))
        st.caption('No se borró inventario ni se creó una base temporal alternativa.')
        if st.button('Volver a comprobar la conexión'):
            st.rerun()
        st.stop()

    @st.cache_resource(show_spinner=False)
    def component_files(asset_signature: str) -> str:
        # Solo archivos estáticos: jamás base de datos o claves en /tmp.
        root = Path(tempfile.gettempdir()) / ('makropetrol_marketing_ui_'+asset_signature)
        root.mkdir(parents=True, exist_ok=True)
        for name, content in [('index.html', motor.FRONTEND_HTML), ('bridge.js', motor.BRIDGE_JS)]:
            temporary = root / (name+'.'+uuid.uuid4().hex+'.tmp')
            temporary.write_text(content, encoding='utf-8')
            os.replace(temporary, root/name)
        return str(root)

    st.caption('MAKROPETROL · MARKETING · Acceso directo · '+VERSION)
    assets = hashlib.sha256((motor.FRONTEND_HTML+motor.BRIDGE_JS).encode()).hexdigest()[:20]
    component = components.declare_component('marketing_workspace', path=component_files(assets))
    request = component(response=st.session_state.get('cloud_response'), frame_height=950,
                        key='marketing_workspace', default=None)
    if isinstance(request, dict) and request.get('id') != st.session_state.get('cloud_seen_request'):
        response = motor.dispatch(request)
        st.session_state['cloud_seen_request'] = request.get('id')
        st.session_state['cloud_response'] = response
        st.rerun()
    with st.expander('Ayuda y acceso'):
        st.caption('Esta versión no incluye inicio de sesión. Si la app es pública, '
                   'cualquier visitante puede consultar y modificar el inventario. '
                   'Puedes restringir las personas autorizadas desde el alojamiento.')
        st.caption('La base se conecta de forma automática mediante Secrets. '
                   'Las credenciales no se envían al espacio de trabajo. Versión '+VERSION)


if __name__ == '__main__':
    run_streamlit()
