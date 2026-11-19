import os
import re
import json
import shutil
import asyncio
import pandas as pd
from datetime import datetime
from playwright.async_api import async_playwright

ARCHIVO_HISTORICO = "tarifa_pistacierto_anterior.csv"
ARCHIVO_NUEVO = "tarifa_pistacierto_nueva.csv"
ARCHIVO_CAMBIOS = "cambios_pistacierto.xlsx"
ARCHIVO_ACUMULADO = "historico_maestro_pistacierto.csv"
CARPETA_HISTORIAL = "historico"

URL_BASE = "https://pistacierto.com"
URL_CATALOGO = "https://pistacierto.com/productos"

def limpiar_numero(texto):
    if not texto:
        return None
    txt = str(texto).replace("€", "").strip()
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    m = re.search(r"\d+(?:\.\d+)?", txt)
    return float(m.group(0)) if m else None

def extraer_nombre_mayorista(imgs_alt_src):
    for alt, src in imgs_alt_src:
        alt_limpio = (alt or "").strip()
        if alt_limpio and len(alt_limpio) > 1 and "logo" not in alt_limpio.lower():
            return alt_limpio
        if src:
            archivo = src.split("/")[-1].split("?")[0]
            nombre = re.sub(r"\.(png|jpg|jpeg|svg|webp)$", "", archivo, flags=re.I)
            nombre = re.sub(r"[-_]", " ", nombre).strip()
            if nombre and len(nombre) > 1:
                return nombre.upper() if len(nombre) <= 4 else nombre.title()
    return "Mayorista Desconocido"

async def ejecutar_scraping():
    sesion_raw = os.environ.get("PC_SESSION", "").strip()
    if not sesion_raw:
        print("Error: Falta el secreto PC_SESSION en GitHub Actions.")
        return False

    try:
        datos_sesion = json.loads(sesion_raw)
    except Exception as e:
        print(f"Error al leer PC_SESSION (formato JSON inválido): {e}")
        return False

    registros = []
    fecha_ejecucion = datetime.now().strftime("%Y-%m-%d %H:%M")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={"width": 1600, "height": 1000},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        )
        await context.route("**/*.{woff,woff2,ttf,eot,mp4,webm}", lambda route: route.abort())

        # Inyectar cookies si existen
        cookies_str = datos_sesion.get("cookies", "")
        if cookies_str:
            lista_cookies = []
            for par in cookies_str.split(";"):
                if "=" in par:
                    k, v = par.strip().split("=", 1)
                    lista_cookies.append({
                        "name": k,
                        "value": v,
                        "domain": "pistacierto.com",
                        "path": "/"
                    })
            if lista_cookies:
                await context.add_cookies(lista_cookies)

        # Inyectar localStorage y sessionStorage antes de que cargue la web
        local_data = datos_sesion.get("local", {})
        session_data = datos_sesion.get("session", {})
        await context.add_init_script(f"""
            const lData = {json.dumps(local_data)};
            const sData = {json.dumps(session_data)};
            for (const [k, v] of Object.entries(lData)) {{ window.localStorage.setItem(k, v); }}
            for (const [k, v] of Object.entries(sData)) {{ window.sessionStorage.setItem(k, v); }}
        """)

        page = await context.new_page()

        print(f"1. Entrando directamente al catálogo con sesión inyectada ({URL_CATALOGO})...")
        await page.goto(URL_CATALOGO, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(3000)

        if "/login" in page.url:
            print("❌ La sesión guardada en PC_SESSION ha caducado o fue rechazada por cambio de IP.")
            await browser.close()
            return False

        print(f"2. ¡Acceso correcto sin bloqueo de IP! URL: {page.url}")

        # Filtrar por ASUS
        buscador = page.locator("input[placeholder*='Buscar']:visible, input[type='search']:visible, input[name*='search']:visible").first
        if await buscador.count() > 0:
            await buscador.fill("ASUS")
            await buscador.press("Enter")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2500)

        filtro_asus = page.locator("label:has-text('ASUS'), span:has-text('ASUS'), a:has-text('ASUS')").first
        try:
            if await filtro_asus.is_visible():
                await filtro_asus.click()
                await page
