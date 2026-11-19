import os
import re
import json
import shutil
import asyncio
import pandas as pd
from datetime import datetime
from urllib.parse import quote
from playwright.async_api import async_playwright

ARCHIVO_HISTORICO = "tarifa_pistacierto_anterior.csv"
ARCHIVO_NUEVO = "tarifa_pistacierto_nueva.csv"
ARCHIVO_CAMBIOS = "cambios_pistacierto.xlsx"
ARCHIVO_ACUMULADO = "historico_maestro_pistacierto.csv"
CARPETA_HISTORIAL = "historico"

TERMINOS_BUSQUEDA = [
    "EXPERTBOOK",
    "EXPERTCENTER",
    "ASUSPRO",
    "90NX",
    "90PF",
    "90PT",
    "90MS"
]

INCLUIR_SIN_STOCK = False
CONCURRENCIA_FICHAS = 6

def log(msg):
    print(msg, flush=True)

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

def limpiar_nombre_mayorista(imgs_info, texto_fila=""):
    up = texto_fila.upper()
    if "PISTACERO" in up or "COMPRA AGRUPADA" in up:
        return "PISTACERO (Compra Agrupada)"

    for alt, title, src in imgs_info:
        for candidato in [alt, title]:
            c = (candidato or "").strip()
            if c and len(c) > 1 and not any(x in c.lower() for x in ["logo", "producto", "asus", "sin imagen", "icon"]):
                return c
        if src:
            archivo = src.split("/")[-1].split("?")[0]
            nombre = re.sub(r"\.(png|jpg|jpeg|svg|webp)$", "", archivo, flags=re.I)
            nombre = re.sub(r"^(logo[-_]?|mayorista[-_]?|prov[-_]?)", "", nombre, flags=re.I)
            nombre = re.sub(r"[-_]", " ", nombre).strip()
            if nombre and len(nombre) > 1 and not nombre.isdigit():
                return nombre.upper() if len(nombre) <= 4 else nombre.title()

    return "Mayorista Desconocido"

def es_producto_asus(texto):
    up = texto.upper()
    claves = ["ASUS", "EXPERTBOOK", "EXPERTCENTER", "ASUSPRO", "ZENBOOK", "VIVOBOOK", "PROART", "CHROMEBOOK", "90NX", "90PF", "90PT", "90MS"]
    return any(c in up for c in claves)

async def cerrar_popups(page):
    for _ in range(2):
        try:
            modal = page.locator(".swal2-container.swal2-backdrop-show")
            if await modal.count() > 0:
                btn_ok = page.locator(".swal2-confirm, .swal2-close, .swal2-cancel").first
                if await btn_ok.count() > 0 and await btn_ok.is_visible():
                    await btn_ok.click(force=True)
                else:
                    await page.keyboard.press("Escape")
                await page.wait_for_timeout(400)
            else:
                break
        except Exception:
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(200)

async def extraer_ficha_producto(context, sem, idx, url_prod, titulo_previo, fecha_ejecucion, progreso):
    async with sem:
        page = await context.new_page()
        filas_resultado = []
        try:
            await page.goto(url_prod, wait_until="domcontentloaded", timeout=35000)
            # Esperar a que carguen por AJAX los precios/mayoristas de la ficha
            try:
                await page.wait_for_selector("text=/€|SKU/i", timeout=8000)
            except Exception:
                pass
            await page.wait_for_timeout(1500)
            await cerrar_popups(page)

            if INCLUIR_SIN_STOCK:
                try:
                    btn_sin_stock = page.locator("text=/Sin stock/i").first
                    if await btn_sin_stock.count() > 0 and await btn_sin_stock.is_visible():
                        await btn_sin_stock.click(force=True)
                        await page.wait_for_timeout(800)
                except Exception:
                    pass

            datos_ficha = await page.evaluate("""() => {
                const bodyTxt = document.body.innerText || "";
                const h1 = document.querySelector("h1, h2, .product-title, .titulo-producto");
                const titulo = h1 ? (h1.innerText || "").trim() : "";

                let pn = "";
                const mPn = bodyTxt.match(/P\\/N:\\s*([A-Z0-9\\-\\.]+)/i);
                if (mPn) {
                    pn = mPn[1].trim();
                }

                // Buscar filas de ofertas por mayorista: contienen precio (ej: 246,37€) y SKU o unidades (47 u.) o canon
                const filas = [];
                const candidatos = Array.from(document.querySelectorAll("div, tr, li")).filter(el => {
                    const t = (el.innerText || "").trim();
                    const tienePrecio = /\\d[\\d\\.]*,\\d{2}\\s*€?/.test(t);
                    const esOferta = /(SKU|\\d+\\s*u\\.|canon|PISTACERO)/i.test(t);
                    return tienePrecio && esOferta && t.length < 450;
                });

                const vistas = new Set();
                for (const el of candidatos) {
                    const t = (el.innerText || "").trim();
                    if (vistas.has(t)) continue;

                    // Quedarnos con la fila individual más interna que tenga precio + indicador de oferta
                    const hijos = Array.from(el.querySelectorAll("div, tr, li")).filter(h => {
                        const ht = (h.innerText || "").trim();
                        return /\\d[\\d\\.]*,\\d{2}\\s*€?/.test(ht) && /(SKU|\\d+\\s*u\\.|canon|PISTACERO)/i.test(ht) && ht.length < 450;
                    });
                    if (hijos.length > 0) continue;

                    vistas.add(t);
                    const imgs = Array.from(el.querySelectorAll("img")).map(img => [
                        img.alt || "",
                        img.title || "",
                        img.src || ""
                    ]);
                    filas.push({ texto: t, imgs: imgs });
                }

                return {
                    titulo: titulo,
                    pn: pn,
                    bodyTxt: bodyTxt.slice(0, 1800),
                    filas: filas
                };
            }""")

            if idx == 0:
                log(f"   [Diagnóstico Ficha 1] URL: {url_prod}")
                log(f"   [Diagnóstico Ficha 1] P/N: '{datos_ficha['pn']}' | Filas detectadas: {len(datos_ficha['filas'])}")
                if datos_ficha["filas"]:
                    log(f"   [Diagnóstico Ficha 1] Ejemplo fila 1: {datos_ficha['filas'][0]}")
                else:
                    muestra_lineas = [l.strip() for l in datos_ficha["bodyTxt"].split("\n") if l.strip()][:25]
                    log(f"   [Diagnóstico Ficha 1] Textos visibles: {muestra_lineas}")

            titulo_final = datos_ficha["titulo"] or titulo_previo
            if es_producto_asus(titulo_final + " " + datos_ficha["bodyTxt"]):
                pn_final = datos_ficha["pn"]
                if not pn_final:
                    m_pn_tit = re.search(r"\b(90[A-Z0-9]{2}[A-Z0-9\-]{4,14}|[A-Z0-9]{5,8}-[A-Z0-9]{4,8})\b", titulo_final)
                    pn_final = m_pn_tit.group(1) if m_pn_tit else titulo_final[:35]

                for fila in datos_ficha["filas"]:
                    txt = fila["texto"]
                    mayorista = limpiar_nombre_mayorista(fila["imgs"], txt)

                    precios = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})", txt)
                    m_canon = re.search(r"\+(\d[\d\.]*[,\.]\d{2})\s*€?\s*canon", txt, re.I)
                    m_stock = re.search(r"(\d+)\s*u\.", txt, re.I)

                    precio = limpiar_numero(precios[0]) if precios else None
                    canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                    stock = int(m_stock.group(1)) if m_stock else 0
                    sin_stock = "sin stock" in txt.lower() or (stock == 0 and "u." in txt.lower())

                    if sin_stock and not INCLUIR_SIN_STOCK:
                        continue

                    if precio is not None:
                        filas_resultado.append({
                            "Fecha": fecha_ejecucion,
                            "P/N": pn_final,
                            "Modelo": titulo_final,
                            "Mayorista": mayorista,
                            "Precio (€)": precio,
                            "Canon (€)": canon,
                            "Stock": stock,
                            "Estado Stock": "Sin stock" if sin_stock else ("Con stock" if stock > 0 else "Consultar"),
                            "URL": url_prod
                        })
        except Exception as e:
            if idx == 0:
                log(f"   [Diagnóstico Ficha 1] Error: {e}")
        finally:
            await page.close()
            progreso["hechas"] += 1
            if progreso["hechas"] % 25 == 0 or progreso["hechas"] == progreso["total"]:
                log(f"   Progreso fichas: {progreso['hechas']}/{progreso['total']} analizadas...")

        return filas_resultado

async def ejecutar_scraping():
    sesion_raw = os.environ.get("PC_SESSION", "").strip()
    if not sesion_raw:
        log("Error: Falta el secreto PC_SESSION en GitHub Actions.")
        return False

    try:
        datos_sesion = json.loads(sesion_raw)
    except Exception as e:
        log(f"Error al leer PC_SESSION (formato JSON inválido): {e}")
        return False

    fecha_ejecucion = datetime.now().strftime("%Y-%m-%d %H:%M")
    productos_encontrados = {}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={"width": 1600, "height": 1000},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        )
        await context.route("**/*.{woff,woff2,ttf,eot,mp4,webm,png,jpg,jpeg,webp,gif,svg}", lambda route: route.abort())

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

        local_data = datos_sesion.get("local", {})
        session_data = datos_sesion.get("session", {})
        await context.add_init_script(f"""
            const lData = {json.dumps(local_data)};
            const sData = {json.dumps(session_data)};
            for (const [k, v] of Object.entries(lData)) {{ window.localStorage.setItem(k, v); }}
            for (const [k, v] of Object.entries(sData)) {{ window.sessionStorage.setItem(k, v); }}
        """)

        page = await context.new_page()

        for termino in TERMINOS_BUSQUEDA:
            pagina_actual = 1
            while True:
                url_pagina = f"https://pistacierto.com/buscador?search={quote(termino)}&page={pagina_actual}&size=100&order=1"
                cargada = False
                for _ in range(3):
                    try:
                        await page.goto(url_pagina, wait_until="domcontentloaded", timeout=30000)
                        await page.wait_for_timeout(2000)
                        cargada = True
                        break
                    except Exception:
                        await page.wait_for_timeout(1000)

                if not cargada:
                    break

                if "/login" in page.url:
                    log("❌ La sesión guardada en PC_SESSION ha caducado.")
                    await browser.close()
                    return False

                await cerrar_popups(page)

                enlaces = await page.evaluate("""() => {
                    return Array.from(document.querySelectorAll("a[href*='/productos/detalle/']"))
                        .map(a => ({ url: a.href.split('?')[0], titulo: (a.innerText || '').trim() }))
                        .filter(x => x.titulo.length > 4);
                }""")

                if not enlaces:
                    break

                nuevos = 0
                for item in enlaces:
                    if item["url"] not in productos_encontrados:
                        productos_encontrados[item["url"]] = item["titulo"]
                        nuevos += 1

                log(f"Buscador '{termino}' pág. {pagina_actual}: +{nuevos} fichas nuevas (Total únicas: {len(productos_encontrados)})")

                if len(enlaces) < 100:
                    break
                pagina_actual += 1

        await page.close()

        if not productos_encontrados:
            log("No se encontraron fichas de productos.")
            await browser.close()
            return False

        log(f"\n2. Extrayendo mayoristas y precios reales de {len(productos_encontrados)} fichas...")
        sem = asyncio.Semaphore(CONCURRENCIA_FICHAS)
        progreso = {"hechas": 0, "total": len(productos_encontrados)}

        tareas = [
            extraer_ficha_producto(context, sem, i, url, tit, fecha_ejecucion, progreso)
            for i, (url, tit) in enumerate(productos_encontrados.items())
        ]

        resultados_por_ficha = await asyncio.gather(*tareas)
        registros = [fila for sublista in resultados_por_ficha for fila in sublista]

        await browser.close()

    if not registros:
        log("No se extrajeron filas de mayoristas dentro de las fichas.")
        return False

    df = pd.DataFrame(registros).drop_duplicates(subset=["P/N", "Mayorista"])
    df.to_csv(ARCHIVO_NUEVO, index=False, sep=";", encoding="utf-8-sig")

    os.makedirs(CARPETA_HISTORIAL, exist_ok=True)
    marca = datetime.now().strftime("%Y-%m-%d_%H-%M")
    ruta_hist = os.path.join(CARPETA_HISTORIAL, f"pistacierto_{marca}.csv")
    shutil.copy2(ARCHIVO_NUEVO, ruta_hist)

    if os.path.exists(ARCHIVO_ACUMULADO):
        df_acum = pd.read_csv(ARCHIVO_ACUMULADO, sep=";", encoding="utf-8-sig")
        df_acum = pd.concat([df_acum, df], ignore_index=True)
    else:
        df_acum = df.copy()
    df_acum.to_csv(ARCHIVO_ACUMULADO, index=False, sep=";", encoding="utf-8-sig")

    log(f"\n3. ¡Completado! Guardados {len(df)} registros reales (Producto x Mayorista) en {ruta_hist}")
    return True

def comparar_tarifas():
    if os.path.exists(ARCHIVO_CAMBIOS):
        os.remove(ARCHIVO_CAMBIOS)

    df_nuevo = pd.read_csv(ARCHIVO_NUEVO, sep=";", encoding="utf-8-sig")

    if not os.path.exists(ARCHIVO_HISTORICO):
        os.replace(ARCHIVO_NUEVO, ARCHIVO_HISTORICO)
        log("Primera ejecución: tarifa base multimayorista guardada correctamente.")
        return

    df_ant = pd.read_csv(ARCHIVO_HISTORICO, sep=";", encoding="utf-8-sig")

    cruce = pd.merge(
        df_nuevo,
        df_ant[["P/N", "Mayorista", "Precio (€)", "Stock"]],
        on=["P/N", "Mayorista"],
        how="inner",
        suffixes=(" Nuevo", " Anterior")
    )

    cambios = cruce[abs(cruce["Precio (€) Nuevo"] - cruce["Precio (€) Anterior"]) >= 0.01].copy()

    if not cambios.empty:
        cambios["Diferencia (€)"] = (cambios["Precio (€) Nuevo"] - cambios["Precio (€) Anterior"]).round(2)
        cambios["Variación (%)"] = ((cambios["Diferencia (€)"] / cambios["Precio (€) Anterior"]) * 100).round(2)

        columnas_excel = [
            "P/N", "Modelo", "Mayorista",
            "Precio (€) Anterior", "Precio (€) Nuevo", "Diferencia (€)", "Variación (%)",
            "Canon (€)", "Stock Anterior", "Stock Nuevo", "Estado Stock", "Fecha", "URL"
        ]
        cambios[columnas_excel].to_excel(ARCHIVO_CAMBIOS, index=False)
        log(f"¡Detectados {len(cambios)} cambios de precio entre mayoristas!")
    else:
        log("Sin variaciones de precio en ningún mayorista.")

    os.replace(ARCHIVO_NUEVO, ARCHIVO_HISTORICO)

if __name__ == "__main__":
    if asyncio.run(ejecutar_scraping()):
        comparar_tarifas()
