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

# Búsquedas de gamas de equipos ASUS (recorrerá TODAS las páginas de cada término)
TERMINOS_BUSQUEDA = [
    "EXPERTBOOK",
    "EXPERTCENTER",
    "ASUSPRO",
    "90NX",      # Portátiles comerciales (ExpertBook, Chromebook)
    "90PF",      # Sobremesas (ExpertCenter D/P/M, ProArt)
    "90PT",      # All-in-One (ExpertCenter E1/E3/E5)
    "90MS",      # Mini PC (ExpertCenter PN/PB)
    "90NB",      # Portátiles VivoBook / ExpertBook
    "90NR"       # Portátiles alto rendimiento
]

# Palabras clave para descartar accesorios, monitores, periféricos y garantías
EXCLUIR_ACCESORIOS = [
    "MONITOR", "ZENSCREEN", "PROART PA", "RATON", "RATÓN", "MOUSE",
    "MOCHILA", "FUNDA", "CARRY BAG", "BACKPACK", "ALFOMBRILLA", "SCABBARD",
    "GARANTIA", "GARANTÍA", "WARRANTY", "PREMIUM CARE", "EXTENSION", "EXTENSIÓN",
    "ADAPTADOR", "TECLADO", "KEYBOARD", "AURICULAR", "HEADSET", "WEBCAM",
    "PROYECTOR", "ROUTER", "EXPERTWIFI", "CONMUTADOR", "SWITCH", "VIVOWATCH",
    "ZENDRIVE", "GRABADORA", "PLACA BASE", "TARJETA", "CAJA ", "TORRE PEQUEÑA",
    "FUENTE DE ALIMENTACION", "CARGADOR", "DOCK", "MALETIN", "MALETÍN"
]

# Cambia a True si quieres incluir también mayoristas sin stock ("Sin stock")
INCLUIR_SIN_STOCK = False
CONCURRENCIA_FICHAS = 6

def log(msg):
    print(msg, flush=True)

def limpiar_numero(texto):
    if texto is None:
        return None
    if isinstance(texto, (int, float)):
        return float(texto)
    txt = str(texto).replace("€", "").strip()
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    m = re.search(r"\d+(?:\.\d+)?", txt)
    return float(m.group(0)) if m else None

def es_equipo_valido(titulo):
    up = titulo.upper()
    if any(exc in up for exc in EXCLUIR_ACCESORIOS):
        return False
    claves_ok = [
        "ASUS", "EXPERTBOOK", "EXPERTCENTER", "ASUSPRO", "ZENBOOK",
        "VIVOBOOK", "CHROMEBOOK", "PORTATIL", "PORTÁTIL", "ORDENADOR",
        "NOTEBOOK", "AIO", "ALL IN ONE", "MINI PC", "MINIPC",
        "90NX", "90PF", "90PT", "90MS", "90NB", "90NR"
    ]
    return any(c in up for c in claves_ok)

def normalizar_mayorista(raw_name, imgs_info=None, texto_fila=""):
    combinado = f"{raw_name or ''} {texto_fila or ''}".upper()
    if imgs_info:
        for alt, title, src in imgs_info:
            combinado += f" {alt or ''} {title or ''} {src or ''}".upper()

    mapa = [
        ("PISTACERO", "PISTACERO (Compra Agrupada)"),
        ("COMPRA AGRUPADA", "PISTACERO (Compra Agrupada)"),
        ("IDIOMUN", "Idiomund"),
        ("GRUPO SII", "Idiomund"),
        ("TD SYNNEX", "TD SYNNEX"),
        ("TDSYNNEX", "TD SYNNEX"),
        ("TECH DATA", "TD SYNNEX"),
        ("DEPAU", "Depau"),
        ("INFORPOR", "Inforpor"),
        ("ESPRINET", "Esprinet"),
        ("VALORISTA", "Valorista"),
        ("MEGASUR", "Megasur"),
        ("V-VALLEY", "V-Valley"),
        ("VVALLEY", "V-Valley"),
        ("INGRAM", "Ingram Micro"),
        ("MCR", "MCR"),
        ("HISPAMICRO", "Hispamicro"),
        ("GLOBOMATIK", "Globomatik"),
        ("BINARY", "Binary"),
        ("GTI", "GTI"),
        ("ARROW", "Arrow"),
        ("DMI", "DMI Computer"),
        ("UNYKA", "Unyka"),
        ("PCBOX", "PcBox"),
        ("TICNOVA", "Ticnova"),
    ]
    for clave, nombre_limpio in mapa:
        if clave in combinado:
            return nombre_limpio

    if raw_name and str(raw_name).strip() and not str(raw_name).strip().isdigit():
        return str(raw_name).strip()

    if imgs_info:
        for alt, title, src in imgs_info:
            for cand in [alt, title]:
                c = (cand or "").strip()
                if c and len(c) > 1 and not any(x in c.lower() for x in ["logo", "producto", "asus", "sin imagen", "icon"]):
                    return c
            if src:
                archivo = src.split("/")[-1].split("?")[0]
                nombre = re.sub(r"\.(png|jpg|jpeg|svg|webp)$", "", archivo, flags=re.I)
                nombre = re.sub(r"^(logo[-_]?|mayorista[-_]?|prov[-_]?|dist[-_]?)", "", nombre, flags=re.I)
                nombre = re.sub(r"[-_]", " ", nombre).strip()
                if nombre and len(nombre) > 1 and not nombre.isdigit():
                    return nombre

    return "Mayorista Desconocido"

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

def extraer_ofertas_de_json_api(data, url_prod, titulo_previo, fecha_ejecucion):
    """Busca recursivamente listas de ofertas de mayoristas en las respuestas JSON de la API de PistaCierto."""
    ofertas = []
    pn_detectado = ""

    def buscar_pn(obj):
        nonlocal pn_detectado
        if pn_detectado or not isinstance(obj, dict):
            return
        for k in ["partNumber", "part_number", "pn", "mpn", "referenciaFabricante", "codigoFabricante"]:
            val = obj.get(k)
            if isinstance(val, str) and len(val.strip()) >= 5:
                pn_detectado = val.strip()
                return
        for v in obj.values():
            if isinstance(v, dict):
                buscar_pn(v)

    buscar_pn(data)

    def recorrer(obj):
        if isinstance(obj, list) and len(obj) > 0 and all(isinstance(x, dict) for x in obj):
            # Comprobar si esta lista contiene ofertas de mayoristas (tienen precio y mayorista/proveedor/stock)
            muestra = obj[0]
            claves = {k.lower() for k in muestra.keys()}
            tiene_precio = any(c in claves for c in ["precio", "price", "pvd", "importe", "coste"])
            tiene_prov = any(c in claves for c in ["mayorista", "proveedor", "wholesaler", "distribuidor", "supplier", "nombre_mayorista", "mayorista_nombre"])
            if tiene_precio and tiene_prov:
                for item in obj:
                    prov_raw = ""
                    for k in ["mayorista", "proveedor", "wholesaler", "distribuidor", "supplier", "nombre_mayorista", "mayorista_nombre"]:
                        v = item.get(k)
                        if isinstance(v, dict):
                            prov_raw = v.get("nombre") or v.get("name") or v.get("alias") or ""
                        elif isinstance(v, str) and v.strip():
                            prov_raw = v.strip()
                        if prov_raw:
                            break

                    precio_val = None
                    for k in ["precio", "price", "pvd", "importe", "precio_final"]:
                        if k in item and item[k] is not None:
                            precio_val = limpiar_numero(item[k])
                            if precio_val is not None:
                                break

                    canon_val = 0.0
                    for k in ["canon", "lpi", "importe_canon"]:
                        if k in item and item[k] is not None:
                            c_num = limpiar_numero(item[k])
                            if c_num is not None:
                                canon_val = c_num
                                break

                    stock_val = 0
                    for k in ["stock", "unidades", "cantidad", "disponible", "qty"]:
                        if k in item and item[k] is not None:
                            try:
                                stock_val = int(float(item[k]))
                                break
                            except Exception:
                                pass

                    if precio_val is not None and precio_val > 0:
                        if stock_val <= 0 and not INCLUIR_SIN_STOCK:
                            continue
                        ofertas.append({
                            "Fecha": fecha_ejecucion,
                            "P/N": pn_detectado,
                            "Modelo": titulo_previo,
                            "Mayorista": normalizar_mayorista(prov_raw),
                            "Precio (€)": precio_val,
                            "Canon (€)": canon_val,
                            "Stock": max(0, stock_val),
                            "Estado Stock": "Con stock" if stock_val > 0 else "Sin stock",
                            "URL": url_prod
                        })
        elif isinstance(obj, dict):
            for v in obj.values():
                recorrer(v)
        elif isinstance(obj, list):
            for v in obj:
                recorrer(v)

    recorrer(data)
    return ofertas, pn_detectado

async def extraer_ficha_producto(context, sem, idx, url_prod, titulo_previo, fecha_ejecucion, progreso):
    async with sem:
        page = await context.new_page()
        filas_resultado = []
        respuestas_json = []

        async def on_response(response):
            try:
                if "pistacero" in response.url or "pistacierto" in response.url:
                    ct = response.headers.get("content-type", "")
                    if "application/json" in ct and response.status == 200:
                        data = await response.json()
                        respuestas_json.append((response.url, data))
            except Exception:
                pass

        page.on("response", on_response)

        try:
            await page.goto(url_prod, wait_until="domcontentloaded", timeout=35000)
            try:
                await page.wait_for_selector("text=/€|SKU|canon/i", timeout=7000)
            except Exception:
                pass
            await page.wait_for_timeout(1500)
            await cerrar_popups(page)

            if INCLUIR_SIN_STOCK:
                try:
                    btn_sin = page.locator("text=/Sin stock/i").first
                    if await btn_sin.count() > 0 and await btn_sin.is_visible():
                        await btn_sin.click(force=True)
                        await page.wait_for_timeout(700)
                except Exception:
                    pass

            # Extraer P/N, título y estructura DOM de las filas de mayoristas
            datos_dom = await page.evaluate("""() => {
                const bodyTxt = document.body.innerText || "";
                const h1 = document.querySelector("h1, h2, .product-title, .titulo-producto");
                const titulo = h1 ? (h1.innerText || "").trim() : "";

                let pn = "";
                const mPn = bodyTxt.match(/P\\/N:\\s*([A-Z0-9\\-\\.]+)/i);
                if (mPn) pn = mPn[1].trim();

                // Buscar filas de mayoristas dentro de la ficha
                const filas = [];
                const candidatos = Array.from(document.querySelectorAll("div, tr, li")).filter(el => {
                    const t = (el.innerText || "").trim();
                    const tienePrecio = /\\d[\\d\\.]*,\\d{2}\\s*€/.test(t);
                    const esFilaOferta = /(u\\.|canon|Sin stock|SKU|PISTACERO)/i.test(t);
                    return tienePrecio && esFilaOferta && t.length < 380;
                });

                const vistas = new Set();
                for (const el of candidatos) {
                    const t = (el.innerText || "").trim();
                    if (vistas.has(t)) continue;
                    const hijos = Array.from(el.querySelectorAll("div, tr, li")).filter(h => {
                        const ht = (h.innerText || "").trim();
                        return /\\d[\\d\\.]*,\\d{2}\\s*€/.test(ht) && /(u\\.|canon|Sin stock|SKU|PISTACERO)/i.test(ht) && ht.length < 380;
                    });
                    if (hijos.length > 0) continue;

                    vistas.add(t);
                    const imgs = Array.from(el.querySelectorAll("img")).map(img => [
                        img.alt || "",
                        img.title || "",
                        img.src || ""
                    ]);
                    filas.push({ texto: t, imgs: imgs, html: el.outerHTML.slice(0, 450) });
                }

                return { titulo: titulo, pn: pn, filas: filas };
            }""")

            titulo_final = datos_dom["titulo"] or titulo_previo
            pn_final = datos_dom["pn"]
            if not pn_final:
                m_pn_tit = re.search(r"\b(90[A-Z0-9]{2}[A-Z0-9\-]{4,14}|[A-Z0-9]{4,10}-[A-Z0-9]{4,10})\b", titulo_final)
                pn_final = m_pn_tit.group(1) if m_pn_tit else titulo_final[:35]

            # 1. Intentar extraer directamente del JSON interceptado de la API
            for _, jdata in respuestas_json:
                ofertas_api, pn_api = extraer_ofertas_de_json_api(jdata, url_prod, titulo_final, fecha_ejecucion)
                if ofertas_api:
                    for of in ofertas_api:
                        if not of["P/N"]:
                            of["P/N"] = pn_final
                    filas_resultado.extend(ofertas_api)

            # 2. Si la API no devolvió la lista en formato estándar, extraer del DOM de la ficha
            if not filas_resultado:
                for fila in datos_dom["filas"]:
                    txt = fila["texto"]
                    mayorista = normalizar_mayorista("", fila["imgs"], txt)

                    precios = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})\s*€", txt)
                    m_canon = re.search(r"\+(\d[\d\.]*[,\.]\d{2})\s*€?\s*canon", txt, re.I)
                    m_stock = re.search(r"(\d+)\s*u\.", txt, re.I)

                    precio = limpiar_numero(precios[0]) if precios else None
                    canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                    stock = int(m_stock.group(1)) if m_stock else 0
                    sin_stock = "sin stock" in txt.lower() or precio == 0.0 or (stock == 0 and "u." in txt.lower())

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
                            "Stock": 0 if sin_stock else stock,
                            "Estado Stock": "Sin stock" if sin_stock else ("Con stock" if stock > 0 else "Consultar"),
                            "URL": url_prod
                        })

            if idx == 0:
                log(f"   [Diagnóstico Ficha 1] URL: {url_prod} | P/N: {pn_final}")
                log(f"   [Diagnóstico Ficha 1] JSONs API capturados: {len(respuestas_json)} | Ofertas extraídas: {len(filas_resultado)}")
                if filas_resultado:
                    log(f"   [Diagnóstico Ficha 1] Muestra fila: {filas_resultado[0]}")
                elif datos_dom["filas"]:
                    log(f"   [Diagnóstico Ficha 1] HTML fila DOM: {datos_dom['filas'][0]['html']}")

        except Exception as e:
            if idx == 0:
                log(f"   [Diagnóstico Ficha 1] Error: {e}")
        finally:
            await page.close()
            progreso["hechas"] += 1
            if progreso["hechas"] % 30 == 0 or progreso["hechas"] == progreso["total"]:
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

        # FASE 1: Recorrer TODAS las páginas de cada término sin detenerse en la página 1
        for termino in TERMINOS_BUSQUEDA:
            pagina_actual = 1
            while pagina_actual <= 30:
                url_pagina = f"https://pistacierto.com/buscador?search={quote(termino)}&page={pagina_actual}&size=100&order=1"
                cargada = False
                for _ in range(3):
                    try:
                        await page.goto(url_pagina, wait_until="domcontentloaded", timeout=35000)
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

                # Solo detenerse cuando la página ya no devuelve ningún producto
                if not enlaces:
                    break

                nuevos_en_esta_pag = 0
                descartados_accesorios = 0
                for item in enlaces:
                    if not es_equipo_valido(item["titulo"]):
                        descartados_accesorios += 1
                        continue
                    if item["url"] not in productos_encontrados:
                        productos_encontrados[item["url"]] = item["titulo"]
                        nuevos_en_esta_pag += 1

                log(f"Buscador '{termino}' pág. {pagina_actual}: {len(enlaces)} items | +{nuevos_en_esta_pag} equipos nuevos (Total únicos: {len(productos_encontrados)})")

                # Si toda la página estaba repetida (el servidor repitió la última página), pasar al siguiente término
                if nuevos_en_esta_pag == 0 and descartados_accesorios == 0:
                    break

                pagina_actual += 1

        await page.close()

        if not productos_encontrados:
            log("No se encontraron equipos ASUS.")
            await browser.close()
            return False

        # FASE 2: Extraer mayoristas, precios y stock reales de cada ficha
        log(f"\n2. Extrayendo mayoristas y precios reales de {len(productos_encontrados)} equipos...")
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
        log("No se extrajeron ofertas de mayoristas.")
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

    log(f"\n✅ Completado: {len(df)} filas reales (Equipo x Mayorista) guardadas en {ruta_hist}")
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
