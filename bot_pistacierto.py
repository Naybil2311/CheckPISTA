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
    "90MS",
    "P1403", "P1503", "P3406", "P3606", "P5405",
    "B1402", "B1502", "B3405", "B5404", "B5604", "B9403",
    "D500", "D700", "D900", "P500", "E1600", "E3402", "E5402"
]

INCLUIR_SIN_STOCK = False

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

def extraer_nombre_mayorista(imgs_info, texto_fila=""):
    up = texto_fila.upper()
    if "PISTACERO" in up or "COMPRA AGRUPADA" in up:
        return "PISTACERO (Compra Agrupada)"
    if "TD SYNNEX" in up or "TECH DATA" in up:
        return "TD SYNNEX"
    if "IDIOMUND" in up:
        return "Idiomund"
    if "DEPAU" in up:
        return "Depau"
    if "INFORPOR" in up:
        return "Inforpor"
    if "ESPRINET" in up:
        return "Esprinet"
    if "VALORISTA" in up:
        return "Valorista"
    if "MEGASUR" in up:
        return "Megasur"
    if "V-VALLEY" in up or "VVALLEY" in up:
        return "V-Valley"
    if "INGRAM" in up:
        return "Ingram Micro"

    for alt, title, src in imgs_info:
        for candidato in [alt, title]:
            c = (candidato or "").strip()
            if c and len(c) > 1 and not any(x in c.lower() for x in ["logo", "producto", "asus", "sin imagen", "icon"]):
                return c
        if src:
            archivo = src.split("/")[-1].split("?")[0]
            nombre = re.sub(r"\.(png|jpg|jpeg|svg|webp)$", "", archivo, flags=re.I)
            nombre = re.sub(r"^(logo[-_]?|mayorista[-_]?|prov[-_]?|dist[-_]?)", "", nombre, flags=re.I)
            nombre = re.sub(r"[-_]", " ", nombre).strip()
            if nombre and len(nombre) > 1 and not nombre.isdigit():
                return nombre.upper() if len(nombre) <= 4 else nombre.title()

    return "Mayorista Desconocido"

def es_producto_asus(texto):
    up = texto.upper()
    claves = [
        "ASUS", "EXPERTBOOK", "EXPERTCENTER", "ASUSPRO", "ZENBOOK",
        "VIVOBOOK", "PROART", "CHROMEBOOK", "90NX", "90PF", "90PT", "90MS",
        "P1403", "P1503", "P3406", "P3606", "B1402", "B1502", "B3405", "B5404", "B9403"
    ]
    return any(c in up for c in claves)

async def cerrar_popups(page):
    for _ in range(3):
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

    registros = []
    urls_procesadas = set()
    diagnostico_hecho = False
    fecha_ejecucion = datetime.now().strftime("%Y-%m-%d %H:%M")

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
                        await page.goto(url_pagina, wait_until="domcontentloaded", timeout=35000)
                        await page.wait_for_timeout(2500)
                        cargada = True
                        break
                    except Exception:
                        await page.wait_for_timeout(1200)

                if not cargada:
                    break

                if "/login" in page.url:
                    log("❌ La sesión guardada en PC_SESSION ha caducado.")
                    await browser.close()
                    return False

                await cerrar_popups(page)

                # Hacer clic real con Playwright en todos los botones "Ver mayoristas" visibles
                botones_ver = page.locator("text=/Ver mayoristas/i")
                total_botones = await botones_ver.count()
                for i in range(total_botones):
                    try:
                        btn = botones_ver.nth(i)
                        if await btn.is_visible():
                            await btn.click(force=True)
                            await page.wait_for_timeout(80)
                    except Exception:
                        continue

                if total_botones > 0:
                    await page.wait_for_timeout(2000)

                tarjetas_pagina = await page.evaluate("""() => {
                    const resultados = [];
                    const enlaces = Array.from(document.querySelectorAll("a[href*='/productos/detalle/']"));
                    const vistas = new Set();

                    for (const enlace of enlaces) {
                        const titulo = (enlace.innerText || "").trim();
                        if (!titulo || titulo.length < 5) continue;
                        const urlProd = enlace.href.split("?")[0];

                        let cont = enlace.parentElement;
                        let mejorCont = null;
                        for (let nivel = 0; nivel < 10 && cont; nivel++) {
                            const enlacesEnCont = cont.querySelectorAll("a[href*='/productos/detalle/']").length;
                            if (enlacesEnCont > 2) break;
                            const txt = cont.innerText || "";
                            if (txt.includes("Ocultar mayoristas") || txt.includes("Ver mayoristas") || txt.includes("STOCK:")) {
                                mejorCont = cont;
                            }
                            cont = cont.parentElement;
                        }

                        if (!mejorCont || vistas.has(mejorCont)) continue;
                        vistas.add(mejorCont);

                        const textoTarjeta = mejorCont.innerText || "";

                        // Buscar filas desplegadas de mayoristas: cada fila tiene precio en € (ej: 608,42€) y unidades o canon
                        const filasMayoristas = [];
                        const candidatos = Array.from(mejorCont.querySelectorAll("div, tr, li")).filter(el => {
                            const t = (el.innerText || "").trim();
                            const tienePrecio = /\\d[\\d\\.]*,\\d{2}\\s*€/.test(t);
                            const esFilaDesplegada = /(\\([0-9]+\\s*u\\.\\)|canon|Sin stock|u\\.\\))/i.test(t);
                            const esCuadroSuperior = /STOCK:\\s*\\d+/i.test(t) && !/canon/i.test(t);
                            return tienePrecio && esFilaDesplegada && !esCuadroSuperior && t.length < 320;
                        });

                        const filasVistas = new Set();
                        for (const fila of candidatos) {
                            const t = (fila.innerText || "").trim();
                            if (filasVistas.has(t)) continue;

                            const sub = Array.from(fila.querySelectorAll("div, tr, li")).filter(s => {
                                const st = (s.innerText || "").trim();
                                return /\\d[\\d\\.]*,\\d{2}\\s*€/.test(st) && /(\\([0-9]+\\s*u\\.\\)|canon|Sin stock|u\\.\\))/i.test(st) && st.length < 320;
                            });
                            if (sub.length > 0) continue;

                            filasVistas.add(t);
                            const imgs = Array.from(fila.querySelectorAll("img")).map(img => [
                                img.alt || "",
                                img.title || "",
                                img.src || ""
                            ]);
                            filasMayoristas.push({ texto: t, imgs: imgs });
                        }

                        resultados.push({
                            titulo: titulo,
                            url: urlProd,
                            textoTarjeta: textoTarjeta,
                            mayoristas: filasMayoristas
                        });
                    }
                    return resultados;
                }""")

                if not diagnostico_hecho and tarjetas_pagina:
                    diagnostico_hecho = True
                    t0 = tarjetas_pagina[0]
                    log(f"   [Diagnóstico Tarjeta 1] Botones 'Ver mayoristas' pulsados: {total_botones}")
                    log(f"   [Diagnóstico Tarjeta 1] Texto tarjeta: {t0['textoTarjeta'][:350].replace(chr(10), ' | ')}")
                    log(f"   [Diagnóstico Tarjeta 1] Filas mayoristas detectadas: {t0['mayoristas']}")

                if not tarjetas_pagina:
                    break

                nuevos_reg = 0
                for prod in tarjetas_pagina:
                    url_prod = prod["url"]
                    if url_prod in urls_procesadas:
                        continue
                    urls_procesadas.add(url_prod)

                    txt_card = prod["textoTarjeta"]
                    if not es_producto_asus(txt_card):
                        continue

                    pn = ""
                    for linea in [l.strip() for l in txt_card.split("\n") if l.strip()]:
                        if linea.lower() in ["asus", "ver mayoristas", "ocultar mayoristas", "añadir", "comparar", "sin stock"]:
                            continue
                        if re.match(r"^[A-Z0-9]{4,10}-[A-Z0-9]{4,10}$", linea, re.I):
                            pn = linea
                            break
                    if not pn:
                        m_pn = re.search(r"\b([A-Z0-9]{4,10}-[A-Z0-9]{4,10})\b", prod["titulo"], re.I)
                        pn = m_pn.group(1) if m_pn else prod["titulo"][:35]

                    for m in prod["mayoristas"]:
                        txt_m = m["texto"]
                        mayorista = extraer_nombre_mayorista(m["imgs"], txt_m)

                        precios = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})\s*€", txt_m)
                        m_canon = re.search(r"\+(\d[\d\.]*[,\.]\d{2})\s*€?\s*canon", txt_m, re.I)
                        m_stock = re.search(r"\((\d+)\s*u\.\)", txt_m, re.I)

                        precio = limpiar_numero(precios[0]) if precios else None
                        canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                        stock = int(m_stock.group(1)) if m_stock else 0
                        sin_stock = "sin stock" in txt_m.lower() or precio == 0.0 or stock == 0

                        if sin_stock and not INCLUIR_SIN_STOCK:
                            continue

                        if precio is not None:
                            registros.append({
                                "Fecha": fecha_ejecucion,
                                "P/N": pn,
                                "Modelo": prod["titulo"],
                                "Mayorista": mayorista,
                                "Precio (€)": precio,
                                "Canon (€)": canon,
                                "Stock": 0 if sin_stock else stock,
                                "Estado Stock": "Sin stock" if sin_stock else "Con stock",
                                "URL": url_prod
                            })
                            nuevos_reg += 1

                log(f"Buscador '{termino}' pág. {pagina_actual}: +{nuevos_reg} ofertas mayoristas (Total acumulado: {len(registros)})")

                if len(tarjetas_pagina) < 100:
                    break
                pagina_actual += 1

        await browser.close()

    if not registros:
        log("No se extrajeron ofertas desplegadas de mayoristas.")
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

    log(f"\n✅ Completado: {len(df)} filas reales (P/N x Mayorista con stock y precio exacto) guardadas en {ruta_hist}")
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
