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

ARCHIVO_PVD = "pvd_oficial.xlsx"
ARCHIVO_BAJO_PVD = "alertas_bajo_pvdr.xlsx"

TERMINOS_BUSQUEDA = [
    "EXPERTBOOK",
    "EXPERTCENTER",
    "ASUSPRO",
    "90NX",
    "90PF",
    "90PT",
    "90MS",
    "90NB",
    "90NR"
]

EXCLUIR_ACCESORIOS = [
    "MONITOR", "ZENSCREEN", "PROART PA", "RATON", "RATÓN", "MOUSE",
    "MOCHILA", "FUNDA", "CARRY BAG", "BACKPACK", "ALFOMBRILLA", "SCABBARD",
    "GARANTIA", "GARANTÍA", "WARRANTY", "PREMIUM CARE", "EXTENSION", "EXTENSIÓN",
    "ADAPTADOR", "TECLADO", "KEYBOARD", "AURICULAR", "HEADSET", "WEBCAM",
    "PROYECTOR", "ROUTER", "EXPERTWIFI", "CONMUTADOR", "SWITCH", "VIVOWATCH",
    "ZENDRIVE", "GRABADORA", "PLACA BASE", "TARJETA", "CAJA ", "TORRE PEQUEÑA",
    "FUENTE DE ALIMENTACION", "CARGADOR", "DOCK", "MALETIN", "MALETÍN"
]

MAPA_MAYORISTAS_UUID = {
    "b53ec8ba-cfaa-43a1-a7f2-c3fbab7ad6b6": "Hispamicro",
    "edf99b17-03e6-4ba2-8898-6339440e4378": "Depau",
    "bff85f34-220c-4b56-a2ce-be10776e1d35": "Infortisa",
    "0b49e0aa-edca-411d-a417-d5e0b5ce2ec6": "Valorista",
    "38d5e378-6eb5-4b9a-b9b8-efa71e36f5d2": "Esprinet",
    "64ed262e-4163-49b2-aa10-7444943d14ac": "TDSynnex",
    "403dd8f5-ed93-4411-80f9-8628218a1479": "Ibertronica",
    "be01559f-4eb3-4b1f-a855-cd598ac0d406": "DMI",
    "64920c2b-a115-43e0-81d2-8ded450e4083": "Desyman",
    "98dd0791-fb9c-4a0b-b5d1-88f4d04297dd": "ASEUROPA",
    "c571c940-2b35-4620-8253-f07d1620f016": "Inforpor",
    "8ede237b-196d-4f2b-9573-0b590bfbe5bb": "Idiomund",
    "7898e114-5073-47c7-960f-36f9c65cfd02": "Globomatik",
}

INCLUIR_SIN_STOCK = True
CONCURRENCIA_FICHAS = 8


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


def precio_oficial_valido(valor):
    if valor is None:
        return None
    if isinstance(valor, str):
        s = valor.strip().lower()
        if not s or any(
            marca in s
            for marca in [
                "no en pricelist",
                "pricelist",
                "sin precio",
                "no disponible",
                "pendiente",
                "n/a",
                "tbd",
                "-",
            ]
        ):
            return None
    num = limpiar_numero(valor)
    if num is None or num <= 0:
        return None
    return num


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


def normalizar_mayorista(pistas_fila=None, texto_fila="", html_fila=""):
    partes = [texto_fila or "", html_fila or ""]
    if pistas_fila:
        for p in pistas_fila:
            partes.append(str(p or ""))
    combinado = " ".join(partes)
    combinado_up = combinado.upper()

    for uuid_key, nombre_comercial in MAPA_MAYORISTAS_UUID.items():
        if uuid_key.lower() in combinado.lower():
            return nombre_comercial

    mapa_texto = [
        ("PISTACERO", "PISTACERO (Compra Agrupada)"),
        ("COMPRA AGRUPADA", "PISTACERO (Compra Agrupada)"),
        ("IDIOMUN", "Idiomund"),
        ("GRUPO SII", "Idiomund"),
        ("TD SYNNEX", "TDSynnex"),
        ("TDSYNNEX", "TDSynnex"),
        ("SYNNEX", "TDSynnex"),
        ("TECH DATA", "TDSynnex"),
        ("TECHDATA", "TDSynnex"),
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
        ("DMI", "DMI"),
        ("IBERTRONICA", "Ibertronica"),
        ("DESYMAN", "Desyman"),
        ("ASEUROPA", "ASEUROPA"),
        ("TICNOVA", "Ticnova"),
        ("ALSO", "ALSO"),
        ("WESTCON", "Westcon"),
        ("EXERTIS", "Exertis"),
        ("DESPEC", "Despec"),
        ("INFORTISA", "Infortisa"),
        ("UNYKA", "Unyka"),
    ]
    for clave, nombre_limpio in mapa_texto:
        if clave in combinado_up:
            return nombre_limpio

    m_uuid = re.search(
        r"/mayoristas/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
        combinado,
        re.I,
    )
    if m_uuid:
        uuid_encontrado = m_uuid.group(1).lower()
        return MAPA_MAYORISTAS_UUID.get(uuid_encontrado, uuid_encontrado)

    m_may = re.search(r"/mayoristas/([a-zA-Z0-9_-]+)\.(?:png|jpg|jpeg|svg|webp)", combinado, re.I)
    if m_may:
        token = m_may.group(1).lower()
        return MAPA_MAYORISTAS_UUID.get(token, f"Mayorista {token[:8]}")

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


async def extraer_ficha_producto(context, sem, idx, url_prod, titulo_previo, fecha_ejecucion, progreso):
    async with sem:
        page = await context.new_page()
        filas_resultado = []

        try:
            await page.goto(url_prod, wait_until="domcontentloaded", timeout=35000)
            await page.wait_for_timeout(1200)
            await cerrar_popups(page)

            datos_dom = await page.evaluate("""() => {
                const bodyTxt = document.body.innerText || "";
                const h1 = document.querySelector("h1, h2, .product-title, .titulo-producto");
                const titulo = h1 ? (h1.innerText || "").trim() : "";

                let pn = "";
                const mPn = bodyTxt.match(/P\\/N:\\s*([A-Z0-9\\-\\.]+)/i);
                if (mPn) pn = mPn[1].trim();

                const bloques = [];
                const hojasPrecio = Array.from(document.querySelectorAll("div, span, td, p, strong, b")).filter(el => {
                    const t = (el.innerText || "").trim();
                    return el.children.length === 0 && /^\\d[\\d\\.]*,\\d{2}\\s*€?$/.test(t);
                });

                const filasVistas = new Set();
                for (const hoja of hojasPrecio) {
                    let fila = hoja.closest("tr");
                    if (!fila) {
                        let cur = hoja.parentElement;
                        for (let nivel = 0; nivel < 4 && cur; nivel++) {
                            const preciosEnFila = (cur.innerText || "").match(/\\d[\\d\\.]*,\\d{2}/g) || [];
                            if (preciosEnFila.length > 2) break;
                            const tieneVisual = cur.querySelectorAll("img").length > 0 ||
                                Array.from(cur.querySelectorAll("*")).some(n => {
                                    const bg = getComputedStyle(n).backgroundImage || n.style.backgroundImage || "";
                                    return bg && bg !== "none";
                                });
                            if (tieneVisual) {
                                fila = cur;
                                break;
                            }
                            cur = cur.parentElement;
                        }
                    }

                    const contenedor = fila || hoja.parentElement?.parentElement || hoja.parentElement;
                    if (!contenedor || filasVistas.has(contenedor)) continue;
                    filasVistas.add(contenedor);

                    const pistas = [];
                    contenedor.querySelectorAll("*").forEach(nodo => {
                        if (nodo.tagName === "IMG") {
                            for (const attr of ["alt", "title", "src", "data-src", "aria-label"]) {
                                const val = nodo.getAttribute(attr);
                                if (val && val.trim()) pistas.push(val.trim());
                            }
                        }
                        const bg = getComputedStyle(nodo).backgroundImage || nodo.style.backgroundImage || "";
                        if (bg && bg !== "none") pistas.push(bg);
                        const aria = nodo.getAttribute("aria-label") || nodo.getAttribute("data-mayorista") || nodo.getAttribute("data-proveedor") || "";
                        if (aria.trim()) pistas.push(aria.trim());
                    });

                    bloques.push({
                        texto: (contenedor.innerText || "").trim(),
                        pistas: pistas,
                        html: contenedor.outerHTML.slice(0, 1000)
                    });
                }

                return { titulo: titulo, pn: pn, bloques: bloques };
            }""")

            titulo_final = datos_dom["titulo"] or titulo_previo
            pn_final = datos_dom["pn"]
            if not pn_final:
                m_pn_tit = re.search(r"\b(90[A-Z0-9]{2}[A-Z0-9\-]{4,14}|[A-Z0-9]{4,10}-[A-Z0-9]{4,10})\b", titulo_final)
                pn_final = m_pn_tit.group(1) if m_pn_tit else titulo_final[:35]

            ofertas_ficha = {}
            for b in datos_dom["bloques"]:
                txt = b["texto"]

                if re.search(r"(Sin Título|Tramitado|\bproductos\b)", txt, re.I):
                    continue
                if not re.search(r"(u\.|canon|SKU|PISTACERO|stock|Añadir|Sin stock)", txt, re.I):
                    continue

                mayorista = normalizar_mayorista(b["pistas"], txt, b["html"])

                precios = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})", txt)
                m_canon = re.search(r"\+?\s*(\d[\d\.]*[,\.]\d{2})\s*€?\s*canon", txt, re.I)
                m_stock = re.search(r"(\d+)\s*u\.", txt, re.I)

                precio = limpiar_numero(precios[0]) if precios else None
                canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                stock = int(m_stock.group(1)) if m_stock else 0
                sin_stock = (stock == 0) and (bool(re.search(r"\bsin stock\b", txt, re.I)) or precio == 0.0)

                if sin_stock and not INCLUIR_SIN_STOCK:
                    continue

                if precio is not None and precio > 0:
                    estado_stock = "Con stock" if stock > 0 else ("Sin stock" if sin_stock else "Consultar")
                    clave = (pn_final, mayorista, precio, stock, estado_stock)
                    if clave not in ofertas_ficha:
                        ofertas_ficha[clave] = {
                            "Fecha": fecha_ejecucion,
                            "P/N": pn_final,
                            "Modelo": titulo_final,
                            "Mayorista": mayorista,
                            "Precio (€)": precio,
                            "Canon (€)": canon,
                            "Stock": stock,
                            "Estado Stock": estado_stock,
                            "URL": url_prod
                        }

            filas_resultado.extend(ofertas_ficha.values())

            if idx == 0:
                log(f"   [Diagnóstico Ficha 1] P/N: {pn_final} | Filas extraídas: {len(filas_resultado)}")
                for i_b, b in enumerate(datos_dom["bloques"][:4]):
                    log(f"      -> Fila {i_b+1} pistas: {b['pistas']} | texto: {b['texto'].replace(chr(10), ' | ')[:140]}")

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

        for termino in TERMINOS_BUSQUEDA:
            url_inicio = f"https://pistacierto.com/buscador?search={quote(termino)}&size=100"
            try:
                await page.goto(url_inicio, wait_until="domcontentloaded", timeout=35000)
                await page.wait_for_timeout(1500)
            except Exception:
                continue

            if "/login" in page.url:
                log("❌ La sesión guardada en PC_SESSION ha caducado.")
                await browser.close()
                return False

            await cerrar_popups(page)
            pagina_actual = 1

            while pagina_actual <= 25:
                enlaces = await page.evaluate("""() => {
                    return Array.from(document.querySelectorAll("a[href*='/productos/detalle/']"))
                        .map(a => ({ url: a.href.split('?')[0], titulo: (a.innerText || '').trim() }))
                        .filter(x => x.titulo.length > 4);
                }""")

                if not enlaces:
                    break

                nuevos_en_pag = 0
                for item in enlaces:
                    if not es_equipo_valido(item["titulo"]):
                        continue
                    if item["url"] not in productos_encontrados:
                        productos_encontrados[item["url"]] = item["titulo"]
                        nuevos_en_pag += 1

                log(f"Buscador '{termino}' pág. {pagina_actual}: {len(enlaces)} items | +{nuevos_en_pag} equipos nuevos (Total únicos: {len(productos_encontrados)})")

                btn_next = page.locator("button:has-text('>'), a:has-text('>')").last
                if await btn_next.count() > 0:
                    try:
                        deshabilitado = await btn_next.evaluate("el => el.disabled || (el.parentElement && el.parentElement.className.includes('disabled'))")
                        if deshabilitado:
                            break
                        primer_antes = enlaces[0]["url"]
                        await btn_next.click(force=True)
                        await page.wait_for_timeout(1500)

                        primer_despues = await page.evaluate("""() => {
                            const a = document.querySelector("a[href*='/productos/detalle/']");
                            return a ? a.href.split('?')[0] : '';
                        }""")
                        if not primer_despues or primer_despues == primer_antes:
                            break
                        pagina_actual += 1
                    except Exception:
                        break
                else:
                    break

        await page.close()

        if not productos_encontrados:
            log("No se encontraron equipos ASUS.")
            await browser.close()
            return False

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

    df = pd.DataFrame(registros).drop_duplicates(subset=["P/N", "Mayorista", "Precio (€)", "Stock", "Estado Stock"])
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


def contrastar_con_pvdr(df_ofertas, ruta_pvd=ARCHIVO_PVD):
    if os.path.exists(ARCHIVO_BAJO_PVD):
        os.remove(ARCHIVO_BAJO_PVD)

    if not os.path.exists(ruta_pvd):
        log(f"Aviso: No se encontró {ruta_pvd}, se omite el contraste contra referencia.")
        return

    if ruta_pvd.lower().endswith(".csv"):
        df_pvd = pd.read_csv(ruta_pvd, sep=None, engine="python")
    else:
        xls = pd.ExcelFile(ruta_pvd)
        df_pvd = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

    df_pvd.columns = [str(c).strip() for c in df_pvd.columns]

    col_pn = next((c for c in df_pvd.columns if c.lower() in ["pn", "p/n", "part number"]), None)
    col_model = next((c for c in df_pvd.columns if c.lower() in ["sales model name", "sales_model_name", "modelo", "model"]), None)
    col_pvdr = next((c for c in df_pvd.columns if c.lower() in ["pvdr", "pvd"]), None)
    col_promo = next((c for c in df_pvd.columns if c.lower() in ["pvd promo", "pvd_promo", "promo"]), None)

    if not col_pvdr and not col_promo:
        log("Error en contraste: no se encontraron columnas 'PvD Promo' ni 'PVDR'.")
        return

    pvdr_series = df_pvd[col_pvdr].apply(precio_oficial_valido) if col_pvdr else pd.Series([None] * len(df_pvd))
    promo_series = df_pvd[col_promo].apply(precio_oficial_valido) if col_promo else pd.Series([None] * len(df_pvd))

    df_pvd["PVDR_num"] = pvdr_series
    df_pvd["PvD_Promo_num"] = promo_series

    # Prioridad: PvD Promo válido (> 0), si no PVDR válido (> 0). Ignora filas sin precio ("No en pricelist", etc.)
    df_pvd["Ref_Precio_num"] = df_pvd["PvD_Promo_num"].where(
        df_pvd["PvD_Promo_num"].notna(),
        df_pvd["PVDR_num"]
    )
    df_pvd["Tipo_Ref"] = df_pvd["PvD_Promo_num"].apply(
        lambda x: "PvD Promo" if pd.notna(x) else "PVDR"
    )

    df_pvd_validos = df_pvd.dropna(subset=["Ref_Precio_num"]).copy()
    if df_pvd_validos.empty:
        log("Aviso: No hay referencias con precio válido en la tarifa oficial tras filtrar 'No en pricelist'.")
        return

    ofertas_prep = df_ofertas.copy()
    ofertas_prep["clave_cod"] = ofertas_prep["P/N"].astype(str).str.strip().str.upper()

    cruce_total = pd.DataFrame()

    if col_pn:
        pvd_pn = df_pvd_validos.copy()
        pvd_pn["clave_cod"] = pvd_pn[col_pn].astype(str).str.strip().str.upper()
        pvd_pn = pvd_pn[pvd_pn["clave_cod"] != ""].drop_duplicates(subset=["clave_cod"], keep="first")
        m1 = pd.merge(ofertas_prep, pvd_pn, on="clave_cod", how="inner")
        cruce_total = pd.concat([cruce_total, m1], ignore_index=True)

    if col_model:
        pvd_mod = df_pvd_validos.copy()
        pvd_mod["clave_cod"] = pvd_mod[col_model].astype(str).str.strip().str.upper()
        pvd_mod = pvd_mod[pvd_mod["clave_cod"] != ""].drop_duplicates(subset=["clave_cod"], keep="first")
        m2 = pd.merge(ofertas_prep, pvd_mod, on="clave_cod", how="inner")
        cruce_total = pd.concat([cruce_total, m2], ignore_index=True)

    if cruce_total.empty:
        log("Aviso: No hubo coincidencias con referencias vigentes en la tarifa oficial.")
        return

    cruce_total = cruce_total.drop_duplicates(subset=["P/N", "Mayorista", "Precio (€)", "Stock"])

    cruce_total["Dif vs Ref (€)"] = (cruce_total["Precio (€)"] - cruce_total["Ref_Precio_num"]).round(2)
    cruce_total["Desv Ref (%)"] = (
        (cruce_total["Dif vs Ref (€)"] / cruce_total["Ref_Precio_num"]) * 100
    ).round(2)

    bajo_ref = cruce_total[cruce_total["Dif vs Ref (€)"] <= -0.01].copy()
    bajo_ref = bajo_ref.sort_values(by="Dif vs Ref (€)", ascending=True)

    if not bajo_ref.empty:
        cols_orden = [
            "P/N", col_pn, col_model, "Modelo", "Mayorista",
            "Precio (€)", "Canon (€)", "Tipo_Ref", "Ref_Precio_num",
            "PvD_Promo_num", "PVDR_num",
            "Dif vs Ref (€)", "Desv Ref (%)",
            "Stock", "Estado Stock", "Fecha", "URL"
        ]
        cols_finales = [c for i, c in enumerate(cols_orden) if c and c in bajo_ref.columns and c not in cols_orden[:i]]
        bajo_ref[cols_finales].to_excel(ARCHIVO_BAJO_PVD, index=False)
        log(f"⚠️ Detectadas {len(bajo_ref)} ofertas por debajo de la referencia (PvD Promo / PVDR) en {ARCHIVO_BAJO_PVD}.")
    else:
        log("✅ Ninguna oferta detectada por debajo de la referencia (PvD Promo / PVDR).")


def comparar_tarifas():
    if os.path.exists(ARCHIVO_CAMBIOS):
        os.remove(ARCHIVO_CAMBIOS)

    df_nuevo = pd.read_csv(ARCHIVO_NUEVO, sep=";", encoding="utf-8-sig")

    # Contrastar siempre contra la tarifa oficial (priorizando PvD Promo)
    contrastar_con_pvdr(df_nuevo)

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
