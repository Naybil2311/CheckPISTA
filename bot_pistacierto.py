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

def extraer_nombre_mayorista(imgs_alt_src, texto_fila=""):
    up = texto_fila.upper()
    if "PISTACERO" in up or "COMPRA AGRUPADA" in up:
        return "PISTACERO (Compra Agrupada)"
    for alt, src in imgs_alt_src:
        alt_limpio = (alt or "").strip()
        if alt_limpio and len(alt_limpio) > 1 and not any(x in alt_limpio.lower() for x in ["logo", "producto", "asus", "sin imagen"]):
            return alt_limpio
        if src and any(k in src.lower() for k in ["mayorista", "proveedor", "dist", "logo", "uploads"]):
            archivo = src.split("/")[-1].split("?")[0]
            nombre = re.sub(r"\.(png|jpg|jpeg|svg|webp)$", "", archivo, flags=re.I)
            nombre = re.sub(r"[-_]", " ", nombre).strip()
            if nombre and len(nombre) > 1 and not nombre.isdigit():
                return nombre.upper() if len(nombre) <= 4 else nombre.title()
    return "Mejor Precio Portal"

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
                await page.wait_for_timeout(600)
            else:
                break
        except Exception:
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(400)

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

        pagina_actual = 1
        while True:
            url_pagina = f"https://pistacierto.com/buscador?search=ASUS&page={pagina_actual}&size=100&order=1"
            print(f"1. Cargando página {pagina_actual} ({url_pagina})...")

            cargada = False
            for intento in range(3):
                try:
                    await page.goto(url_pagina, wait_until="domcontentloaded", timeout=45000)
                    await page.wait_for_timeout(3500)
                    cargada = True
                    break
                except Exception as e:
                    print(f"   Reintentando carga de página {pagina_actual} ({intento + 1}/3)...")
                    await page.wait_for_timeout(2000)

            if not cargada:
                print(f"   No se pudo cargar la página {pagina_actual} tras 3 intentos. Deteniendo paginación.")
                break

            if "/login" in page.url:
                print("❌ La sesión guardada en PC_SESSION ha caducado o fue rechazada.")
                await browser.close()
                return False

            await cerrar_popups(page)

            # Desplegar todos los botones de mayoristas / sin stock si están presentes
            botones_ver = page.locator("text=/Ver mayoristas|Sin stock/i")
            total_botones = await botones_ver.count()
            for i in range(total_botones):
                try:
                    btn = botones_ver.nth(i)
                    if await btn.is_visible():
                        await btn.click(force=True)
                        await page.wait_for_timeout(100)
                except Exception:
                    continue

            if total_botones > 0:
                await page.wait_for_timeout(1500)

            productos_pagina = await page.evaluate("""() => {
                const tarjetas = [];
                const enlacesProd = Array.from(document.querySelectorAll("a[href*='/productos/detalle/']"));
                const tarjetasVistas = new Set();

                for (const enlace of enlacesProd) {
                    const titulo = (enlace.innerText || "").trim();
                    if (!titulo || titulo.length < 5) continue;

                    let cont = enlace.parentElement;
                    let mejorCont = null;
                    for (let nivel = 0; nivel < 10 && cont; nivel++) {
                        const enlacesEnCont = cont.querySelectorAll("a[href*='/productos/detalle/']").length;
                        if (enlacesEnCont > 2) break;
                        const txt = cont.innerText || "";
                        if (/\\d[\\d\\.]*,\\d{2}/.test(txt)) {
                            mejorCont = cont;
                        }
                        cont = cont.parentElement;
                    }

                    if (!mejorCont || tarjetasVistas.has(mejorCont)) continue;
                    tarjetasVistas.add(mejorCont);

                    const textoTarjeta = mejorCont.innerText || "";
                    const urlProd = enlace.href;

                    const lineas = textoTarjeta.split("\\n").map(l => l.trim()).filter(Boolean);
                    let pn = "";
                    let esAsus = textoTarjeta.toUpperCase().includes("ASUS");

                    for (const l of lineas) {
                        if (/^[0-9A-Z]{2}[0-9A-Z]{4,8}-[0-9A-Z]{4,8}$/i.test(l)) {
                            pn = l;
                            break;
                        }
                    }

                    // Buscar filas de mayoristas que incluyan stock o logo, o el precio principal de la tarjeta
                    const filasMayoristas = [];
                    const candidatos = Array.from(mejorCont.querySelectorAll("div, tr, li")).filter(el => {
                        const t = (el.innerText || "").trim();
                        return /\\d[\\d\\.]*,\\d{2}/.test(t) && (t.includes("u.") || t.toLowerCase().includes("stock") || el.querySelectorAll("img").length > 0) && t.length < 280;
                    });

                    const filasVistas = new Set();
                    for (const fila of candidatos) {
                        const t = (fila.innerText || "").trim();
                        if (filasVistas.has(t)) continue;
                        const sub = Array.from(fila.querySelectorAll("div, tr, li")).filter(s => {
                            const st = (s.innerText || "").trim();
                            return /\\d[\\d\\.]*,\\d{2}/.test(st) && (st.includes("u.") || st.toLowerCase().includes("stock") || s.querySelectorAll("img").length > 0) && st.length < 280;
                        });
                        if (sub.length > 0) continue;
                        filasVistas.add(t);

                        const imgs = Array.from(fila.querySelectorAll("img")).map(img => [img.alt || "", img.src || ""]);
                        filasMayoristas.push({ texto: t, imgs: imgs });
                    }

                    // Si en la vista de buscador no hay subfilas con logo/stock, extraer el precio principal de la tarjeta
                    if (filasMayoristas.length === 0) {
                        const imgsTarjeta = Array.from(mejorCont.querySelectorAll("img")).map(img => [img.alt || "", img.src || ""]);
                        filasMayoristas.push({ texto: textoTarjeta, imgs: imgsTarjeta });
                    }

                    tarjetas.push({
                        titulo: titulo,
                        pn: pn,
                        esAsus: esAsus,
                        url: urlProd,
                        mayoristas: filasMayoristas
                    });
                }
                return tarjetas;
            }""")

            if not productos_pagina:
                print(f"   No hay más productos en la página {pagina_actual}. Fin del recorrido.")
                break

            nuevos_en_pagina = 0
            for prod in productos_pagina:
                if not prod["esAsus"]:
                    continue

                pn = prod["pn"] or prod["titulo"][:35]
                for m in prod["mayoristas"]:
                    txt = m["texto"]
                    imgs = m["imgs"]

                    mayorista = extraer_nombre_mayorista(imgs, txt)

                    precios_encontrados = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})", txt)
                    m_canon = re.search(r"\+(\d[\d\.]*[,\.]\d{2})\s*€?\s*canon", txt, re.I)
                    m_stock = re.search(r"(\d+)\s*u\.", txt, re.I)

                    precio = limpiar_numero(precios_encontrados[0]) if precios_encontrados else None
                    canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                    stock = int(m_stock.group(1)) if m_stock else 0
                    sin_stock = "sin stock" in txt.lower()

                    if precio is not None:
                        registros.append({
                            "Fecha": fecha_ejecucion,
                            "P/N": pn,
                            "Modelo": prod["titulo"],
                            "Mayorista": mayorista,
                            "Precio (€)": precio,
                            "Canon (€)": canon,
                            "Stock": stock,
                            "Estado Stock": "Sin stock" if sin_stock else ("Con stock" if stock > 0 else "Consultar"),
                            "URL": prod["url"]
                        })
                        nuevos_en_pagina += 1

            print(f"   Página {pagina_actual}: +{nuevos_en_pagina} registros | Total acumulado: {len(registros)}")

            if len(productos_pagina) < 100:
                break

            pagina_actual += 1

        await browser.close()

    if not registros:
        print("No se extrajeron registros.")
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

    print(f"4. Guardados {len(df)} registros únicos (P/N x Mayorista) en {ruta_hist}")
    return True

def comparar_tarifas():
    if os.path.exists(ARCHIVO_CAMBIOS):
        os.remove(ARCHIVO_CAMBIOS)

    df_nuevo = pd.read_csv(ARCHIVO_NUEVO, sep=";", encoding="utf-8-sig")

    if not os.path.exists(ARCHIVO_HISTORICO):
        os.replace(ARCHIVO_NUEVO, ARCHIVO_HISTORICO)
        print("Primera ejecución: tarifa base multimayorista guardada correctamente.")
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
        print(f"¡Detectados {len(cambios)} cambios de precio entre mayoristas!")
    else:
        print("Sin variaciones de precio en ningún mayorista.")

    os.replace(ARCHIVO_NUEVO, ARCHIVO_HISTORICO)

if __name__ == "__main__":
    if asyncio.run(ejecutar_scraping()):
        comparar_tarifas()
