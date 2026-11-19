import os
import re
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

URL_LOGIN = "https://pistacierto.com/login"
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
    usuario = os.environ.get("PC_USER", "").strip()
    password = os.environ.get("PC_PASS", "").strip()

    if not usuario or not password:
        print("Error: Faltan los secretos PC_USER o PC_PASS en GitHub Actions.")
        return False

    registros = []
    fecha_ejecucion = datetime.now().strftime("%Y-%m-%d %H:%M")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={"width": 1600, "height": 1000})
        await context.route("**/*.{woff,woff2,ttf,eot,mp4,webm}", lambda route: route.abort())
        page = await context.new_page()

        print(f"1. Accediendo a {URL_LOGIN} para iniciar sesión...")
        await page.goto(URL_LOGIN, wait_until="networkidle", timeout=60000)
        await page.wait_for_selector("input[type='password']", timeout=20000)

        inputs_texto = page.locator("input[type='email'], input[type='text']")
        inputs_pass = page.locator("input[type='password']")

        await inputs_texto.first.fill(usuario)
        await inputs_pass.first.fill(password)

        # Hacer clic en el botón de iniciar sesión si existe, o pulsar Enter
        btn_submit = page.locator("button[type='submit'], button:has-text('Iniciar'), button:has-text('Entrar'), button:has-text('Acceder')").first
        if await btn_submit.count() > 0:
            await btn_submit.click()
        else:
            await inputs_pass.first.press("Enter")

        # Esperar a que la URL deje de ser /login
        try:
            await page.wait_for_url(lambda url: "/login" not in url, timeout=20000)
        except Exception:
            await page.wait_for_timeout(4000)

        await page.wait_for_load_state("networkidle")
        print(f"2. Sesión iniciada correctamente. URL actual: {page.url}")

        if "/login" in page.url:
            print("Error: Sigue en /login. Verifica que PC_USER y PC_PASS sean correctos.")
            await browser.close()
            return False

        # Ir al catálogo y buscar ASUS
        await page.goto(URL_CATALOGO, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(2000)

        buscador = page.locator("input[placeholder*='Buscar'], input[type='search'], input[name*='search']").first
        if await buscador.count() > 0:
            await buscador.fill("ASUS")
            await buscador.press("Enter")
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2500)

        filtro_asus = page.locator("label:has-text('ASUS'), span:has-text('ASUS'), a:has-text('ASUS')").first
        try:
            if await filtro_asus.is_visible():
                await filtro_asus.click()
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(2500)
        except Exception:
            pass

        pagina_actual = 1
        while True:
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1500)
            print(f"3. Procesando página {pagina_actual} del catálogo (URL: {page.url})...")

            # Pulsar todos los botones "Ver mayoristas" cerrados en la página actual
            botones_ver = page.locator("text=/Ver mayoristas/i")
            total_botones = await botones_ver.count()
            for i in range(total_botones):
                try:
                    btn = botones_ver.nth(i)
                    if await btn.is_visible():
                        await btn.click()
                        await page.wait_for_timeout(200)
                except Exception:
                    continue

            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1500)

            # Extraer tarjetas con reintento si ocurre una navegación intermedia
            productos_pagina = []
            for intento in range(3):
                try:
                    productos_pagina = await page.evaluate("""() => {
                        const tarjetas = [];
                        const enlacesProd = Array.from(document.querySelectorAll("a[href*='/productos/detalle/']"));
                        const tarjetasVistas = new Set();

                        for (const enlace of enlacesProd) {
                            const titulo = (enlace.innerText || "").trim();
                            if (!titulo || titulo.length < 5) continue;

                            let cont = enlace.parentElement;
                            while (cont && cont.parentElement) {
                                const txt = cont.innerText || "";
                                if ((txt.includes("Ocultar mayoristas") || txt.includes("Ver mayoristas")) && txt.includes("€")) {
                                    break;
                                }
                                cont = cont.parentElement;
                            }
                            if (!cont || tarjetasVistas.has(cont)) continue;
                            tarjetasVistas.add(cont);

                            const textoTarjeta = cont.innerText || "";
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

                            const filasMayoristas = [];
                            const candidatos = Array.from(cont.querySelectorAll("div, tr")).filter(el => {
                                const t = el.innerText || "";
                                return t.includes("€") && (t.includes("u.)") || t.includes("Sin stock")) && t.length < 220;
                            });

                            const filasVistas = new Set();
                            for (const fila of candidatos) {
                                const t = fila.innerText.trim();
                                if (filasVistas.has(t)) continue;
                                const sub = Array.from(fila.querySelectorAll("div, tr")).filter(s => {
                                    const st = s.innerText || "";
                                    return st.includes("€") && (st.includes("u.)") || st.includes("Sin stock")) && st.length < 220;
                                });
                                if (sub.length > 0) continue;
                                filasVistas.add(t);

                                const imgs = Array.from(fila.querySelectorAll("img")).map(img => [img.alt || "", img.src || ""]);
                                filasMayoristas.push({ texto: t, imgs: imgs });
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
                    break
                except Exception as e:
                    print(f"   Reintentando lectura tras navegación interna ({intento + 1}/3)...")
                    await page.wait_for_load_state("networkidle")
                    await page.wait_for_timeout(2000)

            for prod in productos_pagina:
                if not prod["esAsus"]:
                    continue

                pn = prod["pn"] or prod["titulo"][:35]
                for m in prod["mayoristas"]:
                    txt = m["texto"]
                    imgs = m["imgs"]

                    if "PISTACERO" in txt.upper() or "COMPRA AGRUPADA" in txt.upper():
                        mayorista = "PISTACERO (Compra Agrupada)"
                    else:
                        mayorista = extraer_nombre_mayorista(imgs)

                    precios_encontrados = re.findall(r"(?<!\+)(\d[\d\.]*,\d{2})\s*€", txt)
                    m_canon = re.search(r"\+(\d[\d\.]*[,\.]\d{2})\s*€\s*canon", txt, re.I)
                    m_stock = re.search(r"\((\d+)\s*u\.\)", txt, re.I)

                    precio = limpiar_numero(precios_encontrados[0]) if precios_encontrados else None
                    canon = limpiar_numero(m_canon.group(1)) if m_canon else 0.0
                    stock = int(m_stock.group(1)) if m_stock else 0
                    sin_stock = "sin stock" in txt.lower() or stock == 0

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
                            "URL": prod["url"]
                        })

            print(f"   Acumulados hasta ahora: {len(registros)} registros (Producto x Mayorista)...")

            btn_siguiente = page.locator("button:has-text('Siguiente'), a:has-text('Siguiente'), li.next:not(.disabled) a, button[aria-label*='Next']").first
            if await btn_siguiente.count() > 0 and await btn_siguiente.is_enabled():
                try:
                    antes_count = len(registros)
                    await btn_siguiente.click()
                    await page.wait_for_load_state("networkidle")
                    await page.wait_for_timeout(2500)
                    if len(registros) == antes_count and pagina_actual > 1:
                        break
                    pagina_actual += 1
                except Exception:
                    break
            else:
                break

        await browser.close()

    if not registros:
        print("No se extrajeron registros. Revisa credenciales o filtros de búsqueda.")
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

    print(f"4. Guardados {len(df)} registros (P/N x Mayorista) en {ruta_hist}")
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
