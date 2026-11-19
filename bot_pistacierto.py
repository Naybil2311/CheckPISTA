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
        context = await browser.new_context(
            viewport={"width": 1600, "height": 1000},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        )
        await context.route("**/*.{woff,woff2,ttf,eot,mp4,webm}", lambda route: route.abort())
        page = await context.new_page()

        # Capturar respuestas de red relacionadas con login/auth para diagnóstico real
        respuestas_auth = []
        async def on_response(response):
            if any(k in response.url.lower() for k in ["login", "auth", "token", "usuario", "user", "session"]):
                try:
                    cuerpo = await response.text()
                    respuestas_auth.append(f"[{response.status}] {response.url} -> {cuerpo[:250]}")
                except Exception:
                    respuestas_auth.append(f"[{response.status}] {response.url}")

        page.on("response", on_response)

        print(f"1. Accediendo a {URL_LOGIN} para iniciar sesión...")
        await page.goto(URL_LOGIN, wait_until="networkidle", timeout=60000)
        await page.wait_for_selector("input[type='password']:visible", timeout=20000)

        # Cerrar banner de cookies si existe
        try:
            btn_cookies = page.locator("button:has-text('Aceptar'), button:has-text('Entendido'), button:has-text('Accept')").first
            if await btn_cookies.is_visible():
                await btn_cookies.click()
                await page.wait_for_timeout(400)
        except Exception:
            pass

        input_pass = page.locator("input[type='password']:visible").first
        form_padre = input_pass.locator("xpath=ancestor::form")
        if await form_padre.count() > 0:
            input_user = form_padre.locator("input:not([type='password']):not([type='hidden']):not([type='checkbox']):visible").first
        else:
            input_user = page.locator("input[type='email']:visible, input[type='text']:visible").first

        # Rellenar y disparar eventos nativos de Angular (input, change, blur)
        await input_user.click()
        await input_user.fill("")
        await input_user.type(usuario, delay=50)
        await input_user.evaluate("el => { el.dispatchEvent(new Event('input', {bubbles: true})); el.dispatchEvent(new Event('change', {bubbles: true})); el.blur(); }")
        await page.wait_for_timeout(300)

        await input_pass.click()
        await input_pass.fill("")
        await input_pass.type(password, delay=50)
        await input_pass.evaluate("el => { el.dispatchEvent(new Event('input', {bubbles: true})); el.dispatchEvent(new Event('change', {bubbles: true})); el.blur(); }")
        await page.wait_for_timeout(400)

        # Marcar checkbox "Mantener Conectado" si es visible
        try:
            chk = page.locator("input[type='checkbox']:visible").first
            if await chk.count() > 0 and not await chk.is_checked():
                await chk.click()
        except Exception:
            pass

        # Hacer clic específicamente en el botón "Acceder"
        btn_acceder = page.locator("button:has-text('Acceder'):visible, input[type='submit']:visible, button[type='submit']:visible").first
        print(f"   Botón encontrado: '{await btn_acceder.inner_text() if await btn_acceder.count() > 0 else 'Ninguno'}' | Habilitado: {await btn_acceder.is_enabled() if await btn_acceder.count() > 0 else False}")

        if await btn_acceder.count() > 0:
            await btn_acceder.click(force=True)
        else:
            await input_pass.press("Enter")

        try:
            await page.wait_for_url(lambda url: "/login" not in url, timeout=15000)
        except Exception:
            await page.wait_for_timeout(5000)

        await page.wait_for_load_state("networkidle")
        print(f"2. URL tras login: {page.url}")

        if "/login" in page.url:
            print("   Respuestas de red detectadas en el login:")
            for r in respuestas_auth:
                print(f"   -> {r}")
            # Leer solo alertas/toasts o mensajes visibles reales (excluyendo modales ocultos)
            mensajes_visibles = await page.evaluate("""() => {
                return Array.from(document.querySelectorAll('div, p, span, mat-error, .alert, .toast'))
                    .filter(el => el.offsetParent !== null && el.children.length === 0 && el.innerText.trim().length > 2)
                    .map(el => el.innerText.trim());
            }""")
            print(f"   Elementos visibles en pantalla: {mensajes_visibles}")
            await browser.close()
            return False

        # Ir al catálogo y buscar ASUS
        await page.goto(URL_CATALOGO, wait_until="networkidle", timeout=60000)
        await page.wait_for_timeout(2000)

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
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(2500)
        except Exception:
            pass

        pagina_actual = 1
        while True:
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(1500)
            print(f"3. Procesando página {pagina_actual} del catálogo (URL: {page.url})...")

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
                except Exception:
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
