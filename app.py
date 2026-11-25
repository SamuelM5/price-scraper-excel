import os
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/catalogue/page-{}.html"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
ARCHIVO_HISTORIAL = "historial_precios.csv"


def scrapear_pagina(numero):
    url = BASE_URL.format(numero)
    response = requests.get(url, timeout=10)
    if response.status_code == 404:
        return None  # no hay más páginas
    response.raise_for_status()
    response.encoding = "utf-8"

    soup = BeautifulSoup(response.text, "html.parser")
    libros = []
    for libro in soup.select("article.product_pod"):
        precio_texto = libro.select_one("p.price_color").text
        clase_rating = libro.select_one("p.star-rating")["class"][1]
        libros.append({
            "Título": libro.h3.a["title"],
            "Precio (£)": float(precio_texto.replace("£", "").replace("Â", "")),
            "Calificación": RATINGS.get(clase_rating, 0),
            "Disponible": "Sí" if "In stock" in libro.select_one("p.availability").text else "No",
        })
    return libros


def detectar_cambios(df_nuevo):
    """Compara con la ejecución anterior y devuelve los libros que cambiaron de precio."""
    if not os.path.exists(ARCHIVO_HISTORIAL):
        return pd.DataFrame()  # primera ejecución: no hay con qué comparar

    anterior = pd.read_csv(ARCHIVO_HISTORIAL)[["Título", "Precio (£)"]]
    anterior = anterior.rename(columns={"Precio (£)": "Precio anterior"})
    nuevo = df_nuevo[["Título", "Precio (£)"]].rename(columns={"Precio (£)": "Precio nuevo"})

    union = nuevo.merge(anterior, on="Título")
    cambios = union[(union["Precio nuevo"] - union["Precio anterior"]).abs() > 0.001].copy()
    cambios["Diferencia"] = (cambios["Precio nuevo"] - cambios["Precio anterior"]).round(2)
    return cambios[["Título", "Precio anterior", "Precio nuevo", "Diferencia"]]


def dar_formato(hoja, df):
    # Encabezado con estilo
    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="1F4E78")
        celda.alignment = Alignment(horizontal="center")

    # Ancho de columnas (+6 para que quepa el botón del filtro)
    for i, columna in enumerate(df.columns, start=1):
        ancho = max(df[columna].astype(str).map(len).max(), len(columna)) + 6
        hoja.column_dimensions[get_column_letter(i)].width = min(ancho, 60)

    hoja.auto_filter.ref = hoja.dimensions
    hoja.freeze_panes = "A2"


def guardar_excel(df, cambios, archivo):
    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Libros")
        hoja = writer.sheets["Libros"]
        dar_formato(hoja, df)
        for fila in hoja.iter_rows(min_row=2, min_col=2, max_col=2):
            fila[0].number_format = "£#,##0.00"

        if not cambios.empty:
            cambios.to_excel(writer, index=False, sheet_name="Cambios")
            dar_formato(writer.sheets["Cambios"], cambios)


def main():
    todos = []
    for pagina in range(1, 51):
        libros = scrapear_pagina(pagina)
        if libros is None:
            break
        todos.extend(libros)
        print(f"Página {pagina}: {len(libros)} libros (total {len(todos)})")
        time.sleep(0.5)  # pausa cortés entre peticiones

    df = pd.DataFrame(todos)
    cambios = detectar_cambios(df)
    try:
        guardar_excel(df, cambios, "libros.xlsx")
    except PermissionError:
        print("Error: close libros.xlsx in Excel and run the script again.")
        return

    # Guardar el historial para la próxima comparación
    df[["Título", "Precio (£)"]].to_csv(ARCHIVO_HISTORIAL, index=False)

    print(f"Listo: {len(df)} libros guardados en libros.xlsx")
    print(f"Cambios de precio detectados: {len(cambios)}")


if __name__ == "__main__":
    main()