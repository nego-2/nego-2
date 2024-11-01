import pandas as pd
from datetime import datetime
import os
import flet as ft

def filtrar_archivos(ruta, prefijo, fin):
    """Filtra los archivos en el directorio que están dentro del rango especificado."""
    try:
        archivos = os.listdir(ruta)
        archivos_filtrados = [archivo for archivo in archivos if archivo.startswith(prefijo) and fin in archivo]
        return archivos_filtrados
    except FileNotFoundError as e:
        print(f"Error al cargar el archivo {ruta}: {e}")
        return []
    except Exception as e:
        print(f"Error en el servidor: {e}")
        return []

def cargar_datos(ruta, archivo, skiprows=9, na_values=[""]):
    """Carga datos de un archivo Excel, omitiendo filas iniciales y manejando valores NA."""
    try:
        df = pd.read_excel(os.path.join(ruta, archivo), skiprows=skiprows, na_values=na_values)
        df = df[['Time Since Last.2', 'Primary Database \n -- Secondary Database']].dropna(how='all')
        return df
    except Exception as e:
        print(f"Error al cargar el archivo {archivo}: {e}")
        return pd.DataFrame()

def unificar_datos(dataframes):
    """Concatena una lista de DataFrames y divide la columna 'Time Since Last.2'."""
    try:
        join = pd.concat(dataframes)
        nom_div = join['Time Since Last.2'].str.split(expand=True)
        nom_div.columns = ['Time Since Last', 'Time Since Last.3']
        join = pd.concat([join, nom_div['Time Since Last'].astype(int)], axis=1)
        return join[['Time Since Last', 'Primary Database \n -- Secondary Database']].dropna(how='all')
    except Exception as e:
        print(f"Error al unificar datos: {e}")
        return pd.DataFrame()

def guardar_resultado(df, ruta, nombre_base):
    """Guarda el DataFrame resultante en un archivo Excel con la fecha actual en el nombre."""
    fecha_actual = datetime.now()
    fecha = fecha_actual.strftime("%Y-%m-%d %H:%M:%S")
    df.insert(2, 'Ultima Replicación', fecha)
    nombre_archivo = f"{nombre_base} - {fecha_actual.strftime('%m%d%Y')} - Union_LogShipping.xlsx"
    try:
        df.to_excel(os.path.join(ruta, nombre_archivo), index=False)
        print(f"Archivo guardado exitosamente en: {os.path.join(ruta, nombre_archivo)}")
        return True
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")
        return False

def main(page: ft.Page):
    page.title = "UNIFICADOR DE ARCHIVOS LOG SHIPPING"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.bgcolor = ft.colors.GREY_300
    page.window_min_height = 200



    ruta = ft.Ref[ft.TextField]()

    def show_dialog(title, content):
        dlg_modal.title = ft.Text(title)
        dlg_modal.content = ft.Text(content)
        dlg_modal.open = True
        page.dialog = dlg_modal
        page.update()

    def close_dlg(e):
        dlg_modal.open = False
        page.update()
        if dlg_modal.title.value == "Éxito":
            page.window_close()


    dlg_modal = ft.AlertDialog(
        modal=True,
        title=ft.Text(""),
        content=ft.Text(""),
        actions=[
            ft.TextButton("Ok", on_click=close_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def unir_rutas(event):
        ruta_path = ruta.current.value
        prefijo = "Transaction Log Shipping Status -"
        fin = ".xlsx"

        archivos = filtrar_archivos(ruta_path, prefijo, fin)

        if not archivos:
            show_dialog("Advertencia", "No se encontraron archivos en la ruta especificada.")
            page.dialog = dlg_modal
            dlg_modal.open = True
            page.update()
            return

        dataframes = [cargar_datos(ruta_path, archivo) for archivo in archivos]
        if not all([not df.empty for df in dataframes]):
            show_dialog("Error", "Error al cargar uno o más archivos.")
            page.dialog = dlg_modal
            dlg_modal.open = True
            page.update()
            return

        resultado = unificar_datos(dataframes)
        if resultado.empty:
            show_dialog("Error", "Error al unificar los datos.")
            page.dialog = dlg_modal
            dlg_modal.open = True
            page.update()
            return

        success = guardar_resultado(resultado, ruta_path, "Transaction Log Shipping Status")
        if success:
            show_dialog("Éxito", "Los archivos se unificaron y guardaron exitosamente.")
        else:
            show_dialog("Error", "Error al guardar el archivo unificado.")

        dlg_modal.open = True
        page.dialog = dlg_modal
        page.update()

    page.add(
        ft.ResponsiveRow(
            [
                ft.Container(
                    ft.TextField(label="Ingrese la ruta donde están los archivos a unificar:", ref=ruta),
                    padding=ft.padding.symmetric(horizontal=10),
                    width=50,
                    bgcolor=ft.colors.WHITE,
                    border_radius=20,
                    height=60,
                    col={"sm": 9, "md": 5, "xl": 9},
                ),
                ft.Container(
                    ft.FilledButton("Unir archivos", icon="add", on_click=unir_rutas),
                    padding=7,
                    width=50,
                    height=60,
                    col={"sm": 9, "md": 5, "xl": 2},
                ),
            ],
            run_spacing={"xs": 10},
        ),
    )

ft.app(target=main)
