import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from datetime import datetime
import os

df = pd.read_excel("D:/OnePage/Data/op_sl_sem44.xlsx", sheet_name = "graficas")

df.fillna(0, inplace=True)
df = df.drop(["semana", "n_reestructuras"], axis=1)

df["dictamen"] = df[["Promesas", "Vía de solución", "Negativa Fraude", "No localizado"]].values.tolist()
df = df.drop(columns = ["Promesas", "Vía de solución", "Negativa Fraude", "No localizado"])

df["pagos_cumplidos"] = df[["pago_parcial", "al_coriente", "liquidados", "monto_reestructuras"]].values.tolist()
df = df.drop(columns = ["pago_parcial", "al_coriente", "liquidados", "monto_reestructuras"])


tipo_dictamen = ["Promesas", "Vía de Solución", "Negativa Fraude", "No Localizada"]

tipo_pagocum = ["Pago Parcial", "Al corriente", "Liquidación", "Reestructura"]

def semana_label(fecha=None):
    fecha = fecha or datetime.now()
    iso_year, iso_week, _ = fecha.isocalendar()
    semana = datetime.now().strftime("%Y%m%d")
    return f"{iso_year}Sem{iso_week - 1:02d}"

semana = semana_label()

base_dir = f"D:/OnePage/Resources/Visualizations/{semana}"

os.makedirs(base_dir, exist_ok=True)

# Colores y estilo
texto_color = "white"
barra_color_dictamen = "skyblue"

for _, row in df.iterrows():
    nombre = row["nombre"]
    nombre_cobrador = nombre.replace(" ", "")

    # ==== GRÁFICA DICTAMEN ====
    valores = [int(v) for v in row["dictamen"]]
    etiquetas = tipo_dictamen.copy()

    valores_ordenados, etiquetas_ordenadas = zip(*sorted(zip(valores, etiquetas), reverse=False))

    y_pos = np.arange(len(etiquetas_ordenadas))

    fig, ax = plt.subplots(figsize=(4.5, 3))

    # Crear degradado azul oscuro -> celeste
    for i, valor in enumerate(valores_ordenados):
        color = plt.cm.Blues(0.3 + 0.7 * (i / len(valores_ordenados)))  # escala de azul
        ax.barh(y_pos[i], valor, color=color)

    # Etiquetas de valores al final de cada barra
    for i, v in enumerate(valores_ordenados):
        ax.text(v + max(valores_ordenados)*0.02, i, f"{v:,}", va='center', color=texto_color, fontsize=12)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(etiquetas_ordenadas, color=texto_color, fontsize=12)
    ax.set_title("Dictamen", color=texto_color)
    ax.tick_params(colors=texto_color)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(texto_color)
    ax.spines['bottom'].set_color(texto_color)
    plt.tight_layout()
    plt.savefig(f"{base_dir}/{nombre_cobrador}_dictamen.png", transparent=True, dpi=250)
    plt.close()

    # ==== TABLA PAGOS CUMPLIDOS ====
    fig, ax = plt.subplots(figsize=(4.5, 3))
    ax.axis('off')

    pagos = row["pagos_cumplidos"]

    # Formatear: 3 primeros enteros, último monetario
    formateados = []
    for i, valor in enumerate(pagos):
        if i < 3:
            formateados.append(f"{int(valor):,}")
        else:
            formateados.append(f"${valor:,.2f}")

    tabla_data = [[tipo_pagocum[i], formateados[i]] for i in range(4)]
    tabla = ax.table(cellText=tabla_data,
                     colWidths=[0.5, 0.5],
                     cellLoc='center',
                     loc='center')

    tabla.scale(1.5, 2)
    for key, cell in tabla.get_celld().items():
        cell.set_edgecolor(texto_color)
        cell.set_text_props(color=texto_color)
        cell.set_facecolor('none')
        cell.set_fontsize(14)

    plt.savefig(f"{base_dir}/{nombre_cobrador}_pagoscumpli.png", transparent=True, dpi=250)
    plt.close()
