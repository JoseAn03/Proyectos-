#!/usr/bin/env python3
"""
Generador de datos de ejemplo para Dashboard Power BI.
Crea CSVs listos para importar desde Power BI Desktop.

Uso:
    python3 generar_datos_pbi.py

Salida:
    - reservas_ejemplo.csv : datos transaccionales de reservas
    - no_show_ejemplo.csv  : datos de no-show por día
    - vuelos_ejemplo.csv   : datos de vuelos SJO
"""

import csv
import os
import random
from datetime import datetime, timedelta

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

random.seed(42)

# ── Datos base ──
NOMBRES = [
    "Alvarado Sanchez Maria", "Brenes Campos Jose", "Calderon Rojas Carlos",
    "Diaz Fernandez Ana", "Espinoza Gomez Luis", "Fernandez Herrera Sofia",
    "Garcia Jimenez Pedro", "Gutierrez Lopez Andrea", "Hernandez Martinez Juan",
    "Jimenez Mora Lucia", "Lopez Navarro Diego", "Madrigal Orozco Pamela",
    "Molina Porras Esteban", "Montes Quiros Valeria", "Mora Rojas Fernando",
    "Navarro Salas Carolina", "Orozco Solano Adrian", "Perez Torres Gabriela",
    "Quesada Urena Miguel", "Ramirez Vargas Daniela", "Rojas Vega Andres",
    "Salazar Zuniga Mariana", "Sanchez Aguilar Felipe", "Sandoval Benavides Karen",
    "Solano Campos Mauricio", "Torres Chacon Rebeca", "Urena Cordero Steven",
    "Valverde Delgado Catalina", "Vargas Esquivel Rodrigo", "Vega Fernandez Sharon",
    "Zamora Gonzalez Alex", "Zuniga Herrera Paola",
]

CLASES = ["ECAR", "ICAR", "CCAR", "FCAR", "SCAR", "PCAR", "LCAR"]
UBICACIONES = ["SJOE71", "SJOT71", "SJOC61", "SJOT61", "SJOE01", "SJOT01"]
MARCAS = {"SJOE71": "Alamo", "SJOT71": "Alamo", "SJOC61": "Enterprise",
          "SJOT61": "Enterprise", "SJOE01": "National", "SJOT01": "National"}
AGENCIAS = [
    "ANC RENTA CAR", "EXPEDIA 11617270", "BOOKING.COM", "HERTZ",
    "ORBITZ", "TRAVELOCITY", "PAKETE 11617270", "DESPEGAR",
    "KAYAK", "PRICELINE",
]
AEROLINEAS = [
    ("American Airlines", "MIA"), ("United Airlines", "IAH"),
    ("Delta Air Lines", "ATL"), ("Delta Air Lines", "JFK"),
    ("Southwest Airlines", "FLL"), ("JetBlue", "MCO"),
    ("Air France", "CDG"), ("Iberia", "MAD"),
    ("KLM", "AMS"), ("Copa Airlines", "PTY"),
    ("Avianca", "BOG"), ("Spirit Airlines", "MCO"),
    ("Air Canada", "YYZ"), ("British Airways", "LHR"),
    ("American Airlines", "DFW"), ("United Airlines", "EWR"),
]


def generar_reservas(n=200):
    """Genera n registros de reservas."""
    rows = []
    base = datetime(2026, 6, 1, 0, 0, 0)
    for i in range(n):
        nombre = random.choice(NOMBRES)
        ubic = random.choice(UBICACIONES)
        marca = MARCAS[ubic]
        dia = random.randint(0, 29)
        hora_rec = random.randint(4, 23)
        min_rec = random.choice([0, 15, 30, 45])
        f_rec = base + timedelta(days=dia, hours=hora_rec, minutes=min_rec)
        duracion = random.randint(1, 14)
        f_ent = f_rec + timedelta(days=duracion)
        tarifa = round(random.uniform(25, 120), 2)
        agencia = random.choice(AGENCIAS)
        clase = random.choice(CLASES)
        no_show = random.choices([0, 1], weights=[85, 15])[0]
        vip = "Sí" if "E71" in ubic or "C61" in ubic or "E01" in ubic else "No"

        rows.append([
            i + 1001,
            nombre,
            marca,
            clase,
            f_rec.strftime("%Y-%m-%d %H:%M"),
            f_ent.strftime("%Y-%m-%d %H:%M"),
            ubic,
            tarifa,
            agencia,
            "Sí" if "11617270" in agencia else "No",
            no_show,
            vip,
            f_rec.strftime("%Y-%m"),
            f_rec.strftime("%A"),
            f_rec.strftime("%H"),
            duracion,
        ])
    return rows

def generar_vuelos(n=80):
    """Genera n registros de vuelos."""
    rows = []
    for _ in range(n):
        aero, dest = random.choice(AEROLINEAS)
        hora = f"{random.randint(5, 23):02d}:{random.choice(['00','15','30','45'])}"
        rows.append([aero, dest, hora])
    rows.sort(key=lambda r: r[2])
    return rows


def main():
    # ── Reservas ──
    reservas = generar_reservas(200)
    headers_res = [
        "ID_Reserva", "Nombre_Cliente", "Marca", "Clase", "Fecha_Recogida",
        "Fecha_Entrega", "Ubicacion", "Tarifa_Diaria", "Agencia",
        "Es_Expedia", "No_Show", "Es_VIP", "Mes", "Dia_Semana", "Hora", "Duracion_Dias"
    ]
    path_res = os.path.join(OUTPUT_DIR, "reservas_ejemplo.csv")
    with open(path_res, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(headers_res)
        w.writerows(reservas)
    print(f"  {path_res}: {len(reservas)} registros")

    # ── No Show ──
    path_ns = os.path.join(OUTPUT_DIR, "no_show_ejemplo.csv")
    with open(path_ns, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ID_Reserva", "Nombre_Cliente", "Marca", "Fecha", "Hora_Recogida", "Tarifa"])
        nos = [r for r in reservas if r[10] == 1]
        for r in nos:
            w.writerow([r[0], r[1], r[2], r[4][:10], r[4][11:], r[7]])
    print(f"  {path_ns}: {len(nos)} registros (No Show)")

    # ── Vuelos ──
    vuelos = generar_vuelos(60)
    path_vu = os.path.join(OUTPUT_DIR, "vuelos_ejemplo.csv")
    with open(path_vu, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Aerolinea", "Destino", "Hora"])
        w.writerows(vuelos)
    print(f"  {path_vu}: {len(vuelos)} registros")

    print("\n✅ Datos generados. Importá los CSVs desde Power BI Desktop.")


if __name__ == "__main__":
    main()
