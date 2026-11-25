import io
import json
import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Modelo de Fertilización - TOMATE", layout="wide")

CONFIG_FILE = "modelo_tomate_config.json"
CATALOGO_FILE = "catalogo_fertilizantes.csv"
DOSIS_FILE = "dosis_semanales.csv"

NUTRIENTES = ["N_NO3", "N_NH4", "P2O5", "K2O", "CaO", "MgO", "S", "Fe", "Mn", "Zn", "B", "Cu", "Mo", "Co"]
REQ_NUT = ["N_Total", "P2O5", "K2O", "CaO", "MgO", "S", "Fe", "Mn", "Zn", "B", "Cu", "Mo", "Co"]
ETIQUETA = {"N_Total": "N Total", "N_NO3": "N-NO3", "N_NH4": "N-NH4", "S": "S Total"}
ETAPAS = ["Etapa 1 - Vegetativo", "Etapa 2 - Desarrollo", "Etapa 3 - Llenado"]

# ------------------------------------------------------------------
# Valores iniciales tomados de "Modelo Tomate.xlsx"
# (hojas Requerimientos Base 2019 y Dosis Etapa 1/2/3 TOM)
# ------------------------------------------------------------------
CONFIG_BASE = {
    "dist_surcos": 1.2,
    "dist_plantas": 0.3,
    "plantas_sembradas": 4000,
    "prod_estandar": 5.0,
    "prod_objetivo": 8.0,
    "factor_correccion": 1.0,
    "base_requerimiento": "Producción estándar",
    "frac_no3": 0.85,
    "extraccion": {"N_Total": 2.2, "P2O5": 1.14, "K2O": 4.7, "CaO": 2.22, "MgO": 0.67, "S": 0.63,
                   "Fe": 0.032, "Mn": 0.015, "Zn": 0.008, "B": 0.006, "Cu": 0.005, "Mo": 0.0001, "Co": 0.001},
    "aporte_suelo": {n: 1.0 for n in REQ_NUT},
    "etapas": {
        "Etapa 1 - Vegetativo": {"semanas": 5, "ferti_semana": 5,
            "pct": {"N_Total": 30, "P2O5": 38, "K2O": 23, "CaO": 30, "MgO": 30, "S": 38, "Fe": 25, "Mn": 25,
                    "Zn": 25, "B": 25, "Cu": 25, "Mo": 25, "Co": 25}},
        "Etapa 2 - Desarrollo": {"semanas": 5, "ferti_semana": 5,
            "pct": {"N_Total": 40, "P2O5": 32, "K2O": 32, "CaO": 40, "MgO": 40, "S": 32, "Fe": 30, "Mn": 30,
                    "Zn": 30, "B": 30, "Cu": 30, "Mo": 30, "Co": 30}},
        "Etapa 3 - Llenado": {"semanas": 12, "ferti_semana": 5,
            "pct": {"N_Total": 30, "P2O5": 30, "K2O": 45, "CaO": 30, "MgO": 40, "S": 30, "Fe": 45, "Mn": 45,
                    "Zn": 45, "B": 45, "Cu": 45, "Mo": 45, "Co": 45}},
    },
}

CATALOGO_BASE = [
    {'Producto': 'NATURMIX -L', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MAG BORO', 'N_NO3': 10.8, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 15.5, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.1, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SOLUTEC INICIO 13-36-12', 'N_NO3': 13.0, 'N_NH4': 0.0, 'P2O5': 36.0, 'K2O': 12.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.35, 'Fe': 0.03, 'Mn': 0.02, 'Zn': 0.3, 'B': 0.03, 'Cu': 0.0, 'Mo': 0.032, 'Co': 0.0},
    {'Producto': 'SOLUTEC DESARROLLO 20-5-10', 'N_NO3': 20.0, 'N_NH4': 0.0, 'P2O5': 5.0, 'K2O': 10.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 2.0, 'Fe': 0.09, 'Mn': 0.03, 'Zn': 0.02, 'B': 0.04, 'Cu': 0.01, 'Mo': 0.034, 'Co': 0.0},
    {'Producto': 'SOLUTEC PRODUCCION 5-10-43', 'N_NO3': 5.0, 'N_NH4': 0.0, 'P2O5': 10.0, 'K2O': 43.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.23, 'Fe': 0.09, 'Mn': 0.03, 'Zn': 0.02, 'B': 0.04, 'Cu': 0.01, 'Mo': 0.034, 'Co': 0.0},
    {'Producto': 'POLY-FEED GG 13-36-13 + ME', 'N_NO3': 3.7, 'N_NH4': 9.3, 'P2O5': 36.0, 'K2O': 13.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.1, 'Mn': 0.05, 'Zn': 0.015, 'B': 0.02, 'Cu': 0.011, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'POLY - FEED GG 19-19-19 + ME', 'N_NO3': 5.5, 'N_NH4': 13.5, 'P2O5': 19.0, 'K2O': 19.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.1, 'Mn': 0.05, 'Zn': 0.015, 'B': 0.02, 'Cu': 0.011, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'POLY-FEED GG 10-10-43 + ME', 'N_NO3': 10.0, 'N_NH4': 0.0, 'P2O5': 10.0, 'K2O': 43.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.1, 'Mn': 0.05, 'Zn': 0.015, 'B': 0.02, 'Cu': 0.011, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'HAIFA - CAL(15-0-0-26)', 'N_NO3': 14.4, 'N_NH4': 1.1, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 26.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'HAIFA UP (17-44-0)', 'N_NO3': 0.0, 'N_NH4': 17.5, 'P2O5': 43.6, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.032, 'Co': 0.0},
    {'Producto': 'HAIFA MAP (12-61-0)', 'N_NO3': 0.0, 'N_NH4': 12.0, 'P2O5': 61.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MULTI NPK (13-3-43)', 'N_NO3': 13.0, 'N_NH4': 0.0, 'P2O5': 3.0, 'K2O': 43.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SOLUPOTASSE(0-0-50-18)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 50.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 18.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'DKP HAIFA', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 41.0, 'K2O': 54.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NATURQUEL FE 4,8', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 6.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'HAIFA MKP', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 52.0, 'K2O': 34.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NATURZINC', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 13.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NATURMANGANESO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 13.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'ACTION BORO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 2.25, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 15.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KLIPBORO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 20.5, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELLUS CU', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 13.5, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MOLIB-K', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 10.0, 'Co': 0.0},
    {'Producto': 'MICRORRIEGO PRODUCCION(10-3-40)', 'N_NO3': 9.7, 'N_NH4': 0.3, 'P2O5': 3.0, 'K2O': 40.0, 'CaO': 0.0, 'MgO': 1.0, 'S': 1.0, 'Fe': 0.3, 'Mn': 0.3, 'Zn': 0.3, 'B': 0.2, 'Cu': 0.03, 'Mo': 0.005, 'Co': 1.0},
    {'Producto': 'SULFATO DE MAGNESIO TÉC PQP (0-0-0-16-13)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 16.0, 'S': 13.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'AGRIMISN FERTIRRIEGO INICIO(15-30-15)', 'N_NO3': 4.0, 'N_NH4': 11.0, 'P2O5': 30.0, 'K2O': 15.0, 'CaO': 0.0, 'MgO': 16.4, 'S': 13.0, 'Fe': 0.03, 'Mn': 0.02, 'Zn': 0.1, 'B': 0.03, 'Cu': 0.02, 'Mo': 0.03, 'Co': 0.0},
    {'Producto': 'AGRIMISN FERTIRRIEGO PRODUCCION(11-6-39)', 'N_NO3': 10.0, 'N_NH4': 1.0, 'P2O5': 6.0, 'K2O': 39.0, 'CaO': 0.0, 'MgO': 16.4, 'S': 13.0, 'Fe': 3.0, 'Mn': 2.0, 'Zn': 10.0, 'B': 3.0, 'Cu': 2.0, 'Mo': 3.0, 'Co': 0.0},
    {'Producto': 'VAN IPEREN INICIO 11-32-11 + ME', 'N_NO3': 0.0, 'N_NH4': 11.0, 'P2O5': 32.0, 'K2O': 11.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 9.5, 'Fe': 0.05, 'Mn': 0.02, 'Zn': 0.01, 'B': 0.02, 'Cu': 0.01, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'VAN IPEREN DESARROLLO 20-20-20 + ME', 'N_NO3': 5.7, 'N_NH4': 14.3, 'P2O5': 20.0, 'K2O': 20.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.1, 'Mn': 0.05, 'Zn': 0.05, 'B': 0.02, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'VAN IPEREN PRODUCCION 12-5-40 + ME', 'N_NO3': 10.5, 'N_NH4': 1.5, 'P2O5': 5.0, 'K2O': 40.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'FOSFATO MONOPOTASICO SYS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 52.0, 'K2O': 34.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULFATO DE MAGNESIO SYS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 29.0, 'S': 22.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELASYS Fe (12)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 12.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELASYS Mn(12)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 12.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELASYS Zn (12)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 12.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SYS BORO K (18)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 4.0, 'K2O': 3.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 18.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELASYS Cu (12)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 12.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SOLUACTIVE INICIO(15-45-10)', 'N_NO3': 10.62, 'N_NH4': 5.04, 'P2O5': 45.5, 'K2O': 10.5, 'CaO': 0.0, 'MgO': 0.01, 'S': 0.25, 'Fe': 0.06, 'Mn': 0.06, 'Zn': 0.07, 'B': 0.02, 'Cu': 0.02, 'Mo': 0.09, 'Co': 0.0},
    {'Producto': 'SOLUACTIVE DESARROLLO (18-18-18)', 'N_NO3': 16.34, 'N_NH4': 2.24, 'P2O5': 18.7, 'K2O': 18.6, 'CaO': 0.0, 'MgO': 0.01, 'S': 1.15, 'Fe': 0.04, 'Mn': 0.07, 'Zn': 0.11, 'B': 0.02, 'Cu': 0.02, 'Mo': 0.5, 'Co': 0.0},
    {'Producto': 'SOLUACTIVE PRODUCCION(13-3-43)', 'N_NO3': 12.88, 'N_NH4': 0.3, 'P2O5': 3.0, 'K2O': 43.0, 'CaO': 0.0, 'MgO': 0.17, 'S': 0.23, 'Fe': 0.08, 'Mn': 0.06, 'Zn': 0.07, 'B': 0.02, 'Cu': 0.02, 'Mo': 0.09, 'Co': 0.0},
    {'Producto': 'HIDROINICIO 10+40+13 + ME', 'N_NO3': 3.8, 'N_NH4': 9.2, 'P2O5': 40.0, 'K2O': 13.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 1.0, 'Fe': 0.025, 'Mn': 0.01, 'Zn': 0.02, 'B': 0.02, 'Cu': 0.01, 'Mo': 0.005, 'Co': 100.0},
    {'Producto': 'HIDROFOLLAJE 30+10+10 + ME', 'N_NO3': 2.7, 'N_NH4': 27.3, 'P2O5': 10.0, 'K2O': 10.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 2.24, 'Fe': 0.025, 'Mn': 0.01, 'Zn': 0.02, 'B': 0.02, 'Cu': 0.01, 'Mo': 0.005, 'Co': 0.0},
    {'Producto': 'HIDROKOSECHA 10+5+40 + ME', 'N_NO3': 5.3, 'N_NH4': 4.7, 'P2O5': 5.0, 'K2O': 40.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 3.0, 'Fe': 0.025, 'Mn': 0.01, 'Zn': 0.02, 'B': 0.02, 'Cu': 0.01, 'Mo': 0.005, 'Co': 100.0},
    {'Producto': 'HIDROKOFI 25+4+24+ ME', 'N_NO3': 0.0, 'N_NH4': 25.0, 'P2O5': 4.0, 'K2O': 24.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.38, 'Fe': 0.03, 'Mn': 0.02, 'Zn': 0.02, 'B': 0.04, 'Cu': 0.0, 'Mo': 0.005, 'Co': 0.0},
    {'Producto': 'HIDROINICIO GEL 7+21+8+ ME (LIQ)', 'N_NO3': 0.0, 'N_NH4': 7.49815, 'P2O5': 21.391562, 'K2O': 7.920059, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.177646, 'Mn': 0.103627, 'Zn': 0.133235, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'REKOLFOS 0+19+25+ ME (LIQ)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 19.10828, 'K2O': 25.477707, 'CaO': 0.031847, 'MgO': 0.286624, 'S': 0.178344, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.127389, 'B': 0.127389, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'Ca-Bo-Mag-K + ME (LIQ)', 'N_NO3': 2.868852, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 3.278689, 'CaO': 4.098361, 'MgO': 0.819672, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.245902, 'B': 0.983607, 'Cu': 0.081967, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KRISTASOL INICIAL', 'N_NO3': 7.5, 'N_NH4': 7.5, 'P2O5': 30.0, 'K2O': 15.0, 'CaO': 0.0, 'MgO': 1.0, 'S': 1.2, 'Fe': 0.04, 'Mn': 0.02, 'Zn': 0.06, 'B': 0.02, 'Cu': 0.01, 'Mo': 0.005, 'Co': 0.0},
    {'Producto': 'KRISTASOL CRECIMIENTO', 'N_NO3': 12.5, 'N_NH4': 12.5, 'P2O5': 10.0, 'K2O': 10.0, 'CaO': 0.0, 'MgO': 1.0, 'S': 1.0, 'Fe': 0.04, 'Mn': 0.02, 'Zn': 0.02, 'B': 0.01, 'Cu': 0.01, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'KRISTASOL DESARROLLO', 'N_NO3': 9.0, 'N_NH4': 9.0, 'P2O5': 6.0, 'K2O': 18.0, 'CaO': 0.0, 'MgO': 2.0, 'S': 7.0, 'Fe': 0.04, 'Mn': 0.02, 'Zn': 0.02, 'B': 0.01, 'Cu': 0.01, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'KRISTASOL PRODUCCIÓN', 'N_NO3': 11.7, 'N_NH4': 1.3, 'P2O5': 6.0, 'K2O': 40.0, 'CaO': 0.0, 'MgO': 1.0, 'S': 0.0, 'Fe': 0.04, 'Mn': 0.02, 'Zn': 0.02, 'B': 0.01, 'Cu': 0.01, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'KRISTASOL MULTIPROPOSITO', 'N_NO3': 9.0, 'N_NH4': 9.0, 'P2O5': 18.0, 'K2O': 18.0, 'CaO': 0.0, 'MgO': 1.0, 'S': 1.0, 'Fe': 0.04, 'Mn': 0.02, 'Zn': 0.02, 'B': 0.01, 'Cu': 0.01, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': "PRODUCCION D'AGRI", 'N_NO3': 12.0, 'N_NH4': 1.0, 'P2O5': 5.0, 'K2O': 40.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.1, 'Mn': 0.05, 'Zn': 0.02, 'B': 0.02, 'Cu': 0.011, 'Mo': 0.001, 'Co': 0.0},
    {'Producto': 'NITRATO P+K(13-3-43) BEDIS', 'N_NO3': 13.0, 'N_NH4': 0.0, 'P2O5': 3.0, 'K2O': 43.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'ACIDO FOSFORICO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 27.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'ACIDO FOSFORICO PQP', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 40.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'ACIDO NÍTRICO', 'N_NO3': 55.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'FOSFATO DIAMONICO (DAP)', 'N_NO3': 0.0, 'N_NH4': 18.0, 'P2O5': 20.2, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULFATO DE MAGNESIO BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 16.4, 'S': 13.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NITRATO DE MAGNESIO', 'N_NO3': 10.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 16.1, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NITRATO DE AMONIO', 'N_NO3': 10.5, 'N_NH4': 10.5, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULFATO DE AMONIO', 'N_NO3': 0.0, 'N_NH4': 21.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 24.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'UREA', 'N_NO3': 0.0, 'N_NH4': 46.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE POTASIO BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 50.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 18.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE CALCIO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE HIERRO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 10.0, 'Fe': 19.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE MANGANESO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE ZINC', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 13.7, 'Fe': 28.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'SULF DE COBRE', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'CALCINIT', 'N_NO3': 14.2, 'N_NH4': 1.2, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 26.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MioQ Fe BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 13.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MioQ Mn BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 13.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MioQ Zn BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 15.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'Mio B BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 20.5, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOQUEL Fe', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 12.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOQUEL Mn', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 12.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOQUEL Zn', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 12.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMO ION B', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 4.0, 'K2O': 3.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 20.5, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOQUEL Cu', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 12.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'KELATEX Mg', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 9.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'BORAX TÉCNICO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 10.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MioQ Cu BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 15.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'MOLIBDATO DE AMONIO', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 90.0, 'Co': 0.0},
    {'Producto': 'NITRATO DE MAGNESIO LÍQUIDO', 'N_NO3': 8.5, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 12.6, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NUTRIPLAN MENORES', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NITRATO DE CALCIO LIQ.COSMOAGRO', 'N_NO3': 14.4, 'N_NH4': 1.5, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 18.72, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'CLORURO DE POTASIO STD (KCL)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 61.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'UREA FOSFATO BEDIS', 'N_NO3': 0.0, 'N_NH4': 17.0, 'P2O5': 44.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'FOSFATO MONO POTASICO BEDIS', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 52.0, 'K2O': 34.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOQUEL BALANCE MENORES (PPM)', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOR 14-8-19', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'COSMOCEL 20-30-10', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'PRODUCCION SOLUCIONES NUTRITIVAS', 'N_NO3': 0.0, 'N_NH4': 6.0, 'P2O5': 7.0, 'K2O': 42.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 3.3, 'Fe': 0.1, 'Mn': 0.0, 'Zn': 0.08, 'B': 0.04, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NITROSMART', 'N_NO3': 35.0, 'N_NH4': 5.0, 'P2O5': 0.0, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 6.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'PEKACID', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 60.0, 'K2O': 20.0, 'CaO': 0.0, 'MgO': 0.0, 'S': 0.0, 'Fe': 0.0, 'Mn': 0.0, 'Zn': 0.0, 'B': 0.0, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
    {'Producto': 'NUTRIFEED MENORES', 'N_NO3': 0.0, 'N_NH4': 0.0, 'P2O5': 0.3, 'K2O': 0.0, 'CaO': 0.0, 'MgO': 13.2, 'S': 11.0, 'Fe': 1.0, 'Mn': 0.5, 'Zn': 0.04, 'B': 0.06, 'Cu': 0.04, 'Mo': 0.02, 'Co': 0.0},
    {'Producto': 'NUTRIFEED INICIO', 'N_NO3': 1.0, 'N_NH4': 11.0, 'P2O5': 30.0, 'K2O': 10.0, 'CaO': 0.0, 'MgO': 0.4, 'S': 8.0, 'Fe': 0.05, 'Mn': 0.08, 'Zn': 0.03, 'B': 0.05, 'Cu': 0.0, 'Mo': 0.0, 'Co': 0.0},
]



# ------------------------------------------------------------------
# PERSISTENCIA (archivos locales, igual que la app HORIBA)
# ------------------------------------------------------------------
def cargar_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in CONFIG_BASE.items():
            cfg.setdefault(k, v)
        return cfg
    return json.loads(json.dumps(CONFIG_BASE))


def guardar_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def cargar_catalogo():
    if os.path.exists(CATALOGO_FILE):
        df = pd.read_csv(CATALOGO_FILE)
    else:
        df = pd.DataFrame(CATALOGO_BASE)
        df.to_csv(CATALOGO_FILE, index=False)
    df = df.drop(columns=[c for c in ["Precio_kg"] if c in df.columns])
    for c in NUTRIENTES:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


COLS_DOSIS = ["Etapa", "Producto", "Dosis_lote_semana"]


def cargar_dosis():
    # Cada etapa empieza vacía: el usuario elige sus fertilizantes.
    if os.path.exists(DOSIS_FILE):
        df = pd.read_csv(DOSIS_FILE)
        if all(c in df.columns for c in COLS_DOSIS):
            return df[COLS_DOSIS]
    return pd.DataFrame(columns=COLS_DOSIS)


def guardar_dosis(df):
    df[COLS_DOSIS].to_csv(DOSIS_FILE, index=False)


# ------------------------------------------------------------------
# CÁLCULOS DEL MODELO
# ------------------------------------------------------------------
def diseno(cfg):
    dens = (100 / cfg["dist_surcos"]) * (100 / cfg["dist_plantas"])
    area = cfg["plantas_sembradas"] / dens
    ton_std = dens * cfg["prod_estandar"] / 1000
    ton_obj = dens * cfg["prod_objetivo"] / 1000
    ton_base = ton_std if cfg["base_requerimiento"] == "Producción estándar" else ton_obj
    return {"densidad": dens, "area_ha": area, "ton_std": ton_std, "ton_obj": ton_obj, "ton_base": ton_base,
            "plantas": cfg["plantas_sembradas"]}


def requerimiento_ciclo(cfg, d):
    return {n: cfg["extraccion"][n] * d["ton_base"] * cfg["factor_correccion"] * cfg["aporte_suelo"][n]
            for n in REQ_NUT}


def requerimiento_etapa(cfg, req, etapa):
    pct = cfg["etapas"][etapa]["pct"]
    out = {}
    n_etapa = req["N_Total"] * pct["N_Total"] / 100
    out["N_NO3"] = n_etapa * cfg["frac_no3"]
    out["N_NH4"] = n_etapa * (1 - cfg["frac_no3"])
    for n in REQ_NUT[1:]:
        out[n] = req[n] * pct[n] / 100
    return out


def calcular_dosis(df_dosis_etapa, catalogo, d, semanas, ferti_sem):
    """Dosis_lote_semana = kg (o L) por semana para TODAS las plantas del lote.
    Internamente se pasa a kg/ha solo para el balance contra el requerimiento."""
    df = df_dosis_etapa[["Producto", "Dosis_lote_semana"]].merge(catalogo, on="Producto", how="left").fillna(0.0)
    df["Dosis_lote_semana"] = pd.to_numeric(df["Dosis_lote_semana"], errors="coerce").fillna(0.0)
    df = df[df["Dosis_lote_semana"] > 0].copy()
    df["Dosis_semana"] = df["Dosis_lote_semana"] / d["area_ha"]
    df["Dosis_kg_ha"] = df["Dosis_semana"] * semanas
    for n in NUTRIENTES:
        df[n] = df["Dosis_kg_ha"] * df[n] / 100
    df["Lote_semana_kg"] = df["Dosis_lote_semana"]
    df["G_planta_semana"] = df["Dosis_lote_semana"] * 1000 / max(d["plantas"], 1)
    df["Por_aplicacion_kg"] = df["Lote_semana_kg"] / max(ferti_sem, 1)
    df["Cantidad_lote_kg"] = df["Lote_semana_kg"] * semanas
    return df


# Nutrientes que usa el cálculo automático y su peso (macros pesan más que menores)
OPT_PESOS = {"N_Total": 1.0, "P2O5": 1.0, "K2O": 1.0, "CaO": 1.0, "MgO": 1.0, "S": 0.7,
             "Fe": 0.4, "Mn": 0.4, "Zn": 0.4, "B": 0.4, "Cu": 0.4, "Mo": 0.2}


def calcular_dosis_automatica(productos, catalogo, req_semana, nutrientes_usar, iteraciones=4000):
    """Mínimos cuadrados no negativos sobre el error RELATIVO de cada nutriente.
    Devuelve kg por semana (en las mismas unidades del requerimiento) de cada producto."""
    import numpy as np
    cat = catalogo.set_index("Producto")
    filas, pesos, metas = [], [], []
    for n in nutrientes_usar:
        meta = req_semana.get(n, 0.0)
        if meta <= 0:
            continue
        if n == "N_Total":
            fila = [(cat.at[p, "N_NO3"] + cat.at[p, "N_NH4"]) / 100 for p in productos]
        else:
            fila = [cat.at[p, n] / 100 for p in productos]
        filas.append(fila); pesos.append(OPT_PESOS.get(n, 0.5)); metas.append(meta)
    if not filas:
        return {p: 0.0 for p in productos}
    A = np.array(filas, dtype=float); b = np.array(metas, dtype=float); w = np.array(pesos, dtype=float)
    Aw = A / b[:, None] * w[:, None]   # error relativo ponderado
    bw = w.copy()
    H = Aw.T @ Aw; g0 = Aw.T @ bw
    x = np.zeros(len(productos))
    diag = np.diag(H).copy()
    for _ in range(iteraciones):          # descenso por coordenadas con x >= 0
        cambio = 0.0
        for j in range(len(x)):
            if diag[j] <= 0:
                continue
            nuevo = max(0.0, x[j] - (H[j] @ x - g0[j]) / diag[j])
            cambio = max(cambio, abs(nuevo - x[j])); x[j] = nuevo
        if cambio < 1e-9:
            break
    return {p: round(float(v), 3) for p, v in zip(productos, x)}


# Conversión a forma elemental / iónica para expresar ppm (mg/L)
PPM_FORMAS = [  # (etiqueta, nutriente del modelo, factor)
    ("N total", "N_Total", 1.0), ("N-NO3", "N_NO3", 1.0), ("NO3⁻ (como ion, HORIBA)", "N_NO3", 4.427),
    ("N-NH4", "N_NH4", 1.0), ("P", "P2O5", 0.4364), ("K", "K2O", 0.8301), ("Ca", "CaO", 0.7147),
    ("Mg", "MgO", 0.6030), ("S", "S", 1.0), ("Fe", "Fe", 1.0), ("Mn", "Mn", 1.0), ("Zn", "Zn", 1.0),
    ("B", "B", 1.0), ("Cu", "Cu", 1.0), ("Mo", "Mo", 1.0)]


def ppm_de(kg_nutriente_por_ferti, litros):
    """kg de nutriente disueltos en una fertilización del lote -> mg/L"""
    return kg_nutriente_por_ferti * 1e6 / litros if litros > 0 else 0.0


def tabla_cumplimiento(sol, aportes):
    calc = {n: float(aportes[n].sum()) if n in aportes else 0.0 for n in NUTRIENTES}
    filas = [{"Nutriente": "N Total", "Solicitado": sol["N_NO3"] + sol["N_NH4"],
              "Calculado": calc["N_NO3"] + calc["N_NH4"]}]
    for n in NUTRIENTES:
        filas.append({"Nutriente": ETIQUETA.get(n, n), "Solicitado": sol.get(n, 0.0), "Calculado": calc[n]})
    t = pd.DataFrame(filas)
    t["Cumplimiento_%"] = t.apply(lambda r: r["Calculado"] / r["Solicitado"] * 100 if r["Solicitado"] else None, axis=1)
    t["Diferencia"] = t["Calculado"] - t["Solicitado"]
    return t


def color_cumpl(v):
    if v is None or pd.isna(v):
        return ""
    if 90 <= v <= 110:
        return "background-color: #d4edda"
    if 75 <= v <= 125:
        return "background-color: #fff3cd"
    return "background-color: #f8d7da"


# ------------------------------------------------------------------
# CARGA
# ------------------------------------------------------------------
cfg = cargar_config()
catalogo = cargar_catalogo()
df_dosis = cargar_dosis()
d = diseno(cfg)
req = requerimiento_ciclo(cfg, d)

with st.sidebar:
    st.header("💾 Respaldo de tu programa")
    st.caption("En Streamlit Cloud los datos se pueden borrar al reiniciarse la app. "
               "Descarga tu programa y vuelve a cargarlo cuando lo necesites.")
    st.download_button("📥 Descargar mi fertilización (CSV)", data=df_dosis.to_csv(index=False).encode("utf-8"),
                       file_name=f"fertilizacion_tomate_{datetime.now():%Y%m%d}.csv", mime="text/csv")
    subido = st.file_uploader("Cargar fertilización guardada", type="csv")
    if subido is not None and st.button("♻️ Restaurar este archivo"):
        df_up = pd.read_csv(subido)
        if all(c in df_up.columns for c in COLS_DOSIS):
            guardar_dosis(df_up)
            for k in [k for k in st.session_state if k.startswith(("dosis_mem_", "sel_"))]:
                del st.session_state[k]
            st.success("✅ Fertilización restaurada")
            st.rerun()
        else:
            st.error("El archivo no tiene las columnas Etapa, Producto, Dosis_lote_semana.")

st.title("🍅 Modelo de Fertilización - Cultivo de TOMATE")
st.write("Diseño agronómico, requerimientos por etapa fenológica, cálculo de dosis con hidrosolubles.")

tab_dis, tab_req, tab_dosis, tab_res, tab_cat = st.tabs([
    "🌱 Diseño del Cultivo", "📆 Requerimientos por Etapa", "🧪 Fertilización Semanal",
    "📋 Resumen del Ciclo", "📚 Catálogo de Fertilizantes"])

# ==========================================
# PESTAÑA 1: DISEÑO DEL CULTIVO
# ==========================================
with tab_dis:
    st.header("Diseño agronómico")
    with st.form("form_diseno"):
        c1, c2, c3 = st.columns(3)
        with c1:
            dist_surcos = st.number_input("Distancia entre surcos / caminos (m)", 0.1, 5.0, float(cfg["dist_surcos"]), 0.05)
            dist_plantas = st.number_input("Espaciamiento entre plantas (m)", 0.05, 3.0, float(cfg["dist_plantas"]), 0.01)
            plantas = st.number_input("Plantas sembradas en el lote", 1, 10_000_000, int(cfg["plantas_sembradas"]), 100)
        with c2:
            prod_std = st.number_input("Producción estándar (kg/planta)", 0.1, 30.0, float(cfg["prod_estandar"]), 0.5)
            prod_obj = st.number_input("Producción objetivo (kg/planta)", 0.1, 30.0, float(cfg["prod_objetivo"]), 0.5)
            fc = st.number_input("Factor de corrección (F.C.)", 0.1, 3.0, float(cfg["factor_correccion"]), 0.05)
        with c3:
            base = st.radio("Calcular requerimiento con:", ["Producción estándar", "Producción objetivo"],
                            index=0 if cfg["base_requerimiento"] == "Producción estándar" else 1)
            frac_no3 = st.slider("Fracción del N como nitrato (NO3)", 0.0, 1.0, float(cfg["frac_no3"]), 0.05)

        st.markdown("**Extracción (kg por tonelada producida) y % de aporte a suplir con fertilizante**")
        df_ext = pd.DataFrame({
            "Nutriente": [ETIQUETA.get(n, n) for n in REQ_NUT],
            "Extraccion_kg_ton": [cfg["extraccion"][n] for n in REQ_NUT],
            "Factor_suelo": [cfg["aporte_suelo"][n] for n in REQ_NUT],
        })
        df_ext_ed = st.data_editor(df_ext, hide_index=True, width="stretch", disabled=["Nutriente"],
                                   key="ed_ext",
                                   column_config={"Factor_suelo": st.column_config.NumberColumn(
                                       "Factor suelo (1 = 100%)", help="Reduce si el análisis de suelo aporta parte del nutriente")})
        guardar = st.form_submit_button("💾 Guardar diseño")
    if guardar:
        cfg.update({"dist_surcos": dist_surcos, "dist_plantas": dist_plantas, "plantas_sembradas": int(plantas),
                    "prod_estandar": prod_std, "prod_objetivo": prod_obj, "factor_correccion": fc,
                    "base_requerimiento": base, "frac_no3": frac_no3})
        cfg["extraccion"] = dict(zip(REQ_NUT, df_ext_ed["Extraccion_kg_ton"].astype(float)))
        cfg["aporte_suelo"] = dict(zip(REQ_NUT, df_ext_ed["Factor_suelo"].astype(float)))
        guardar_config(cfg)
        st.success("✅ Diseño guardado")
        st.rerun()

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Densidad de siembra", f"{d['densidad']:,.0f} pl/ha")
    m2.metric("Área del lote", f"{d['area_ha']:.3f} ha")
    m3.metric("Prod. estándar", f"{d['ton_std']:,.1f} t/ha")
    m4.metric("Prod. objetivo", f"{d['ton_obj']:,.1f} t/ha")

    st.subheader(f"Requerimiento del ciclo (kg/ha) — base: {cfg['base_requerimiento'].lower()}")
    df_req = pd.DataFrame({"Nutriente": [ETIQUETA.get(n, n) for n in REQ_NUT],
                           "kg/ha-ciclo": [req[n] for n in REQ_NUT],
                           "kg en el lote": [req[n] * d["area_ha"] for n in REQ_NUT],
                           "g por planta-ciclo": [req[n] * d["area_ha"] * 1000 / d["plantas"] for n in REQ_NUT]})
    st.dataframe(df_req.style.format({"kg/ha-ciclo": "{:,.3f}", "kg en el lote": "{:,.3f}", "g por planta-ciclo": "{:,.3f}"}),
                 hide_index=True, width="stretch")

# ==========================================
# PESTAÑA 2: REQUERIMIENTOS POR ETAPA
# ==========================================
with tab_req:
    st.header("Distribución del requerimiento por etapa fenológica")
    st.write("✏️ Edita el % de cada nutriente por etapa, las semanas y las fertilizaciones por semana. Cada columna debería sumar 100%.")
    filas = []
    for e in ETAPAS:
        fila = {"Etapa": e, "Semanas": cfg["etapas"][e]["semanas"], "Ferti/semana": cfg["etapas"][e]["ferti_semana"]}
        fila.update({ETIQUETA.get(n, n): cfg["etapas"][e]["pct"][n] for n in REQ_NUT})
        filas.append(fila)
    df_pct = pd.DataFrame(filas)
    df_pct_ed = st.data_editor(df_pct, hide_index=True, width="stretch", disabled=["Etapa"], key="ed_pct")

    sumas = {ETIQUETA.get(n, n): df_pct_ed[ETIQUETA.get(n, n)].sum() for n in REQ_NUT}
    malas = {k: v for k, v in sumas.items() if abs(v - 100) > 0.01}
    if malas:
        st.warning("⚠️ Estos nutrientes no suman 100% entre las 3 etapas: " +
                   ", ".join(f"{k} = {v:g}%" for k, v in malas.items()))
    if st.button("💾 Guardar distribución por etapas"):
        for _, r in df_pct_ed.iterrows():
            e = r["Etapa"]
            cfg["etapas"][e]["semanas"] = int(r["Semanas"])
            cfg["etapas"][e]["ferti_semana"] = int(r["Ferti/semana"])
            cfg["etapas"][e]["pct"] = {n: float(r[ETIQUETA.get(n, n)]) for n in REQ_NUT}
        guardar_config(cfg)
        st.success("✅ Distribución guardada")
        st.rerun()

    total_sem = sum(cfg["etapas"][e]["semanas"] for e in ETAPAS)
    st.caption(f"Duración total de las etapas: {total_sem} semanas ({total_sem * 7} días)")

    st.subheader(f"Requerimiento por etapa (kg para {d['plantas']:,} plantas)")
    filas = []
    for e in ETAPAS:
        r = requerimiento_etapa(cfg, req, e)
        a = d["area_ha"]
        fila = {"Etapa": e, "N Total": (r["N_NO3"] + r["N_NH4"]) * a}
        fila.update({ETIQUETA.get(n, n): r[n] * a for n in NUTRIENTES})
        filas.append(fila)
    df_req_et = pd.DataFrame(filas)
    st.dataframe(df_req_et.style.format({c: "{:,.3f}" for c in df_req_et.columns if c != "Etapa"}),
                 hide_index=True, width="stretch")
    df_plot = df_req_et.melt(id_vars="Etapa", value_vars=["N Total", "P2O5", "K2O", "CaO", "MgO", "S Total"],
                             var_name="Nutriente", value_name="kg lote")
    st.plotly_chart(px.bar(df_plot, x="Nutriente", y="kg lote", color="Etapa", barmode="group"),
                    width="stretch")

# ==========================================
# PESTAÑA 3: FERTILIZACIÓN SEMANAL
# ==========================================
with tab_dosis:
    st.header("Arma tu fertilización semanal")
    etapa = st.radio("Etapa", ETAPAS, horizontal=True, key="etapa_sel")
    info = cfg["etapas"][etapa]
    sem, fxs = max(int(info["semanas"]), 1), max(int(info["ferti_semana"]), 1)
    sol = requerimiento_etapa(cfg, req, etapa)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Duración", f"{sem} semanas")
    c2.metric("Fertilizaciones / semana", fxs)
    c3.metric("Plantas del lote", f"{d['plantas']:,}")
    c4.metric("Total fertilizaciones", sem * fxs)
    st.caption(f"Todas las dosis son para tus {d['plantas']:,} plantas (se cambian en la pestaña Diseño del Cultivo).")

    # Dosis en memoria por etapa (para no perder lo escrito al agregar productos)
    ss_key = f"dosis_mem_{etapa}"
    if ss_key not in st.session_state:
        guardadas = df_dosis[df_dosis["Etapa"] == etapa]
        st.session_state[ss_key] = dict(zip(guardadas["Producto"], guardadas["Dosis_lote_semana"].astype(float)))
    memoria = st.session_state[ss_key]

    st.subheader("1. Elige tus fertilizantes")
    elegidos = st.multiselect("Fertilizantes para esta etapa (escribe para buscar)",
                              options=sorted(catalogo["Producto"].tolist()),
                              default=[p for p in memoria if p in set(catalogo["Producto"])],
                              key=f"sel_{etapa}", placeholder="Ej: Calcinit, Nitrato de potasio, MKP...")

    st.subheader("2. Dosis por semana")
    ver_key = f"ver_{etapa}"
    st.session_state.setdefault(ver_key, 0)
    fac = d["area_ha"] / sem   # kg/ha-etapa -> kg por semana para las plantas del lote
    req_sem = {"N_Total": (sol["N_NO3"] + sol["N_NH4"]) * fac, **{n: sol[n] * fac for n in REQ_NUT[1:]}}
    with st.expander("🧮 Calcular dosis automáticamente con los fertilizantes elegidos", expanded=True):
        st.caption("El programa busca la cantidad por semana de cada fertilizante que más se acerca al "
                   "requerimiento semanal. Elige qué nutrientes quieres cubrir.")
        nut_usar = st.multiselect("Nutrientes a cubrir", options=list(OPT_PESOS),
                                  default=["N_Total", "P2O5", "K2O", "CaO", "MgO", "S"],
                                  format_func=lambda n: ETIQUETA.get(n, n), key=f"nut_{etapa}")
        if st.button("🧮 Calcular cuánto aplicar por semana", disabled=not elegidos, key=f"calc_{etapa}"):
            sugeridas = calcular_dosis_automatica(elegidos, catalogo, req_sem, nut_usar)
            st.session_state[ss_key] = sugeridas
            memoria = sugeridas
            st.session_state[ver_key] += 1
            ceros = [p for p, v in sugeridas.items() if v == 0]
            st.success("✅ Dosis calculadas. Puedes ajustarlas a mano en la tabla y luego guardar.")
            if ceros:
                st.info("Estos productos quedaron en 0 porque no ayudan a cubrir los nutrientes elegidos "
                        "o los cubren mejor otros: " + ", ".join(ceros))
    if not elegidos:
        st.info("👆 Elige uno o más fertilizantes para empezar.")
        df_in = pd.DataFrame({"Producto": [], "Dosis_lote_semana": []})
    else:
        df_in = pd.DataFrame({"Producto": elegidos, "Dosis_lote_semana": [float(memoria.get(p, 0.0)) for p in elegidos]})
    editado = st.data_editor(
        df_in, hide_index=True, width="stretch", disabled=["Producto"],
        key=f"ed_{etapa}_{st.session_state[ver_key]}_{abs(hash(tuple(elegidos)))}",
        column_config={
            "Producto": st.column_config.TextColumn("Fertilizante", width="large"),
            "Dosis_lote_semana": st.column_config.NumberColumn(f"kg por semana para {d['plantas']:,} plantas",
                                                               min_value=0.0, step=0.1, format="%.3f",
                                                               help="Para líquidos: litros por semana"),
        })
    editado["Dosis_lote_semana"] = pd.to_numeric(editado["Dosis_lote_semana"], errors="coerce").fillna(0.0)
    st.session_state[ss_key] = dict(zip(editado["Producto"], editado["Dosis_lote_semana"]))

    if st.button("💾 Guardar fertilización de esta etapa", type="primary"):
        nuevo = editado.copy()
        nuevo["Etapa"] = etapa
        df_dosis = pd.concat([df_dosis[df_dosis["Etapa"] != etapa], nuevo[COLS_DOSIS]], ignore_index=True)
        guardar_dosis(df_dosis)
        st.success("✅ Fertilización guardada")

    res = calcular_dosis(editado, catalogo, d, sem, fxs)
    tabla = tabla_cumplimiento(sol, res)
    for c in ["Solicitado", "Calculado", "Diferencia"]:
        tabla[c] = tabla[c] * fac
    tabla["g/planta requerido"] = tabla["Solicitado"] * 1000 / d["plantas"]
    tabla["g/planta aportado"] = tabla["Calculado"] * 1000 / d["plantas"]
    tabla = tabla.rename(columns={"Solicitado": "Requerido kg/semana", "Calculado": "Aportado kg/semana",
                                  "Diferencia": "Diferencia kg/semana"})
    tabla = tabla[["Nutriente", "Requerido kg/semana", "Aportado kg/semana", "Cumplimiento_%", "Diferencia kg/semana",
                   "g/planta requerido", "g/planta aportado"]]

    st.subheader(f"3. ¿Cuánto cubres del requerimiento? (por semana, para {d['plantas']:,} plantas)")
    col_t, col_g = st.columns([1, 1])
    with col_t:
        st.dataframe(tabla.style.format({"Requerido kg/semana": "{:,.3f}", "Aportado kg/semana": "{:,.3f}",
                                         "Diferencia kg/semana": "{:+,.3f}", "Cumplimiento_%": "{:,.0f}%",
                                         "g/planta requerido": "{:,.3f}", "g/planta aportado": "{:,.3f}"}, na_rep="-")
                     .map(color_cumpl, subset=["Cumplimiento_%"]),
                     hide_index=True, width="stretch", height=560)
        st.caption("🟩 90–110%   🟨 75–125%   🟥 fuera de rango.  Diferencia negativa = lo que te falta por semana.")
    with col_g:
        tg = tabla.dropna(subset=["Cumplimiento_%"])
        fig = go.Figure(go.Bar(x=tg["Nutriente"], y=tg["Cumplimiento_%"],
                               marker_color=["#2e7d32" if 90 <= v <= 110 else "#f9a825" if 75 <= v <= 125 else "#c62828"
                                             for v in tg["Cumplimiento_%"]]))
        fig.add_hline(y=100, line_dash="dash")
        fig.update_layout(yaxis_title="% cumplimiento", height=560, margin=dict(t=20))
        st.plotly_chart(fig, width="stretch")

    st.subheader("4. ¿Cuántas ppm aporta en el agua de riego?")
    litros = st.number_input("Litros de agua por fertilización en el lote (L)", min_value=0.0,
                             value=float(info.get("litros_ferti", 5000.0)), step=100.0, key=f"litros_{etapa}",
                             help="Volumen total de agua en que se aplica UNA fertilización en el lote")
    if litros != float(info.get("litros_ferti", 5000.0)):
        cfg["etapas"][etapa]["litros_ferti"] = litros
        guardar_config(cfg)
    if res.empty or litros <= 0:
        st.info("Escribe dosis y litros de agua para ver las ppm.")
    else:
        n_apl = sem * fxs
        kg_ferti = {n: res[n].sum() * d["area_ha"] / n_apl for n in NUTRIENTES}  # kg nutriente por fertilización
        kg_ferti["N_Total"] = kg_ferti["N_NO3"] + kg_ferti["N_NH4"]
        req_ferti = {"N_Total": (sol["N_NO3"] + sol["N_NH4"]) * d["area_ha"] / n_apl,
                     **{n: sol[n] * d["area_ha"] / n_apl for n in NUTRIENTES}}
        filas = [{"Elemento": et, "ppm aportadas": ppm_de(kg_ferti[n] * f, litros),
                  "ppm requeridas": ppm_de(req_ferti[n] * f, litros)} for et, n, f in PPM_FORMAS]
        t_ppm = pd.DataFrame(filas)
        g_l = res["Por_aplicacion_kg"].sum() * 1000 / litros
        p1, p2 = st.columns([1, 1])
        with p1:
            st.dataframe(t_ppm.style.format({"ppm aportadas": "{:,.2f}", "ppm requeridas": "{:,.2f}"}),
                         hide_index=True, width="stretch", height=560)
        with p2:
            st.metric("Concentración total de fertilizantes", f"{g_l:,.3f} g/L")
            st.caption("ppm = mg por litro en el agua de riego durante la fertilización. "
                       "P, K, Ca y Mg se expresan como elemento (no como óxido). "
                       "NO3⁻ como ion sirve para comparar con la lectura del HORIBA en agua de riego.")
            if g_l > 2:
                st.warning("⚠️ Más de 2 g/L de fertilizantes: revisa la conductividad (CE) de la solución.")
        with st.expander("Ver ppm que aporta cada fertilizante"):
            filas = []
            for _, r in res.iterrows():
                fila = {"Producto": r["Producto"], "g/L": r["Por_aplicacion_kg"] * 1000 / litros}
                kg = {n: r[n] * d["area_ha"] / n_apl for n in NUTRIENTES}
                kg["N_Total"] = kg["N_NO3"] + kg["N_NH4"]
                for et, n, f in PPM_FORMAS:
                    fila[et] = ppm_de(kg[n] * f, litros)
                filas.append(fila)
            pp = pd.DataFrame(filas)
            pp = pp.loc[:, (pp != 0).any(axis=0)]
            st.dataframe(pp.style.format({c: "{:,.2f}" for c in pp.columns if c != "Producto"}),
                         hide_index=True, width="stretch")

    st.subheader("5. Programa semanal para el lote")
    if res.empty:
        st.info("Escribe dosis mayores a cero para ver las cantidades a aplicar.")
    else:
        prog = res[["Producto", "Lote_semana_kg", "G_planta_semana", "Por_aplicacion_kg", "Cantidad_lote_kg"]].rename(columns={
            "G_planta_semana": "g por planta por semana", "Lote_semana_kg": "Lote: kg por semana",
            "Por_aplicacion_kg": f"Lote: kg por fertilización (÷{fxs})", "Cantidad_lote_kg": f"Lote: kg toda la etapa (×{sem} sem)"})
        tot = prog.drop(columns="Producto").sum()
        prog = pd.concat([prog, pd.DataFrame([{"Producto": "TOTAL", **tot}])], ignore_index=True)
        st.dataframe(prog.style.format({c: "{:,.3f}" for c in prog.columns if c != "Producto"}),
                     hide_index=True, width="stretch")
        with st.expander("Ver aporte de nutrientes de cada fertilizante (kg por semana en el lote)"):
            ap = res[["Producto"] + NUTRIENTES].copy()
            ap[NUTRIENTES] = ap[NUTRIENTES] * fac
            st.dataframe(ap.rename(columns=ETIQUETA).style.format({ETIQUETA.get(n, n): "{:,.3f}" for n in NUTRIENTES}),
                         hide_index=True, width="stretch")

# ==========================================
# PESTAÑA 4: RESUMEN DEL CICLO
# ==========================================
with tab_res:
    st.header("Resumen del ciclo")
    resumen, aportes_ciclo, detalle = [], [], []
    for e in ETAPAS:
        info_e = cfg["etapas"][e]
        r_e = calcular_dosis(df_dosis[df_dosis["Etapa"] == e], catalogo, d,
                             max(int(info_e["semanas"]), 1), max(int(info_e["ferti_semana"]), 1))
        t_e = tabla_cumplimiento(requerimiento_etapa(cfg, req, e), r_e)
        t_e["Etapa"] = e
        aportes_ciclo.append(t_e)
        r_e["Etapa"] = e
        detalle.append(r_e)
        resumen.append({"Etapa": e, "Semanas": info_e["semanas"], "Productos": len(r_e),
                        "g/planta por semana": r_e["G_planta_semana"].sum(), "Total kg lote": r_e["Cantidad_lote_kg"].sum(),
                        "Lote kg/semana": r_e["Lote_semana_kg"].sum(),
                        "Lote kg por fertilización": r_e["Por_aplicacion_kg"].sum()})
    df_resumen = pd.DataFrame(resumen)
    k1, k2, k3 = st.columns(3)
    k1.metric("Semanas de fertilización", f"{df_resumen['Semanas'].sum()}")
    k2.metric("Plantas del lote", f"{d['plantas']:,}")
    k3.metric("Total fertilizantes del lote (kg)", f"{df_resumen['Total kg lote'].sum():,.2f}")
    st.dataframe(df_resumen.style.format({"g/planta por semana": "{:,.2f}",
                                          "Total kg lote": "{:,.2f}", "Lote kg/semana": "{:,.3f}",
                                          "Lote kg por fertilización": "{:,.3f}"}),
                 hide_index=True, width="stretch")

    st.subheader("Cumplimiento por etapa (%)")
    df_ap = pd.concat(aportes_ciclo)
    piv = df_ap.pivot_table(index="Nutriente", columns="Etapa", values="Cumplimiento_%", sort=False)
    st.dataframe(piv.style.format("{:,.0f}%", na_rep="-").map(color_cumpl), width="stretch")

    st.subheader("Productos necesarios (todo el ciclo)")
    df_det = pd.concat(detalle) if detalle else pd.DataFrame()
    if not df_det.empty:
        cot = df_det.groupby("Producto", as_index=False).agg(
            Kg_lote=("Cantidad_lote_kg", "sum")).sort_values("Kg_lote", ascending=False)
        cot["g_por_planta_ciclo"] = cot["Kg_lote"] * 1000 / d["plantas"]
        st.dataframe(cot.style.format({"Kg_lote": "{:,.2f}", "g_por_planta_ciclo": "{:,.2f}"}),
                     hide_index=True, width="stretch")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame([{**{k: v for k, v in cfg.items() if not isinstance(v, dict)}, **d}]).T.to_excel(
                writer, sheet_name="Diseño", header=["Valor"])
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen")
            cot.to_excel(writer, index=False, sheet_name="Productos")
            df_ap.to_excel(writer, index=False, sheet_name="Cumplimiento")
            df_det.to_excel(writer, index=False, sheet_name="Detalle dosis")
        st.download_button("📥 Descargar programa en Excel", data=output.getvalue(),
                           file_name=f"programa_tomate_{datetime.now():%Y%m%d}.xlsx")

# ==========================================
# PESTAÑA 5: CATÁLOGO
# ==========================================
with tab_cat:
    st.header("Catálogo de fertilizantes")
    st.write("✏️ Composición en % (N y formas como N; P como P2O5; K como K2O; Ca como CaO; Mg como MgO; "
             "S y menores como elemento). Para líquidos, el % ya está convertido a kg por litro aplicado.")
    filtro = st.text_input("🔎 Buscar producto")
    vista = catalogo[catalogo["Producto"].str.contains(filtro, case=False, na=False)] if filtro else catalogo
    cat_ed = st.data_editor(vista, num_rows="dynamic", width="stretch", hide_index=True, key="ed_cat")
    if st.button("💾 Guardar catálogo"):
        resto = catalogo[~catalogo.index.isin(vista.index)]
        nuevo = pd.concat([resto, cat_ed], ignore_index=True).dropna(subset=["Producto"])
        nuevo = nuevo.drop_duplicates(subset="Producto", keep="last").sort_values("Producto")
        nuevo.to_csv(CATALOGO_FILE, index=False)
        st.success("✅ Catálogo actualizado")
        st.rerun()
