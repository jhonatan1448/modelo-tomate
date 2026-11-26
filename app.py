import io
import json
import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Modelo de Fertilización - TOMATE", layout="centered")

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


def proponer_tanque(fila):
    """Regla clásica A/B: el calcio no puede ir con fosfatos ni sulfatos (precipitan)."""
    ca, p, s = fila["CaO"] >= 1, fila["P2O5"] >= 1, fila["S"] >= 1
    micro_quelato = (fila["Fe"] + fila["Mn"] + fila["Zn"] + fila["Cu"]) > 0 and not (p or s)
    if ca and (p or s):
        return "A", "⚠️ Trae calcio y también fosfato/sulfato: aplícalo aparte"
    if ca:
        return "A", "Lleva calcio"
    if p and s:
        return "B", "Lleva fosfato y sulfato"
    if p:
        return "B", "Lleva fosfato"
    if s:
        return "B", "Lleva sulfato"
    if micro_quelato:
        return "A", "Quelato de micros (mejor lejos de fosfatos)"
    return "A/B", "Compatible con ambos (se reparte para equilibrar)"


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

# ------------------------------------------------------------------
# AYUDAS PARA LA PANTALLA SENCILLA
# ------------------------------------------------------------------
NOMBRES = {"N_Total": "Nitrógeno (N)", "P2O5": "Fósforo (P)", "K2O": "Potasio (K)", "CaO": "Calcio (Ca)",
           "MgO": "Magnesio (Mg)", "S": "Azufre (S)", "Fe": "Hierro (Fe)", "Mn": "Manganeso (Mn)",
           "Zn": "Zinc (Zn)", "B": "Boro (B)", "Cu": "Cobre (Cu)", "Mo": "Molibdeno (Mo)"}
MACROS = ["N_Total", "P2O5", "K2O", "CaO", "MgO", "S"]
MICROS = ["Fe", "Mn", "Zn", "B", "Cu", "Mo"]
ICONO_ETAPA = {"Etapa 1 - Vegetativo": "🌱", "Etapa 2 - Desarrollo": "🌿", "Etapa 3 - Llenado": "🍅"}


def num(x, dec=2):
    """Número al estilo colombiano: 1.234,56"""
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cantidad(kg, liquido=False):
    if liquido:
        return f"{num(kg * 1000, 0)} mL" if kg < 1 else f"{num(kg, 2)} L"
    return f"{num(kg * 1000, 0)} g" if kg < 1 else f"{num(kg, 2)} kg"


def es_liquido(nombre):
    n = nombre.upper()
    return "LIQ" in n or "LÍQ" in n


def estado(pct):
    if pct is None:
        return "⚪", "Sin dato"
    if pct < 75:
        return "🔴", "Falta"
    if pct < 90:
        return "🟡", "Un poco bajo"
    if pct <= 110:
        return "✅", "Bien"
    if pct <= 125:
        return "🟡", "Un poco alto"
    return "🔴", "Se pasa"


def contenido(cat_row, n):
    return cat_row["N_NO3"] + cat_row["N_NH4"] if n == "N_Total" else cat_row[n]


@st.cache_data(show_spinner=False)
def dosis_auto_cache(productos, req_items, nutrientes, catalogo_df):
    return calcular_dosis_automatica(list(productos), catalogo_df, dict(req_items), list(nutrientes))


def agregar_producto(key, prod):
    actual = list(st.session_state.get(key, []))
    if prod not in actual:
        actual.append(prod)
    st.session_state[key] = actual


def guardar_si_cambia(e, clave, valor):
    if cfg["etapas"][e].get(clave) != valor:
        cfg["etapas"][e][clave] = valor
        guardar_config(cfg)


# ------------------------------------------------------------------
# ENCABEZADO
# ------------------------------------------------------------------
st.title("🍅 Fertilización del Tomate")
st.write("Elige tus fertilizantes y la app te dice **cuánto aplicar** y **cómo preparar los tanques**.")

tab_main, tab_res, tab_adv = st.tabs(["🍅 Mi fertilización", "📋 Resumen y copia", "⚙️ Configuración avanzada"])

# ==========================================
# PESTAÑA PRINCIPAL (SENCILLA)
# ==========================================
def _principal():
    global df_dosis
    with st.expander("❓ ¿Cómo se usa? (toca aquí)"):
        st.markdown(
            "1. Escribe **cuántas plantas** tienes y en qué **etapa** está el cultivo.\n"
            "2. **Elige los fertilizantes** que tienes en la finca.\n"
            "3. La app calcula **sola** cuánto aplicar y te dice si a la planta le **falta** o le **sobra** algo.\n"
            "4. Mira **cómo preparar los tanques A y B** y descarga la receta.\n\n"
            "Todo se guarda solo. Lo técnico (extracciones, % por etapa, catálogo) está en ⚙️ Configuración avanzada.")

    # ---------------- PASO 1 ----------------
    st.markdown("### 1️⃣ Tu cultivo")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        if "plantas_main" not in st.session_state:
            st.session_state["plantas_main"] = int(cfg["plantas_sembradas"])
        plantas = c1.number_input("🌱 ¿Cuántas plantas tienes?", min_value=1, step=100, key="plantas_main")
        campo_prod = "prod_estandar" if cfg["base_requerimiento"] == "Producción estándar" else "prod_objetivo"
        if "prod_main" not in st.session_state:
            st.session_state["prod_main"] = float(cfg[campo_prod])
        prod = c2.number_input("🍅 ¿Cuántos kg de tomate esperas por planta?", min_value=0.5, max_value=30.0,
                               step=0.5, format="%.1f", key="prod_main")
        if int(plantas) != cfg["plantas_sembradas"] or float(prod) != float(cfg[campo_prod]):
            cfg["plantas_sembradas"], cfg[campo_prod] = int(plantas), float(prod)
            guardar_config(cfg)
        d = diseno(cfg)
        req = requerimiento_ciclo(cfg, d)

        etapa = st.radio("📅 ¿En qué etapa está el cultivo?", ETAPAS, horizontal=True, key="etapa_sel",
                         format_func=lambda e: f"{ICONO_ETAPA[e]} {e.split(' - ')[1]} ({cfg['etapas'][e]['semanas']} sem)")
        info = cfg["etapas"][etapa]
        k_fxs = f"fxs_{etapa}"
        if k_fxs not in st.session_state:
            st.session_state[k_fxs] = int(info["ferti_semana"])
        fxs = st.number_input("💧 ¿Cuántas veces fertilizas por semana?", min_value=1, max_value=14, step=1, key=k_fxs)
        guardar_si_cambia(etapa, "ferti_semana", int(fxs))
        sem = max(int(info["semanas"]), 1)
        fxs = max(int(fxs), 1)

    sol = requerimiento_etapa(cfg, req, etapa)
    fac = d["area_ha"] / sem   # kg/ha-etapa -> kg por semana para las plantas
    req_sem = {"N_Total": (sol["N_NO3"] + sol["N_NH4"]) * fac, **{n: sol[n] * fac for n in REQ_NUT[1:]}}

    # ---------------- PASO 2 ----------------
    st.markdown("### 2️⃣ Elige tus fertilizantes")
    k_sel = f"sel_{etapa}"
    if k_sel not in st.session_state:
        previos = info.get("productos")
        if previos is None:
            previos = df_dosis[df_dosis["Etapa"] == etapa]["Producto"].tolist()
        st.session_state[k_sel] = [p for p in previos if p in set(catalogo["Producto"])]
    with st.container(border=True):
        elegidos = st.multiselect("Escribe el nombre para buscarlo y tócalo para agregarlo",
                                  options=sorted(catalogo["Producto"].tolist()), key=k_sel,
                                  placeholder="Ej: Calcinit, MKP, Sulfato de magnesio…")
        guardar_si_cambia(etapa, "productos", list(elegidos))
        k_mic = f"micros_{etapa}"
        if k_mic not in st.session_state:
            st.session_state[k_mic] = bool(info.get("micros", False))
        micros = st.toggle("Incluir también micronutrientes (hierro, zinc, boro…)", key=k_mic)
        guardar_si_cambia(etapa, "micros", bool(micros))

    nut_eval = MACROS + (MICROS if micros else [])

    if not elegidos:
        st.info("👆 Elige al menos un fertilizante para ver cuánto aplicar.")
        return

    # Cálculo automático (o manual si la persona lo pide)
    auto = dosis_auto_cache(tuple(elegidos), tuple(sorted(req_sem.items())), tuple(nut_eval), catalogo)
    k_man = f"manual_{etapa}"
    if k_man not in st.session_state:
        st.session_state[k_man] = bool(info.get("modo_manual", False))
    manual = st.toggle("✏️ Quiero poner yo mismo las cantidades", key=k_man,
                       help="Si lo dejas apagado, la app calcula las cantidades sola")
    guardar_si_cambia(etapa, "modo_manual", bool(manual))
    if manual:
        previas = info.get("dosis_manual", {})
        df_man = pd.DataFrame({"Producto": elegidos,
                               "kg_semana": [float(previas.get(p, auto.get(p, 0.0))) for p in elegidos]})
        st.caption(f"Escribe los **kg por semana** para tus {num(d['plantas'], 0)} plantas (litros si es líquido).")
        df_man = st.data_editor(df_man, hide_index=True, width="stretch", disabled=["Producto"],
                                key=f"ed_man_{etapa}_{abs(hash(tuple(elegidos)))}",
                                column_config={"Producto": st.column_config.TextColumn("Fertilizante"),
                                               "kg_semana": st.column_config.NumberColumn("kg por semana", min_value=0.0,
                                                                                          step=0.1, format="%.3f")})
        dosis_final = {p: float(v or 0) for p, v in zip(df_man["Producto"], df_man["kg_semana"])}
        guardar_si_cambia(etapa, "dosis_manual", dosis_final)
    else:
        dosis_final = {p: float(auto.get(p, 0.0)) for p in elegidos}

    # guardar dosis usadas (para el resumen del ciclo)
    actuales = df_dosis[df_dosis["Etapa"] == etapa]
    if dict(zip(actuales["Producto"], actuales["Dosis_lote_semana"].round(6))) != {p: round(v, 6) for p, v in dosis_final.items()}:
        df_dosis = pd.concat([df_dosis[df_dosis["Etapa"] != etapa],
                              pd.DataFrame({"Etapa": etapa, "Producto": list(dosis_final),
                                            "Dosis_lote_semana": list(dosis_final.values())})], ignore_index=True)
        guardar_dosis(df_dosis)

    df_uso = pd.DataFrame({"Producto": list(dosis_final), "Dosis_lote_semana": list(dosis_final.values())})
    res = calcular_dosis(df_uso, catalogo, d, sem, fxs)
    cat_i = catalogo.set_index("Producto")

    # ---------------- PASO 3 ----------------
    st.markdown("### 3️⃣ Cuánto aplicar")
    st.caption(f"Para tus **{num(d['plantas'], 0)} plantas** · {fxs} fertilizaciones por semana · etapa de {sem} semanas")
    if res.empty:
        st.warning("Con estos fertilizantes no se cubre ningún nutriente elegido. Prueba agregando otros.")
    for _, r in res.iterrows():
        liq = es_liquido(r["Producto"])
        with st.container(border=True):
            st.markdown(
                f"**{r['Producto']}**  \n"
                f"<span style='font-size:1.7rem;font-weight:700'>{cantidad(r['Por_aplicacion_kg'], liq)}</span>"
                f" <span style='color:gray'>en cada fertilización</span>  \n"
                f"<span style='color:gray'>{cantidad(r['Lote_semana_kg'], liq)} por semana · "
                f"{num(r['G_planta_semana'], 2)} {'mL' if liq else 'g'} por planta a la semana</span>",
                unsafe_allow_html=True)
    ceros = [p for p in elegidos if dosis_final.get(p, 0) <= 0]
    if ceros:
        st.caption("No hace falta aplicar: " + ", ".join(ceros) +
                   " (otros fertilizantes ya cubren lo que aportan).")

    # ¿Cubre lo que necesita la planta?
    st.markdown("#### ¿La planta recibe lo que necesita?")
    aporte = {n: 0.0 for n in nut_eval}
    for n in nut_eval:
        if n == "N_Total":
            aporte[n] = (res["N_NO3"].sum() + res["N_NH4"].sum()) * fac if not res.empty else 0.0
        else:
            aporte[n] = res[n].sum() * fac if not res.empty else 0.0
    pcts = {n: (aporte[n] / req_sem[n] * 100 if req_sem.get(n) else None) for n in nut_eval}
    with st.container(border=True):
        lineas = []
        for n in nut_eval:
            ico, txt = estado(pcts[n])
            lineas.append(f"{ico} **{NOMBRES[n]}** — {txt}" + (f" ({num(pcts[n], 0)}%)" if pcts[n] is not None else ""))
        st.markdown("  \n".join(lineas))

    faltan = [n for n in nut_eval if pcts[n] is not None and pcts[n] < 75]
    sobran = [n for n in nut_eval if pcts[n] is not None and pcts[n] > 125]
    for n in faltan:
        cands = catalogo[~catalogo["Producto"].isin(elegidos)].copy()
        cands["cont"] = cands.apply(lambda r: contenido(r, n), axis=1)
        if n != "CaO":   # sugerir primero productos sin calcio, para no complicar los tanques
            cands["cont"] = cands["cont"] - cands["CaO"] * 0.01
        top = cands[cands["cont"] > 0].sort_values("cont", ascending=False).head(3)["Producto"].tolist()
        with st.container(border=True):
            st.markdown(f"💡 **Le falta {NOMBRES[n].split(' (')[0].lower()}.** Agrega uno de estos:")
            if top:
                bcols = st.columns(len(top))
                for bc, p in zip(bcols, top):
                    bc.button(f"➕ {p}", key=f"add_{etapa}_{n}_{p}", on_click=agregar_producto, args=(k_sel, p),
                              width="stretch")
            else:
                st.caption("No hay en el catálogo un producto que lo aporte.")
    if sobran:
        st.info("🔺 Se pasa de " + ", ".join(NOMBRES[n].split(" (")[0].lower() for n in sobran) +
                ". Suele pasar cuando un fertilizante trae varios nutrientes juntos. Prueba quitando ese "
                "fertilizante o cambiándolo por otro.")

    with st.expander("🔢 Ver el detalle en números"):
        det = pd.DataFrame({"Nutriente": [NOMBRES[n] for n in nut_eval],
                            "Necesita (kg/semana)": [req_sem[n] for n in nut_eval],
                            "Recibe (kg/semana)": [aporte[n] for n in nut_eval],
                            "%": [pcts[n] for n in nut_eval]})
        st.dataframe(det.style.format({"Necesita (kg/semana)": "{:,.3f}", "Recibe (kg/semana)": "{:,.3f}",
                                       "%": "{:,.0f}%"}, na_rep="-"), hide_index=True, width="stretch")

    # ---------------- PASO 4 ----------------
    st.markdown("### 4️⃣ Cómo preparar los tanques")
    tq_txt = pd.DataFrame()
    if res.empty:
        st.info("Cuando haya cantidades a aplicar, aquí verás qué va en cada tanque.")
    else:
        with st.container(border=True):
            t1, t2, t3 = st.columns(3)
            k_base = f"tq_base_{etapa}"
            base_tq = t1.radio("Preparo los tanques para:", ["Cada fertilización", "Toda la semana"], key=k_base)
            k_va, k_vb = f"vol_a_{etapa}", f"vol_b_{etapa}"
            st.session_state.setdefault(k_va, float(info.get("vol_a", 200.0)))
            st.session_state.setdefault(k_vb, float(info.get("vol_b", 200.0)))
            vol_a = t2.number_input("Litros del tanque A", min_value=1.0, step=10.0, format="%.0f", key=k_va)
            vol_b = t3.number_input("Litros del tanque B", min_value=1.0, step=10.0, format="%.0f", key=k_vb)
            guardar_si_cambia(etapa, "vol_a", float(vol_a))
            guardar_si_cambia(etapa, "vol_b", float(vol_b))

        col_kg = "Por_aplicacion_kg" if base_tq == "Cada fertilización" else "Lote_semana_kg"
        tq = res[["Producto", col_kg]].rename(columns={col_kg: "kg"}).reset_index(drop=True)
        props = [proponer_tanque(cat_i.loc[p]) for p in tq["Producto"]]
        tq["Tanque"] = [t for t, _ in props]
        tq["Motivo"] = [m for _, m in props]
        carga = {"A": tq.loc[tq["Tanque"] == "A", "kg"].sum(), "B": tq.loc[tq["Tanque"] == "B", "kg"].sum()}
        for i_, r in tq[tq["Tanque"] == "A/B"].sort_values("kg", ascending=False).iterrows():
            destino = "A" if carga["A"] <= carga["B"] else "B"
            tq.at[i_, "Tanque"] = destino
            carga[destino] += r["kg"]
        cambios = dict(info.get("tanques", {}))
        tq["Tanque"] = [cambios.get(p_, t_) for p_, t_ in zip(tq["Producto"], tq["Tanque"])]

        ca_, cb_ = st.columns(2)
        for col_, t_, vol, color in [(ca_, "A", vol_a, "🟦"), (cb_, "B", vol_b, "🟧")]:
            sub = tq[tq["Tanque"] == t_]
            with col_.container(border=True):
                st.markdown(f"#### {color} Tanque {t_}")
                st.caption(f"{num(vol, 0)} litros de agua")
                if sub.empty:
                    st.write("— Vacío —")
                    continue
                for _, r in sub.iterrows():
                    st.markdown(f"- **{cantidad(r['kg'], es_liquido(r['Producto']))}** de {r['Producto']}")
                conc = sub["kg"].sum() * 1000 / vol
                if conc > 250:
                    st.warning(f"Muy concentrado ({num(conc, 0)} g por litro): puede no disolverse. "
                               "Usa más agua o prepáralo para cada fertilización.")

        comp_tq = {t_: tq.loc[tq["Tanque"] == t_, "Producto"].tolist() for t_ in "AB"}
        problema = False
        for t_ in "AB":
            con_ca = [p_ for p_ in comp_tq[t_] if cat_i.at[p_, "CaO"] >= 1]
            con_ps = [p_ for p_ in comp_tq[t_] if cat_i.at[p_, "P2O5"] >= 1 or cat_i.at[p_, "S"] >= 1]
            if [x for x in con_ca if x not in con_ps] and [x for x in con_ps if x not in con_ca]:
                problema = True
                st.error(f"❌ En el tanque {t_} hay calcio junto con fosfatos o sulfatos. Se forma una piedra "
                         "que tapa los goteros. Cámbialos de tanque abajo.")
            for x in [x for x in con_ca if x in con_ps]:
                st.warning(f"⚠️ {x} trae calcio y fosfato/sulfato juntos. Revisa la etiqueta o aplícalo aparte.")
        if not problema:
            st.success("✅ La mezcla está bien: el calcio (tanque A) no se junta con fosfatos ni sulfatos (tanque B).")

        with st.expander("🔄 Cambiar un fertilizante de tanque"):
            st.caption("La app ya los separó. Cámbialos solo si sabes lo que haces.")
            for _, r in tq.iterrows():
                cc1, cc2 = st.columns([3, 1])
                cc1.markdown(f"{r['Producto']}  \n<small>{r['Motivo']}</small>", unsafe_allow_html=True)
                nuevo = cc2.radio("Tanque", ["A", "B"], index=0 if r["Tanque"] == "A" else 1, horizontal=True,
                                  key=f"mv_{etapa}_{r['Producto']}", label_visibility="collapsed")
                if nuevo != r["Tanque"]:
                    cambios[r["Producto"]] = nuevo
                    guardar_si_cambia(etapa, "tanques", cambios)
                    st.rerun()
            if cambios and st.button("↩️ Volver a la propuesta de la app", key=f"reset_tq_{etapa}"):
                guardar_si_cambia(etapa, "tanques", {})
                for k in [k for k in st.session_state if k.startswith(f"mv_{etapa}_")]:
                    del st.session_state[k]
                st.rerun()
        st.caption("👉 Llena 2/3 del tanque con agua, disuelve un producto a la vez revolviendo y completa con agua.")
        tq_txt = tq

    # ---------------- PPM (opcional) ----------------
    with st.expander("💧 Ver ppm en el agua de riego (para comparar con el HORIBA)"):
        k_l = f"litros_{etapa}"
        st.session_state.setdefault(k_l, float(info.get("litros_ferti", 5000.0)))
        litros = st.number_input("¿Cuántos litros de agua lleva una fertilización en todo el lote?",
                                 min_value=1.0, step=100.0, format="%.0f", key=k_l)
        guardar_si_cambia(etapa, "litros_ferti", float(litros))
        if not res.empty:
            n_apl = sem * fxs
            kg_f = {n: res[n].sum() * d["area_ha"] / n_apl for n in NUTRIENTES}
            kg_f["N_Total"] = kg_f["N_NO3"] + kg_f["N_NH4"]
            rq_f = {"N_Total": (sol["N_NO3"] + sol["N_NH4"]) * d["area_ha"] / n_apl,
                    **{n: sol[n] * d["area_ha"] / n_apl for n in NUTRIENTES}}
            t_ppm = pd.DataFrame([{"Elemento": et, "ppm que pones": ppm_de(kg_f[n] * f, litros),
                                   "ppm que necesita": ppm_de(rq_f[n] * f, litros)} for et, n, f in PPM_FORMAS])
            st.dataframe(t_ppm.style.format({"ppm que pones": "{:,.1f}", "ppm que necesita": "{:,.1f}"}),
                         hide_index=True, width="stretch")
            st.caption("ppm = miligramos por litro. NO3⁻ (ion) es el valor que se compara con el HORIBA.")

    # ---------------- DESCARGA ----------------
    receta = res[["Producto", "Por_aplicacion_kg", "Lote_semana_kg", "G_planta_semana", "Cantidad_lote_kg"]].rename(columns={
        "Por_aplicacion_kg": "kg en cada fertilización", "Lote_semana_kg": "kg por semana",
        "G_planta_semana": "g por planta por semana", "Cantidad_lote_kg": "kg toda la etapa"})
    if not tq_txt.empty:
        receta = receta.merge(tq_txt[["Producto", "Tanque"]], on="Producto", how="left")
    estado_df = pd.DataFrame({"Nutriente": [NOMBRES[n] for n in nut_eval],
                              "Estado": [estado(pcts[n])[1] for n in nut_eval],
                              "% cubierto": [round(pcts[n] or 0) for n in nut_eval]})
    datos_df = pd.DataFrame({"Dato": ["Etapa", "Plantas", "Fertilizaciones por semana", "Semanas de la etapa", "Fecha"],
                             "Valor": [etapa, d["plantas"], fxs, sem, f"{datetime.now():%Y-%m-%d}"]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        datos_df.to_excel(w, index=False, sheet_name="Receta", startrow=0)
        receta.to_excel(w, index=False, sheet_name="Receta", startrow=len(datos_df) + 2)
        estado_df.to_excel(w, index=False, sheet_name="Nutrientes")
    st.download_button("📥 Descargar la receta (Excel)", data=buf.getvalue(), type="primary", width="stretch",
                       file_name=f"receta_{etapa.split(' - ')[1].lower()}_{datetime.now():%Y%m%d}.xlsx")


with tab_main:
    _principal()

d = diseno(cfg)
req = requerimiento_ciclo(cfg, d)

# ==========================================
# PESTAÑA RESUMEN Y COPIA
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

    st.markdown("---")
    st.subheader("💾 Copia de seguridad")
    st.caption("Si la app se reinicia en internet puede perder lo que guardaste. Descarga una copia de vez en "
               "cuando y súbela si ves que se borró.")
    copia = json.dumps({"config": cfg, "dosis": df_dosis.to_dict(orient="records")}, ensure_ascii=False, indent=1)
    st.download_button("📥 Descargar copia", data=copia.encode("utf-8"),
                       file_name=f"copia_fertilizacion_tomate_{datetime.now():%Y%m%d}.json", mime="application/json")
    subido = st.file_uploader("Subir una copia guardada", type="json")
    if subido is not None and st.button("♻️ Recuperar esta copia"):
        try:
            datos = json.load(subido)
            guardar_config(datos["config"])
            guardar_dosis(pd.DataFrame(datos["dosis"], columns=COLS_DOSIS))
            st.session_state.clear()
            st.success("✅ Copia recuperada")
            st.rerun()
        except Exception:
            st.error("Ese archivo no es una copia válida de esta app.")

# ==========================================
# PESTAÑA CONFIGURACIÓN AVANZADA
# ==========================================
with tab_adv:
    st.info("🔧 Esta parte es para el técnico. Si solo vas a fertilizar, no necesitas tocar nada aquí.")
    sub_dis, sub_etp, sub_cat = st.tabs(["Diseño y extracción", "Reparto por etapas", "Catálogo de fertilizantes"])
    with sub_dis:
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
            for k in ["plantas_main", "prod_main"]:
                st.session_state.pop(k, None)
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
    with sub_etp:
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
            for k in [k for k in st.session_state if k.startswith("fxs_")]:
                del st.session_state[k]
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
    with sub_cat:
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
