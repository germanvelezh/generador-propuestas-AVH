# -*- coding: utf-8 -*-
"""
Prototipo web (Streamlit) para generar propuestas comerciales.
Reutiliza el motor de calculo de motor_propuestas.py (sin reprogramar la logica).
Ejecutar:  streamlit run app.py
"""
import io, os, copy, json, tempfile
import streamlit as st
import pandas as pd
import motor_propuestas as mp

st.set_page_config(page_title="Generador de propuestas", page_icon="📄", layout="wide")

# ---------- estado ----------
if "cfg" not in st.session_state:
    st.session_state.cfg = copy.deepcopy(mp.DEFAULT_CONFIG)

def recs(df, required=None):
    """DataFrame -> lista de dicts con tipos nativos (sin numpy) y sin filas vacias."""
    rows = json.loads(df.to_json(orient="records"))
    if required:
        rows = [r for r in rows if all(r.get(k) not in (None, "") for k in required)]
    return rows

def cargos_nombres(cfg):
    return [c["nombre"] for c in cfg["catalogo"]["cargos"]]

def log_nombres(cfg):
    return [l["nombre"] for l in cfg["catalogo"]["logistica"]]

def generar_archivos(cfg):
    budget = mp.compute_budget(cfg)
    tmp = tempfile.mkdtemp()
    cod = cfg["proyecto"]["codigo"] or "PROPUESTA"
    xlsx = os.path.join(tmp, f"{cod}_Presupuesto.xlsx")
    docx = os.path.join(tmp, f"{cod}_Propuesta.docx")
    mp.build_excel(cfg, xlsx)
    mp.build_docx(cfg, budget, docx)
    with open(xlsx, "rb") as f: xb = f.read()
    with open(docx, "rb") as f: db = f.read()
    return budget, xb, db, os.path.basename(xlsx), os.path.basename(docx)

def money(v):
    return "$ {:,.0f}".format(round(v or 0))

# ---------- encabezado ----------
st.title("Generador de propuestas comerciales")
st.caption("Completa los datos, genera y descarga el Excel de cálculo y el documento de propuesta. "
           "El motor de cálculo es el mismo de la plantilla maestra.")

cfg = st.session_state.cfg

# ---------- cargar / guardar configuracion ----------
with st.sidebar:
    st.header("Proyecto")
    up = st.file_uploader("Cargar configuración (.json)", type=["json"])
    if up is not None:
        try:
            st.session_state.cfg = json.load(up)
            st.success("Configuración cargada. Revisa las pestañas.")
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo leer el JSON: {e}")
    st.download_button("Descargar configuración actual (.json)",
                       data=json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8"),
                       file_name="config_proyecto.json", mime="application/json")
    if st.button("Restablecer ejemplo (Flora)"):
        st.session_state.cfg = copy.deepcopy(mp.DEFAULT_CONFIG)
        st.rerun()

tabs = st.tabs(["1. Datos", "2. Actividades", "3. Entregables", "4. Desembolsos", "5. Catálogo", "▶ Generar"])

# ===== 1. Datos =====
with tabs[0]:
    c1, c2 = st.columns(2)
    pj = cfg["proyecto"]
    pj["codigo"] = c1.text_input("Código del proyecto", pj.get("codigo", ""))
    pj["cliente"] = c2.text_input("Cliente", pj.get("cliente", ""))
    pj["objeto"] = st.text_area("Objeto / Nombre de la oferta", pj.get("objeto", ""), height=70)
    c3, c4, c5 = st.columns(3)
    pj["fecha"] = c3.text_input("Fecha", str(pj.get("fecha", "")))
    pj["n_titulos"] = c4.number_input("N° de títulos mineros", min_value=1, value=int(pj.get("n_titulos", 5)), step=1)
    com = cfg["comercial"]
    cc1, cc2 = st.columns(2)
    com["utilidad"] = cc1.number_input("Utilidad (%)", min_value=0.0, max_value=1.0,
                                       value=float(com.get("utilidad", 0.25)), step=0.01, format="%.2f")
    com["iva"] = cc2.number_input("IVA (%)", min_value=0.0, max_value=1.0,
                                  value=float(com.get("iva", 0.19)), step=0.01, format="%.2f")

# ===== 2. Actividades =====
with tabs[1]:
    st.write("Activa las actividades, define los tiempos (días) y, si quieres, ajusta personal y logística.")
    for i, act in enumerate(cfg["actividades"]):
        with st.expander(act["nombre"], expanded=(i == 0)):
            a1, a2, a3 = st.columns([1, 1, 1])
            act["activa"] = a1.checkbox("Activa", value=bool(act.get("activa", False)), key=f"act{i}")
            act["tps_campo"] = a2.number_input("Días Fase de campo", min_value=0,
                                               value=int(act.get("tps_campo", 0)), step=1, key=f"tc{i}")
            act["tps_entregables"] = a3.number_input("Días Fase de entregables", min_value=0,
                                                     value=int(act.get("tps_entregables", 0)), step=1, key=f"te{i}")
            st.markdown("**Personal**")
            dfp = pd.DataFrame(act["personal"])
            dfp = st.data_editor(dfp, num_rows="dynamic", key=f"per{i}",
                                 column_config={"cargo": st.column_config.SelectboxColumn("cargo", options=cargos_nombres(cfg))})
            act["personal"] = recs(dfp, ["cargo"])
            st.markdown("**Logística e insumos**")
            dfl = pd.DataFrame(act["logistica"])
            dfl = st.data_editor(dfl, num_rows="dynamic", key=f"log{i}")
            act["logistica"] = recs(dfl, ["concepto"])

# ===== 3. Entregables =====
with tabs[2]:
    st.write("Productos de la fase de entregables. Tiempo en meses por producto.")
    ent = cfg["entregables"]
    ent["activa"] = st.checkbox("Incluir entregables", value=bool(ent.get("activa", True)))
    dfe = pd.DataFrame(ent["items"])
    dfe = st.data_editor(dfe, num_rows="dynamic", key="ent",
                         column_config={"cargo": st.column_config.SelectboxColumn("cargo", options=cargos_nombres(cfg))})
    ent["items"] = recs(dfe, ["cargo"])

# ===== 4. Desembolsos =====
with tabs[3]:
    st.write("Porcentajes de pago. Deben sumar 100%.")
    dfd = pd.DataFrame(cfg["desembolsos"])
    dfd = st.data_editor(dfd, num_rows="dynamic", key="des")
    cfg["desembolsos"] = recs(dfd, ["pct"])
    suma = sum(float(d.get("pct", 0) or 0) for d in cfg["desembolsos"])
    if abs(suma - 1.0) > 0.001:
        st.warning(f"Los desembolsos suman {suma*100:.0f}% (deberían sumar 100%).")
    else:
        st.success("Desembolsos = 100%.")

# ===== 5. Catálogo =====
with tabs[4]:
    st.write("Precios base (única fuente). Normalmente se ajusta pocas veces.")
    st.markdown("**Tabla salarial**")
    dfc = pd.DataFrame(cfg["catalogo"]["cargos"])
    dfc = st.data_editor(dfc, num_rows="dynamic", key="cat_cargos")
    cfg["catalogo"]["cargos"] = recs(dfc, ["nombre"])
    st.markdown("**Costos logísticos (valor unitario)**")
    dfcl = pd.DataFrame(cfg["catalogo"]["logistica"])
    dfcl = st.data_editor(dfcl, num_rows="dynamic", key="cat_log")
    cfg["catalogo"]["logistica"] = recs(dfcl, ["nombre"])

# ===== 6. Generar =====
with tabs[5]:
    st.write("Genera la propuesta con los datos actuales.")
    if st.button("Generar propuesta", type="primary"):
        try:
            budget, xb, db, xn, dn = generar_archivos(cfg)
            st.session_state["result"] = {"budget": budget, "xb": xb, "db": db, "xn": xn, "dn": dn}
        except Exception as e:
            st.error(f"Error al generar: {e}")

    if "result" in st.session_state:
        res = st.session_state["result"]; b = res["budget"]
        m1, m2, m3 = st.columns(3)
        m1.metric("Total antes de IVA", money(b["cu"]))
        m2.metric("IVA", money(b["iva"]))
        m3.metric("TOTAL GENERAL", money(b["total"]))
        activos = [x for x in b["bloques"] if x["activa"]]
        tabla = pd.DataFrame([{
            "Código": f"A{i+1}",
            "Ítem": x["nombre"],
            "Valor + utilidad": money(x["subtotal_u"]),
            "Valor por título": money(x["subtotal_u"] / b["n_titulos"]),
        } for i, x in enumerate(activos)])
        st.dataframe(tabla, hide_index=True)
        d1, d2 = st.columns(2)
        d1.download_button("⬇ Descargar Excel (cálculo)", data=res["xb"], file_name=res["xn"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("⬇ Descargar propuesta (Word)", data=res["db"], file_name=res["dn"],
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
