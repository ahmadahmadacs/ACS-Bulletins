"""
ACS — Générateur de Bulletins
Application Streamlit — Interface utilisateur
"""

import os
import shutil
import tempfile
import subprocess
from pathlib import Path

# ─────────────────────────────────────────────
# FONT SETUP — refresh cache pour LibreOffice
# ─────────────────────────────────────────────
def setup_fonts():
    """Force le refresh du cache polices pour que LibreOffice utilise
    les polices Microsoft installées via packages.txt (ttf-mscorefonts-installer)."""
    try:
        subprocess.run(["fc-cache", "-f", "-v"],
                       capture_output=True, timeout=30)
    except Exception:
        pass

setup_fonts()

import streamlit as st
from generator import (
    ensure_libreoffice, ensure_calibri,
    generate_all, zip_pdfs, zip_xlsx, GROUPS
)

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ACS — Bulletins",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
  /* Header */
  .acs-header {
    background: linear-gradient(135deg, #1A3A6B 0%, #1A73C8 100%);
    padding: 24px 32px;
    border-radius: 16px;
    margin-bottom: 24px;
    display: flex;
    align-items: center;
    gap: 20px;
  }
  .acs-header h1 {
    color: white !important;
    margin: 0;
    font-size: 2rem;
  }
  .acs-header p {
    color: rgba(255,255,255,0.82);
    margin: 4px 0 0;
    font-size: 1rem;
  }

  /* Step cards */
  .step-card {
    background: #F8FAFF;
    border: 1.5px solid #D6E4FF;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 10px;
  }
  .step-title {
    font-weight: 700;
    color: #1A3A6B;
    font-size: 1.05rem;
    margin-bottom: 6px;
  }

  /* Status badges */
  .badge-ok   { background:#D4EDDA; color:#155724; padding:3px 10px; border-radius:20px; font-size:0.85rem; font-weight:600; }
  .badge-warn { background:#FFF3CD; color:#856404; padding:3px 10px; border-radius:20px; font-size:0.85rem; font-weight:600; }
  .badge-err  { background:#F8D7DA; color:#721C24; padding:3px 10px; border-radius:20px; font-size:0.85rem; font-weight:600; }

  /* Log box */
  .log-box {
    background: #0D1117;
    color: #C9D1D9;
    font-family: monospace;
    font-size: 0.82rem;
    padding: 16px;
    border-radius: 10px;
    max-height: 320px;
    overflow-y: auto;
    line-height: 1.7;
  }

  /* Download button */
  .stDownloadButton > button {
    background: linear-gradient(135deg, #1A3A6B, #1A73C8) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 12px 28px !important;
    font-size: 1.05rem !important;
    font-weight: 700 !important;
    width: 100%;
  }

  /* Sidebar */
  section[data-testid="stSidebar"] {
    background: #F0F4FF;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="acs-header">
  <div>
    <h1>🎓 ACS — Générateur de Bulletins</h1>
    <p>Akroum College of Sciences &nbsp;·&nbsp; مدرسة أكروم للعلوم &nbsp;·&nbsp; Année scolaire 2025-2026</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR — Guide
# ─────────────────────────────────────────────
with st.sidebar:
    st.image("https://via.placeholder.com/200x80/1A3A6B/FFFFFF?text=ACS",
             use_column_width=True)
    st.markdown("---")
    st.markdown("### 📖 Guide rapide")
    st.markdown("""
**Étape 1** — Uploadez le fichier Excel des notes  
*(doit contenir les feuilles `S1/S2/S3 Notes EB...`)*

**Étape 2** — Uploadez les templates Excel  
*(un par groupe : EB1-2, EB3-6, EB7, EB8, EB9)*

**Étape 3** — Sélectionnez les groupes à générer

**Étape 4** — Cliquez **Générer** et téléchargez le ZIP
""")
    st.markdown("---")
    st.markdown("### ℹ️ Groupes disponibles")
    for g, cfg in GROUPS.items():
        st.markdown(f"- **{g}** — {cfg['nb_cols']} colonnes de notes")
    st.markdown("---")
    st.caption("v7.0 · ACS 2025-2026")

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "logs"       not in st.session_state: st.session_state.logs = []
if "pdf_zip"    not in st.session_state: st.session_state.pdf_zip = None
if "xlsx_zip"   not in st.session_state: st.session_state.xlsx_zip = None
if "pdf_count"  not in st.session_state: st.session_state.pdf_count = 0
if "errors"     not in st.session_state: st.session_state.errors = []
if "sys_ready"  not in st.session_state: st.session_state.sys_ready = False
if "semestre"      not in st.session_state: st.session_state.semestre = "S3"
if "pdf_zip_name"  not in st.session_state: st.session_state.pdf_zip_name = "bulletins_ACS_S3.zip"
if "xlsx_zip_name" not in st.session_state: st.session_state.xlsx_zip_name = "bulletins_ACS_S3_excel.zip"

# ─────────────────────────────────────────────
# STEP 1 — Système (LibreOffice + Calibri)
# ─────────────────────────────────────────────
col_sys, col_info = st.columns([3, 1])
with col_sys:
    st.markdown('<div class="step-title">⚙️ Étape 0 — Vérification du système</div>',
                unsafe_allow_html=True)
with col_info:
    if st.session_state.sys_ready:
        st.markdown('<span class="badge-ok">✅ Prêt</span>', unsafe_allow_html=True)

if not st.session_state.sys_ready:
    if st.button("🔧 Vérifier LibreOffice + Calibri", use_container_width=True):
        logs = []
        with st.spinner("Vérification du système..."):
            ensure_libreoffice(log=logs.append)
            ensure_calibri(log=logs.append)
        st.session_state.sys_ready = True
        st.session_state.logs = logs
        st.rerun()
else:
    st.success("✅ LibreOffice et Calibri sont prêts.")

st.markdown("---")

# ─────────────────────────────────────────────
# STEP 2 — Upload fichier notes
# ─────────────────────────────────────────────
st.markdown('<div class="step-title">📊 Étape 1 — Fichier des notes (.xlsx)</div>',
            unsafe_allow_html=True)

notes_file = st.file_uploader(
    "Uploadez le fichier Excel contenant toutes les notes",
    type=["xlsx"],
    key="notes_upload",
    help="Le fichier doit contenir les feuilles : S1 Notes EB1-2, S2 Notes EB3-6, etc."
)

if notes_file:
    st.success(f"✅ **{notes_file.name}** chargé ({notes_file.size // 1024} Ko)")

st.markdown("---")

# ─────────────────────────────────────────────
# STEP 3 — Upload templates
# ─────────────────────────────────────────────
st.markdown('<div class="step-title">📄 Étape 2 — Templates Excel (un par groupe)</div>',
            unsafe_allow_html=True)

template_names = {
    "EB1-2": "template-eb-1to2.xlsx",
    "EB3-6": "template-eb-3to6.xlsx",
    "EB7":   "template-eb7.xlsx",
    "EB8":   "template-eb8.xlsx",
    "EB9":   "template-eb9.xlsx",
}

cols = st.columns(5)
template_files = {}
for i, (group, fname) in enumerate(template_names.items()):
    with cols[i]:
        f = st.file_uploader(
            f"**{group}**",
            type=["xlsx"],
            key=f"tmpl_{group}",
            label_visibility="visible",
        )
        template_files[group] = f
        if f:
            st.caption(f"✅ {f.name}")
        else:
            st.caption(f"⬆️ {fname}")

st.markdown("---")

# ─────────────────────────────────────────────
# STEP 4 — Sélection des groupes
# ─────────────────────────────────────────────
st.markdown('<div class="step-title">🎯 Étape 3 — Groupes à générer</div>',
            unsafe_allow_html=True)

available_groups = [g for g, f in template_files.items() if f is not None]
if notes_file and available_groups:
    selected_groups = st.multiselect(
        "Sélectionnez les groupes",
        options=available_groups,
        default=available_groups,
        format_func=lambda x: f"🏫 {x}",
    )
else:
    selected_groups = []
    if not notes_file:
        st.info("👆 Uploadez d'abord le fichier notes pour voir les groupes disponibles.")
    elif not available_groups:
        st.info("👆 Uploadez au moins un template pour commencer.")

st.markdown("---")

# ─────────────────────────────────────────────
# STEP 4b — Semestre cible
# ─────────────────────────────────────────────
st.markdown('<div class="step-title">📅 Étape 3b — Semestre à générer</div>',
            unsafe_allow_html=True)

col_s1, col_s2, col_s3 = st.columns(3)
sem_choice = st.session_state.semestre
with col_s1:
    if st.button("📘 S1 — Semestre 1 seulement", use_container_width=True,
                 type="primary" if sem_choice == "S1" else "secondary"):
        st.session_state.semestre = "S1"
        st.rerun()
    st.caption("Notes S1 uniquement · Pas de moyenne")
with col_s2:
    if st.button("📗 S2 — Semestres 1 + 2", use_container_width=True,
                 type="primary" if sem_choice == "S2" else "secondary"):
        st.session_state.semestre = "S2"
        st.rerun()
    st.caption("Notes S1 + S2 · Moyenne S1+S2")
with col_s3:
    if st.button("📕 S3 — Semestres 1 + 2 + 3", use_container_width=True,
                 type="primary" if sem_choice == "S3" else "secondary"):
        st.session_state.semestre = "S3"
        st.rerun()
    st.caption("Notes S1 + S2 + S3 · Moyenne finale")

st.info(f"✅ Semestre sélectionné : **{st.session_state.semestre}**")

st.markdown("---")

# ─────────────────────────────────────────────
# STEP 5 — GENERATE
# ─────────────────────────────────────────────
st.markdown('<div class="step-title">🚀 Étape 4 — Génération</div>',
            unsafe_allow_html=True)

ready = (
    st.session_state.sys_ready
    and notes_file is not None
    and len(selected_groups) > 0
)

if not ready:
    missing = []
    if not st.session_state.sys_ready: missing.append("système non initialisé")
    if not notes_file:                  missing.append("fichier notes manquant")
    if not selected_groups:             missing.append("aucun groupe sélectionné")
    st.warning("⚠️ " + " · ".join(missing))

generate_btn = st.button(
    "▶️  Générer les bulletins PDF",
    disabled=not ready,
    use_container_width=True,
    type="primary",
)

if generate_btn and ready:
    st.session_state.logs   = []
    st.session_state.pdf_zip = None
    st.session_state.errors  = []

    log_placeholder  = st.empty()
    prog_placeholder = st.empty()
    status_text      = st.empty()

    logs = []

    def log(msg):
        logs.append(msg)
        log_html = "\n".join(logs[-40:])  # dernières 40 lignes
        log_placeholder.markdown(
            f'<div class="log-box">{log_html}</div>',
            unsafe_allow_html=True)

    def progress_cb(current, total, name):
        pct = current / total
        prog_placeholder.progress(pct, text=f"[{current}/{total}] {name}")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Sauvegarder les fichiers uploadés
        notes_path = tmpdir / notes_file.name
        notes_path.write_bytes(notes_file.read())

        templates = {}
        for group in selected_groups:
            f = template_files[group]
            tmpl_path = tmpdir / f"template_{group}.xlsx"
            tmpl_path.write_bytes(f.read())
            templates[group] = tmpl_path

        output_dir = tmpdir / "output"

        log(f"🚀 Démarrage — {len(selected_groups)} groupe(s) · Semestre cible : {st.session_state.semestre}")

        pdf_list, xlsx_list, error_list = generate_all(
            notes_path=notes_path,
            templates=templates,
            output_dir=output_dir,
            log=log,
            progress_cb=progress_cb,
            semestre_cible=st.session_state.semestre,
        )

        sem = st.session_state.semestre
        if pdf_list:
            zip_path = tmpdir / f"bulletins_ACS_{sem}.zip"
            zip_pdfs(pdf_list, zip_path)
            st.session_state.pdf_zip      = zip_path.read_bytes()
            st.session_state.pdf_count    = len(pdf_list)
            st.session_state.pdf_zip_name = f"bulletins_ACS_{sem}.zip"
            log(f"\n🎉 {len(pdf_list)} bulletins générés avec succès !")
        else:
            log("❌ Aucun bulletin généré.")

        if xlsx_list:
            xlsx_zip_path = tmpdir / f"bulletins_ACS_{sem}_excel.zip"
            zip_xlsx(xlsx_list, xlsx_zip_path)
            st.session_state.xlsx_zip      = xlsx_zip_path.read_bytes()
            st.session_state.xlsx_zip_name = f"bulletins_ACS_{sem}_excel.zip"

        st.session_state.errors = error_list
        st.session_state.logs   = logs

    prog_placeholder.empty()
    status_text.empty()
    st.rerun()

# ─────────────────────────────────────────────
# RÉSULTATS
# ─────────────────────────────────────────────
if st.session_state.logs:
    st.markdown("### 📋 Journal de génération")
    log_html = "<br>".join(st.session_state.logs)
    st.markdown(f'<div class="log-box">{log_html}</div>', unsafe_allow_html=True)

if st.session_state.pdf_zip:
    st.markdown("---")
    col_dl, col_xlsx, col_stat = st.columns([2, 2, 1])
    with col_dl:
        st.download_button(
            label=f"⬇️  Télécharger les {st.session_state.pdf_count} bulletins PDF (ZIP)",
            data=st.session_state.pdf_zip,
            file_name=st.session_state.pdf_zip_name,
            mime="application/zip",
            use_container_width=True,
        )
    with col_xlsx:
        if st.session_state.xlsx_zip:
            st.download_button(
                label=f"📊  Télécharger les {st.session_state.pdf_count} bulletins Excel (ZIP)",
                data=st.session_state.xlsx_zip,
                file_name=st.session_state.xlsx_zip_name,
                mime="application/zip",
                use_container_width=True,
            )
    with col_stat:
        size_kb = len(st.session_state.pdf_zip) // 1024
        st.metric("Bulletins générés", st.session_state.pdf_count)
        st.caption(f"Taille du ZIP : {size_kb} Ko")

if st.session_state.errors:
    with st.expander(f"⚠️ {len(st.session_state.errors)} erreur(s) à vérifier", expanded=False):
        for err in st.session_state.errors:
            st.markdown(f"- `{err}`")
