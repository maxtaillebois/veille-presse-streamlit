#!/usr/bin/env python3
"""
Application de Veille Presse v4.0 — Cloud Edition
===================================================
Deployable sur Streamlit Community Cloud.
Recupere les donnees (JSON + PDF) depuis Google Drive.
Le workflow N8N depose chaque semaine :
  - veille_data.json (analyse IA des articles)
  - Les PDF originaux des articles

Lancer en local : streamlit run veille_presse_app.py
"""

import streamlit as st
from pypdf import PdfReader, PdfWriter
from docx import Document as DocxDocument
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import re
import io
import json
import zipfile
import requests
from datetime import datetime


# ==================== CONFIGURATION ====================

# ID du dossier Google Drive partage (contient les PDF + veille_data.json)
GDRIVE_FOLDER_ID = "15MwTpntChfAgDQYiDSvc5cJ0KYmuW9_p"

# URL du webhook N8N pour l'envoi du mail a Stephanie
N8N_WEBHOOK_URL = "https://maximetaillebois.app.n8n.cloud/webhook/veille-presse-envoi"

# Mots-cles de reference (pour le badge couleur)
KEYWORDS = ["Yannick Borde", "Procivis", "Immo de France", "Maisons d'en France"]


def get_google_api_key():
    """Recupere la cle API Google depuis les secrets Streamlit ou l'environnement."""
    try:
        return st.secrets["GOOGLE_API_KEY"]
    except Exception:
        return os.environ.get("GOOGLE_API_KEY", "")


# ==================== FONCTIONS GOOGLE DRIVE ====================

@st.cache_data(ttl=300)  # Cache 5 minutes
def list_drive_files(folder_id: str, api_key: str) -> list:
    """Liste les fichiers dans un dossier Google Drive partage."""
    url = "https://www.googleapis.com/drive/v3/files"
    params = {
        "q": f"'{folder_id}' in parents and trashed = false",
        "fields": "files(id,name,mimeType,modifiedTime,size)",
        "orderBy": "modifiedTime desc",
        "pageSize": 100,
        "key": api_key
    }
    try:
        resp = requests.get(url, params=params, timeout=15)
        resp.raise_for_status()
        return resp.json().get("files", [])
    except Exception as e:
        st.error(f"Erreur lors de la lecture du Google Drive : {e}")
        return []


@st.cache_data(ttl=300)
def download_drive_file(file_id: str, api_key: str) -> bytes:
    """Telecharge un fichier depuis Google Drive."""
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    params = {"alt": "media", "key": api_key}
    try:
        resp = requests.get(url, params=params, timeout=60)
        resp.raise_for_status()
        return resp.content
    except Exception as e:
        st.error(f"Erreur de telechargement : {e}")
        return b""


def find_file_in_drive(files: list, filename: str) -> dict:
    """Trouve un fichier par son nom dans la liste Drive."""
    for f in files:
        if f["name"] == filename:
            return f
    return {}


def load_veille_data_from_drive(files: list, api_key: str) -> dict:
    """Charge le veille_data.json depuis Google Drive."""
    json_file = find_file_in_drive(files, "veille_data.json")
    if not json_file:
        return None
    content = download_drive_file(json_file["id"], api_key)
    if not content:
        return None
    try:
        return json.loads(content.decode("utf-8"))
    except json.JSONDecodeError as e:
        st.error(f"Erreur de lecture du fichier JSON : {e}")
        return None


# ==================== CONFIGURATION PAGE ====================

st.set_page_config(
    page_title="Veille Presse — Procivis",
    page_icon="📰",
    layout="wide"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A5F;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .keyword-highlight {
        background-color: #FFEB3B;
        padding: 2px 4px;
        border-radius: 3px;
        font-weight: bold;
    }
    .media-badge {
        background-color: #1976D2;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.85rem;
        font-weight: bold;
        display: inline-block;
    }
    .date-badge {
        background-color: #43A047;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.85rem;
        display: inline-block;
    }
    .keyword-badge {
        background-color: #FF9800;
        color: white;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.85rem;
        display: inline-block;
    }
    .context-text {
        font-style: italic;
        color: #424242;
        background-color: #F5F5F5;
        padding: 10px;
        border-radius: 5px;
        margin: 8px 0;
        border-left: 3px solid #1976D2;
    }
    .summary-text {
        color: #333;
        background-color: #E8F5E9;
        padding: 10px;
        border-radius: 5px;
        margin: 8px 0;
        border-left: 3px solid #43A047;
    }
    .stButton > button { width: 100%; }
</style>
""", unsafe_allow_html=True)


# ==================== FONCTIONS UTILITAIRES ====================

def highlight_keywords_html(text: str, keywords: list) -> str:
    """Surligne les mots-cles dans le texte (HTML)."""
    result = text
    for keyword in keywords:
        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
        result = pattern.sub(
            f'<span class="keyword-highlight">{keyword}</span>',
            result
        )
    return result


def parse_date_for_sorting(date_str: str) -> datetime:
    """Parse une date JJ/MM/AAAA pour le tri."""
    if not date_str:
        return datetime.min
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"]:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return datetime.min


# ==================== FONCTIONS DE COMPILATION ====================

def create_compilation_pdf_bytes(pdf_contents: list) -> bytes:
    """Compile plusieurs PDF (en bytes) en un seul fichier."""
    writer = PdfWriter()
    for pdf_bytes, filename in pdf_contents:
        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            st.warning(f"Impossible d'ajouter {filename} : {e}")

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)
    return output.getvalue()


def create_recap_docx_bytes(articles: list) -> bytes:
    """Cree le Word recapitulatif en memoire."""
    sorted_articles = sorted(
        articles,
        key=lambda x: parse_date_for_sorting(x.get('date', '')),
        reverse=True
    )

    doc = DocxDocument()

    title = doc.add_heading('RECAPITULATIF VEILLE PRESSE', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    info_para = doc.add_paragraph()
    info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info_para.add_run(
        f"Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}"
    )
    info_run.font.size = Pt(10)
    info_run.font.italic = True

    count_para = doc.add_paragraph()
    count_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    count_run = count_para.add_run(f"{len(articles)} article(s)")
    count_run.font.size = Pt(10)

    doc.add_paragraph("-" * 50)

    for article in sorted_articles:
        media = article.get('media', 'Media inconnu')
        title_text = article.get('title', 'Titre inconnu')
        date = article.get('date', 'Date inconnue')

        para = doc.add_paragraph()
        media_run = para.add_run(media)
        media_run.bold = True
        para.add_run(" | ")
        title_run = para.add_run(f'"{title_text}"')
        title_run.italic = True
        para.add_run(" | ")
        para.add_run(date)

    doc.add_paragraph("-" * 50)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


# ==================== FONCTION WEBHOOK ====================

def send_to_stephanie(pdf_bytes: bytes, docx_bytes: bytes, articles: list,
                      week_start: str, week_end: str) -> bool:
    """Declenche le webhook N8N pour envoyer le mail a Stephanie."""
    if not N8N_WEBHOOK_URL:
        st.error("L'URL du webhook N8N n'est pas configuree.")
        return False

    try:
        date_suffix = datetime.now().strftime('%Y%m%d')
        files = {}
        if pdf_bytes:
            files['pdf'] = (
                f"compilation_veille_{date_suffix}.pdf",
                io.BytesIO(pdf_bytes),
                'application/pdf'
            )
        if docx_bytes:
            files['docx'] = (
                f"compilation_veille_{date_suffix}_recap.docx",
                io.BytesIO(docx_bytes),
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

        data = {
            'week_start': week_start,
            'week_end': week_end,
            'article_count': str(len(articles)),
            'generated_at': datetime.now().isoformat(),
            'selected_articles': json.dumps([
                {"media": a.get("media", ""), "title": a.get("title", ""), "date": a.get("date", "")}
                for a in articles
            ])
        }

        response = requests.post(N8N_WEBHOOK_URL, data=data, files=files, timeout=30)
        return response.status_code == 200

    except Exception as e:
        st.error(f"Erreur lors de l'envoi : {e}")
        return False


# ==================== INTERFACE PRINCIPALE ====================

def main():
    st.markdown(
        '<h1 class="main-header">📰 Veille Presse</h1>',
        unsafe_allow_html=True
    )

    # Verifier la cle API
    api_key = get_google_api_key()
    if not api_key:
        st.error(
            "**Cle API Google manquante.**\n\n"
            "Configurez `GOOGLE_API_KEY` dans les secrets Streamlit "
            "(Settings > Secrets) ou en variable d'environnement."
        )
        st.code('GOOGLE_API_KEY = "votre-cle-api-google"', language="toml")
        return

    # ===== SIDEBAR =====
    with st.sidebar:
        st.header("Configuration")

        st.caption(f"Dossier Google Drive : `{GDRIVE_FOLDER_ID}`")

        if st.button("Rafraichir les donnees", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.divider()

        st.subheader("Envoi mail")
        webhook_url = st.text_input(
            "URL webhook N8N",
            value=N8N_WEBHOOK_URL,
            help="URL du webhook N8N pour l'envoi automatique",
            type="password"
        )

    # ===== CHARGEMENT DES DONNEES =====

    if 'selected_ids' not in st.session_state:
        st.session_state.selected_ids = set()
    if 'compilation_done' not in st.session_state:
        st.session_state.compilation_done = False
    if 'compilation_data' not in st.session_state:
        st.session_state.compilation_data = {}

    # Lister les fichiers du Drive
    with st.spinner("Chargement depuis Google Drive..."):
        drive_files = list_drive_files(GDRIVE_FOLDER_ID, api_key)

    if not drive_files:
        st.info(
            "Aucun fichier trouve dans le dossier Google Drive.\n\n"
            "Le workflow N8N s'execute chaque vendredi a 9h00. "
            "Une fois les fichiers generes, rechargez cette page."
        )
        return

    # Charger le JSON
    data = load_veille_data_from_drive(drive_files, api_key)

    if data is None:
        st.info(
            "En attente du fichier `veille_data.json` dans Google Drive.\n\n"
            "Le workflow N8N s'execute chaque vendredi a 9h00."
        )
        with st.expander("Fichiers presents dans le Drive"):
            for f in drive_files:
                st.text(f"{f['name']} ({f.get('size', '?')} octets)")
        return

    articles = data.get('articles', [])
    week_start = data.get('week_start', '')
    week_end = data.get('week_end', '')
    generated_at = data.get('generated_at', '')

    if not articles:
        st.warning("Le fichier JSON ne contient aucun article.")
        return

    # Indexer les PDF du Drive par nom de fichier
    pdf_drive_index = {}
    for f in drive_files:
        if f["name"].lower().endswith(".pdf"):
            pdf_drive_index[f["name"]] = f["id"]

    # ===== EN-TETE =====
    if week_start and week_end:
        st.markdown(
            f'<p class="sub-header">Semaine du {week_start} au {week_end} '
            f'| {len(articles)} article(s) analyse(s) '
            f'| Donnees du {generated_at[:10] if generated_at else "?"}</p>',
            unsafe_allow_html=True
        )

    # ===== STATISTIQUES =====
    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
    with col_stat1:
        st.metric("Articles", len(articles))
    with col_stat2:
        kw_counts = {}
        for a in articles:
            for kw in a.get('keywords_found', []):
                kw_counts[kw] = kw_counts.get(kw, 0) + 1
        top_kw = max(kw_counts, key=kw_counts.get) if kw_counts else "—"
        st.metric("Mot-cle principal", top_kw)
    with col_stat3:
        medias = set(a.get('media', '') for a in articles)
        st.metric("Medias distincts", len(medias))
    with col_stat4:
        selected_count = len(st.session_state.selected_ids)
        st.metric("Selectionnes", f"{selected_count}/{len(articles)}")

    st.divider()

    # ===== BOUTONS DE SELECTION RAPIDE =====
    col_sel1, col_sel2, col_sel3 = st.columns([1, 1, 3])
    with col_sel1:
        if st.button("Tout selectionner", use_container_width=True):
            st.session_state.selected_ids = set(
                a.get('id', f'article_{i}') for i, a in enumerate(articles)
            )
            st.rerun()
    with col_sel2:
        if st.button("Tout deselectionner", use_container_width=True):
            st.session_state.selected_ids = set()
            st.rerun()

    st.divider()

    # ===== AFFICHAGE DES ARTICLES =====
    st.header("Articles de la semaine")

    for i, article in enumerate(articles):
        article_id = article.get('id', f'article_{i}')
        media = article.get('media', 'Media inconnu')
        title = article.get('title', 'Titre inconnu')
        date = article.get('date', 'Date inconnue')
        keywords_found = article.get('keywords_found', [])
        context = article.get('context', '')
        summary = article.get('summary', '')
        pdf_filename = article.get('pdf_filename', article.get('fileName', ''))

        with st.container():
            col_main, col_select = st.columns([5, 1])

            with col_main:
                st.markdown(f"### {title}")

                badges_html = f'<span class="media-badge">{media}</span> '
                badges_html += f'<span class="date-badge">{date}</span> '
                for kw in keywords_found:
                    badges_html += f'<span class="keyword-badge">{kw}</span> '
                st.markdown(badges_html, unsafe_allow_html=True)

                if pdf_filename:
                    has_pdf = pdf_filename in pdf_drive_index
                    icon = "📎" if has_pdf else "⚠️"
                    st.caption(f"{icon} {pdf_filename}")

                if summary:
                    st.markdown(
                        f'<div class="summary-text"><b>Resume :</b> {summary}</div>',
                        unsafe_allow_html=True
                    )

                if context:
                    highlighted = highlight_keywords_html(context, keywords_found)
                    with st.expander("Voir le contexte de citation", expanded=False):
                        st.markdown(
                            f'<div class="context-text">{highlighted}</div>',
                            unsafe_allow_html=True
                        )

            with col_select:
                is_selected = st.checkbox(
                    "Garder",
                    key=f"select_{article_id}",
                    value=article_id in st.session_state.selected_ids
                )
                if is_selected:
                    st.session_state.selected_ids.add(article_id)
                else:
                    st.session_state.selected_ids.discard(article_id)

            st.divider()

    # ===== SECTION DE COMPILATION =====
    st.header("Compilation & Envoi")

    selected_count = len(st.session_state.selected_ids)
    st.info(f"**{selected_count}** article(s) selectionne(s)")

    if selected_count > 0:
        selected_articles = []
        for i, a in enumerate(articles):
            aid = a.get('id', f'article_{i}')
            if aid in st.session_state.selected_ids:
                selected_articles.append(a)

        date_suffix = datetime.now().strftime('%Y%m%d')

        col_compile, col_send = st.columns(2)

        with col_compile:
            if st.button("Generer la compilation", type="primary", use_container_width=True):
                with st.spinner("Generation en cours (telechargement des PDF depuis Drive)..."):
                    # 1. Generer le Word recapitulatif
                    docx_bytes = create_recap_docx_bytes(selected_articles)

                    # 2. Telecharger et compiler les PDF
                    pdf_contents = []
                    missing_pdfs = []
                    for a in selected_articles:
                        pdf_filename = a.get('pdf_filename', a.get('fileName', ''))
                        if pdf_filename and pdf_filename in pdf_drive_index:
                            file_id = pdf_drive_index[pdf_filename]
                            pdf_bytes = download_drive_file(file_id, api_key)
                            if pdf_bytes:
                                pdf_contents.append((pdf_bytes, pdf_filename))
                            else:
                                missing_pdfs.append(pdf_filename)
                        else:
                            missing_pdfs.append(pdf_filename or a.get('title', '?'))

                    compiled_pdf_bytes = b""
                    if pdf_contents:
                        compiled_pdf_bytes = create_compilation_pdf_bytes(pdf_contents)

                    if missing_pdfs:
                        st.warning(
                            f"{len(missing_pdfs)} PDF non trouve(s) : "
                            f"{', '.join(missing_pdfs)}"
                        )

                    st.session_state.compilation_done = True
                    st.session_state.compilation_data = {
                        'pdf_bytes': compiled_pdf_bytes,
                        'docx_bytes': docx_bytes,
                    }

                    st.success("Recapitulatif Word genere")
                    if compiled_pdf_bytes:
                        st.success(f"Compilation PDF generee ({len(pdf_contents)} articles)")
                    st.balloons()

        # Boutons de telechargement
        if st.session_state.compilation_done:
            comp_data = st.session_state.compilation_data

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                if comp_data.get('docx_bytes'):
                    zf.writestr(f"compilation_veille_{date_suffix}_recap.docx", comp_data['docx_bytes'])
                if comp_data.get('pdf_bytes'):
                    zf.writestr(f"compilation_veille_{date_suffix}.pdf", comp_data['pdf_bytes'])
            zip_buffer.seek(0)

            st.download_button(
                label="Telecharger tout (PDF + Word)",
                data=zip_buffer.getvalue(),
                file_name=f"veille_presse_{date_suffix}.zip",
                mime="application/zip",
                use_container_width=True
            )

        # Bouton d'envoi a Stephanie
        with col_send:
            if st.session_state.compilation_done:
                if st.button("Envoyer a Stephanie", type="secondary", use_container_width=True):
                    comp_data = st.session_state.compilation_data
                    effective_webhook = webhook_url or N8N_WEBHOOK_URL

                    if not effective_webhook:
                        st.error("L'URL du webhook N8N n'est pas configuree.")
                    else:
                        with st.spinner("Envoi en cours..."):
                            success = send_to_stephanie(
                                pdf_bytes=comp_data.get('pdf_bytes', b''),
                                docx_bytes=comp_data.get('docx_bytes', b''),
                                articles=selected_articles,
                                week_start=week_start,
                                week_end=week_end
                            )
                            if success:
                                st.success("Mail envoye a Stephanie !")
                            else:
                                st.error("L'envoi a echoue. Verifiez le webhook N8N.")
            else:
                st.button(
                    "Envoyer a Stephanie",
                    type="secondary",
                    use_container_width=True,
                    disabled=True,
                    help="Generez d'abord la compilation"
                )
    else:
        st.warning("Selectionnez au moins un article pour creer une compilation.")


if __name__ == "__main__":
    main()
