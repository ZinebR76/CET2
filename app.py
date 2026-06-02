import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import requests
from io import BytesIO
import time

# Configuration de la page
st.set_page_config(
    page_title="Carte Interactive - Fournisseurs CET",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé
st.markdown("""
    <style>
        .main {
            padding-top: 2rem;
        }
        .stMetric {
            background-color: #f0f2f6;
            padding: 1rem;
            border-radius: 0.5rem;
            border-left: 4px solid #1e3c72;
        }
    </style>
""", unsafe_allow_html=True)

# Titre et description
st.markdown("# 🗺️ Carte Interactive - Fournisseurs CET")
st.markdown("**Bouygues Construction** - Sourcing CVC Plomberie")

# URL du fichier Excel sur SharePoint
EXCEL_URL = "https://bouyguesconstruction.sharepoint.com/:x:/r/sites/bybatgocete/Shared%20Documents/05%20CVC%20PLOMBERIE/13-%20Fournisseurs/SOURCING%20PARTAGE-CET.xlsx?d=w967f744c974542f7bb359724bdb394a2&csf=1&web=1&e=FFLkVJ"

# Fonction pour charger le fichier Excel depuis SharePoint
@st.cache_data(ttl=300)  # Cache de 5 minutes
def load_excel_from_sharepoint():
    try:
        # Modifier l'URL pour forcer le téléchargement
        download_url = EXCEL_URL.replace("web=1&e=", "download=1&e=")
        
        # Télécharger le fichier
        response = requests.get(download_url, timeout=30)
        response.raise_for_status()
        
        # Lire le fichier Excel
        excel_file = BytesIO(response.content)
        df = pd.read_excel(excel_file, sheet_name='SOURCING FOURNISSEURS')
        
        return df, None
    except Exception as e:
        return None, str(e)

# Coordonnées des zones géographiques françaises
GEO_COORDINATES = {
    'Ile-de-France': [48.8566, 2.3522],
    'IDF': [48.8566, 2.3522],
    'Provence-Alpes-Côte d'Azur': [43.9352, 6.6246],
    'PACA': [43.9352, 6.6246],
    'Auvergne-Rhône-Alpes': [45.7772, 4.6305],
    'Rhône-Alpes': [45.7772, 4.6305],
    'Nouvelle-Aquitaine': [46.5667, 0.5667],
    'Aquitaine': [46.5667, 0.5667],
    'Occitanie': [43.6108, 3.8767],
    'Midi-Pyrénées': [43.6108, 3.8767],
    'Languedoc-Roussillon': [43.6108, 3.8767],
    'Bourgogne-Franche-Comté': [47.2829, 5.0317],
    'Bourgogne': [47.2829, 5.0317],
    'Franche-Comté': [47.2829, 5.0317],
    'Hauts-de-France': [50.4501, 3.8196],
    'Nord-Pas-de-Calais': [50.4501, 3.8196],
    'Normandie': [49.1829, 0.3667],
    'Bretagne': [48.1095, -3.3616],
    'Pays de la Loire': [47.2184, -1.5536],
    'Centre-Val de Loire': [47.5882, 1.5037],
    'Centre': [47.5882, 1.5037],
    'Corse': [42.0366, 8.7853],
    'Guadeloupe': [16.2667, -61.5333],
    'Réunion': [-21.1351, 55.4920],
    'Martinique': [14.6349, -61.0242],
    'Guyane': [3.9339, -53.1256],
    'Mayotte': [-12.7596, 45.2271],
    'Paris': [48.8566, 2.3522],
    'Lyon': [45.7640, 4.8357],
    'Marseille': [43.2965, 5.3698],
    'Toulouse': [43.6047, 1.4442],
    'Nice': [43.7102, 7.2620],
    'Nantes': [47.2184, -1.5536],
    'Strasbourg': [48.5734, 7.7521],
    'Montpellier': [43.6108, 3.8767],
    'Bordeaux': [44.8378, -0.5792],
    'Lille': [50.6292, 3.0573],
    'Rennes': [48.1095, -1.6802],
    'Le Havre': [49.4944, 0.1079],
    'Saint-Étienne': [45.4398, 4.3890],
    'Grenoble': [45.1885, 5.7245],
    'Dijon': [47.3220, 5.0341],
    'Angers': [47.4667, -0.5500],
    'Nîmes': [43.8345, 4.3605],
    'Toulon': [43.1256, 5.9355],
    'Douai': [50.3733, 3.0743],
    'Clermont-Ferrand': [45.7772, 3.0870],
    'Le Mans': [48.0061, 0.1996],
    'Amiens': [49.8941, 2.2959],
    'Limoges': [45.8336, 1.2611],
    'Dunkerque': [51.0364, 2.3826],
    'Brest': [48.3905, -4.4860],
    'Perpignan': [42.7039, 2.8945],
    'Mulhouse': [47.7412, 7.3362],
    'Rouen': [49.4432, 1.0977],
    'Nancy': [48.6921, 6.1844],
    'Valence': [44.9329, 4.8905],
    'Metz': [49.1193, 6.1757],
    'Thionville': [49.3600, 6.1597],
    'Lorraine': [49.1193, 6.1757],
    'Alsace': [48.5734, 7.7521],
    'Champagne-Ardenne': [48.9627, 4.0267],
    'Reims': [49.2583, 4.0317],
    'Calais': [50.9520, 1.8621],
    'Avignon': [43.9493, 4.8055],
    'Aix-en-Provence': [43.5298, 5.4455],
    'Cannes': [43.5527, 7.0174],
    'Antibes': [43.5808, 7.1239],
    'Versailles': [48.8041, 2.1302],
    'Boulogne-Billancourt': [48.8355, 2.2399],
    'Neuilly-sur-Seine': [48.8822, 2.2726],
    'La Défense': [48.8920, 2.2466],
    'Courbevoie': [48.8973, 2.2560]
}

# Charger les données
df, error = load_excel_from_sharepoint()

if error:
    st.error(f"❌ Erreur lors du chargement du fichier Excel: {error}")
    st.info("Veuillez vérifier que le fichier Excel existe sur SharePoint et que vous avez les droits d'accès.")
else:
    # Nettoyer les données
    df = df.dropna(subset=['Entreprise'])
    df = df.fillna('')
    
    # Afficher les colonnes disponibles pour debug
    with st.expander("ℹ️ Infos sur les données"):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Fournisseurs", len(df))
        with col2:
            st.metric("Catégories", df['Catégorie'].nunique())
        with col3:
            st.metric("Zones", df['Zone géographique'].nunique())
    
    # Sidebar avec filtres
    st.sidebar.header("🔍 Filtres")
    
    # Filtrer par catégorie
    categories = ['Toutes'] + sorted(df['Catégorie'].unique().tolist())
    selected_category = st.sidebar.selectbox("Catégorie", categories)
    
    if selected_category != 'Toutes':
        df_filtered = df[df['Catégorie'] == selected_category].copy()
    else:
        df_filtered = df.copy()
    
    # Filtrer par zone
    zones = ['Toutes'] + sorted(df_filtered['Zone géographique'].unique().tolist())
    selected_zone = st.sidebar.selectbox("Zone géographique", zones)
    
    if selected_zone != 'Toutes':
        df_filtered = df_filtered[df_filtered['Zone géographique'] == selected_zone].copy()
    
    # Créer la carte
    st.subheader(f"🗺️ Carte des Fournisseurs ({len(df_filtered)} affichés)")
    
    # Initialiser la carte Folium
    m = folium.Map(location=[46.2276, 2.2137], zoom_start=6, tiles='OpenStreetMap')
    
    # Ajouter les marqueurs
    displayed_count = 0
    for idx, row in df_filtered.iterrows():
        zone = str(row.get('Zone géographique', '')).strip()
        
        # Chercher les coordonnées
        coords = GEO_COORDINATES.get(zone)
        
        if coords:
            lat, lng = coords
            
            # Ajouter un peu de variation pour éviter les superpositions exactes
            import random
            jitter_lat = lat + (random.random() - 0.5) * 0.08
            jitter_lng = lng + (random.random() - 0.5) * 0.08
            
            # Créer le popup
            popup_text = f"""
            <div style="width: 250px;">
                <h4 style="color:#1e3c72; margin: 0 0 10px 0;">{row.get('Entreprise', 'N/A')}</h4>
                <b>Catégorie:</b> {row.get('Catégorie', 'N/A')}<br>
                <b>Type:</b> {row.get('Type de produits', 'N/A')}<br>
                <b>Zone:</b> {zone}<br>
                <b>Contact:</b> {row.get('contact', 'N/A')}<br>
                <b>Fonction:</b> {row.get('fonction', 'N/A')}<br>
                <b>Tél:</b> <a href="tel:{row.get('Numéro', '')}">{row.get('Numéro', 'N/A')}</a><br>
                <b>Email:</b> <a href="mailto:{row.get('Email', '')}">{row.get('Email', 'N/A')}</a><br>
                <b>Remarques:</b> {row.get('Remarques', 'N/A')}
            </div>
            """
            
            # Ajouter le marqueur
            folium.CircleMarker(
                location=[jitter_lat, jitter_lng],
                radius=8,
                popup=folium.Popup(popup_text, max_width=300),
                color='#1e3c72',
                fill=True,
                fillColor='#2a5298',
                fillOpacity=0.7,
                weight=2,
                tooltip=row.get('Entreprise', 'N/A')
            ).add_to(m)
            
            displayed_count += 1
    
    # Afficher la carte
    st_folium(m, width=1400, height=600)
    
    # Tableau des fournisseurs
    st.subheader("📋 Tableau des Fournisseurs")
    
    # Préparer les données pour l'affichage
    display_df = df_filtered[[
        'Entreprise', 'Catégorie', 'Type de produits', 'Zone géographique',
        'contact', 'fonction', 'Numéro', 'Email', 'Remarques'
    ]].copy()
    
    display_df.columns = [
        'Entreprise', 'Catégorie', 'Type', 'Zone',
        'Contact', 'Fonction', 'Téléphone', 'Email', 'Remarques'
    ]
    
    st.dataframe(display_df, use_container_width=True, height=400)
    
    # Bouton de téléchargement
    csv = display_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Télécharger les données (CSV)",
        data=csv,
        file_name="fournisseurs_cet.csv",
        mime="text/csv"
    )
    
    # Footer
    st.markdown("---")
    st.caption(f"✅ Mis à jour depuis SharePoint | {displayed_count} fournisseur(s) affichés | Dernière actualisation: {pd.Timestamp.now().strftime('%H:%M:%S')}")
