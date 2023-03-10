# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
import streamlit as st
import time
import plotly.express as px
import plotly.graph_objects as go
warnings.filterwarnings('ignore')
from st_aggrid import AgGrid, GridUpdateMode, JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
from user import *
import streamlit.components.v1 as stc
from streamlit_option_menu import option_menu
st.set_page_config(
        page_title="RATP",
        page_icon="chart_with_upwards_trend",
        layout="wide",
    )
from sign import *
if "estimer" not in st.session_state:
    st.session_state.estimer =False
if "qte" not in st.session_state:
        myFile = open("quantite_projet.xlsx", "w+")
        dataframe=pd.DataFrame(columns=['Sous Système', 'N° préstation', 'Désignation','Travaux','Quantité','Taux forfaitaire unitaire JOUR',"Taux forfaitaire unitaire NUIT LONGUE","Fournitures unitaires","CMP","Délai d'appro","SAV"])
        st.session_state.qte = dataframe
        dataframe.to_excel("quantite_projet.xlsx",index=False)
if "pswd" not in st.session_state:
    st.session_state.pswd = ''       

if "data" not in st.session_state:
        st.session_state.data = load_data("BPU.xlsx")

if "eq" not in st.session_state:
        st.session_state.eq = load_data("Equipements.xlsx")
if "syst" not in st.session_state:
    st.session_state.syst = load_data("Sous_Systeme.xlsx")
if "soc" not in st.session_state:
        st.session_state.soc = load_data("prestation_equipement.xlsx")
if "con" not in st.session_state:
        st.session_state.con = False
if "admin" not in st.session_state:
        st.session_state.admin = False
st.markdown('<style>body{background-color:blue;}</style>',unsafe_allow_html=True)    
HTML_BANNER = """
    <div style="background-color:#034980;padding:10px;border-radius:15px">
    <h1 style="color:white;text-align:center;">Outil de chiffrage détaillé RATP </h1>
    <h2 style="color:white;text-align:center;">Courants faibles et télécom</h2>
    </div>
    """  
stc.html(HTML_BANNER)

st.sidebar.image("R.png", use_column_width=True)

def main() :
            if st.session_state.con==True:
                                with st.sidebar:
                                           res = option_menu("MENU", ['ACCEUIL','BPU', 'EQUIPEMENTS',"ASSOCIATION",'QUANTITES DU PROJET','ESTIMATION DES COUTS'],key="id1",
                                                     icons=['house', 'list-task','gear' ,"table", 'activity'],
                                                     menu_icon="app-indicator", default_index=0,
                                                     styles={
                                    "container": {"padding": "5!important", "background-color": "#5cb8a7"},
                                    "icon": {"color": "blue", "font-size": "25px"}, 
                                    "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#c8dfe3"},
                                    "nav-link-selected": {"background-color": "#034980"},
                                          }
                                    )

                                if res=='ACCEUIL':
                                      Welcome()
                                elif res=='BPU':
                                        st.markdown(""" <style> .font {
                                font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
                                </style> """, unsafe_allow_html=True)
                                        st.markdown('<p class="font">BASE DES PRIX UNITAIRES</p>', unsafe_allow_html=True)
                                        col1, col2 ,col3 = st.columns(3)
                                        with col1:pass
                                        with col3:pass
                                        with col2:
                                               res = option_menu("MENU", ['BPU par défaut', 'Importer une BPU'],key="id10",
                                                         icons=['house', 'list-task'],
                                                         menu_icon="app-indicator", default_index=0,
                                                         styles={
                                        "container": {"padding": "5!important", "background-color": "#5cb8a7"},
                                        "icon": {"color": "blue", "font-size": "25px"}, 
                                        "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#c8dfe3"},
                                        "nav-link-selected": {"background-color": "#034980"},
                                              }
                                        )
                                        if res=='BPU par défaut':
                                              
                                              f1()
                                        else:
                                            if st.session_state.admin==False:
                                                placeholder1 = st.empty()
                                                uploaded_file = placeholder1.file_uploader("Importer la BPU", accept_multiple_files=False)
                                                if uploaded_file is not None:
                                                    st.session_state.data=pd.read_excel(uploaded_file)
                                                    placeholder1.empty()
                                                    f1()
                                            else:
                                                st.warning("Impossible d'importer la BPU")

                                elif res=='EQUIPEMENTS':
                                            st.markdown(""" <style> .font {
                                font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
                                </style> """, unsafe_allow_html=True)
                                            st.markdown('<p class="font">EQUIPEMENTS</p>', unsafe_allow_html=True)
                                            
                                            col1, col2 = st.columns(2)
                                            with col1:pass
                                            
                                            with col2:
                                                 res = option_menu("MENU", ["Liste d'équipements par défaut", "Importer une liste d'équipements"],key="id133",
                                                         icons=['house', 'list-task'],
                                                         menu_icon="app-indicator", default_index=0,
                                                         styles={
                                        "container": {"padding": "5!important", "background-color": "#5cb8a7"},
                                        "icon": {"color": "blue", "font-size": "25px"}, 
                                        "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#c8dfe3"},
                                        "nav-link-selected": {"background-color": "#034980"},
                                              }
                                        )
                                            if res=="Liste d'équipements par défaut":
                                                  f2()
                                            else:
                                                if st.session_state.admin==False:
                                                    placeholder1 = st.empty()
                                                    uploaded_file = placeholder1.file_uploader("Importer la liste d'équipements", accept_multiple_files=False)

                                                    if uploaded_file is not None:
                                                        st.session_state.eq=pd.read_excel(uploaded_file)
                                                        placeholder1.empty()
                                                        f2()
                                                        
                                                else:
                                                    st.warning("Impossible d'importer la liste des équipements")
                                            
                                            
                                elif res=='QUANTITES DU PROJET':
                                            st.markdown(""" <style> .font {
                                font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
                                </style> """, unsafe_allow_html=True)
                                            st.markdown('<p class="font">QUANTITES DU PROJET</p>', unsafe_allow_html=True)
                                            manage_quantite()

                                elif res=='ESTIMATION DES COUTS':
                                            st.markdown(""" <style> .font {
                                font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
                                </style> """, unsafe_allow_html=True)
                                            st.markdown('<p class="font">ESTIMATION DES COUTS</p>', unsafe_allow_html=True)
                                            estimation_totale()

                                elif res=='ASSOCIATION':
                                            st.markdown(""" <style> .font {
                                font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
                                </style> """, unsafe_allow_html=True)
                                            st.markdown('<p class="font">PRESTATION - EQUIPEMENT</p>', unsafe_allow_html=True)
                                            association()

                                else:
                                      pass
            else:
                     pass
        
        
            
def get_base64(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


def set_background(png_file):
    bin_str = get_base64(png_file)
    page_bg_img = '''
    <style>
    .stApp {
    background-image: url("data:image/png;base64,%s");
    background-size: cover;
    }
    </style>
    ''' % bin_str
    st.markdown(page_bg_img, unsafe_allow_html=True)



def Welcome():
    st.markdown(""" <style> .font {
            font-size:22px ; font-family: 'Cooper Black'; color: #016C9A;} 
            </style> """, unsafe_allow_html=True)
    st.markdown('<p class="font">Un pro logiciel développé en interne avec une interface utilisateur simple et pratique qui permettra le chiffrage détaillé des Coûts des Travaux Courants Faibles . </p>', unsafe_allow_html=True)
    st.markdown(""" <style> .font {
            font-size:22px ; font-family: 'Cooper Black'; color: #016C9A;} 
            </style> """, unsafe_allow_html=True)
    st.markdown('<p class="font">Les caractéristiques</p>', unsafe_allow_html=True)
    col1,col2,col3=st.columns(3)
    with col1:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Simplicité ,clarté et rapidité </h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col2:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Automatisme et Ergonomie</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col3:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Chiffrage de bonne qualité</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    col4,col5,col6=st.columns(3)
    with col4:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Données évolutives</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col5:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Outil intuitif et pratique</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col6:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Dimensionnement détaillé</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    st.markdown('<p class="font">Les bénifices</p>', unsafe_allow_html=True)
    col1,col2,col3=st.columns(3)
    with col1:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Uniformisation des pratiques de chiffrage au sein de l’équipe CFA</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col2:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Source de données structurées, fiables et exploitables</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    with col3:
            st.markdown('<style>body{background-color:green;}</style>',unsafe_allow_html=True)    
            HTML_BANNER = """
            <div style="background-color:#034980;padding:10px;border-radius:15px">
            <h2 style="color:white;text-align:center;">Répondre à la demande  de chiffrages détaillés</h2>
            
            </div>
            """  
            stc.html(HTML_BANNER)
    
    st.markdown(""" <style> .font {
            font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
            </style> """, unsafe_allow_html=True)
    st.markdown('<p class="font">Architecture générale et fonctionnalités</p>', unsafe_allow_html=True)
    st.image("Structure1.png")
    st.markdown(""" <style> .font {
            font-size:35px ; font-family: 'Cooper Black'; color: #95c2a9;} 
            </style> """, unsafe_allow_html=True)
    st.markdown('<p class="font">La logique relationnelle du modéle de données</p>', unsafe_allow_html=True)
    st.image("Structure2.png")
def connexion():
       
        with st.sidebar:
            choice = option_menu("Connexion", ["Se connecter","S'inscrire","Se déconnecter"],key="id2",
                         icons=[  'person lines fill','bi bi-person-plus-fill','gear'],
                         menu_icon="app-indicator", default_index=0,
                         styles={
        "container": {"padding": "5!important", "background-color": "#5cb8a7"},
        "icon": {"color": "blue", "font-size": "25px"}, 
        "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#c8dfe3"},
        "nav-link-selected": {"background-color": "#034980"},
              }
        )
       
        if choice == "Se connecter":
            if st.session_state.con==True:
                        pass
            else:
                placeholder1 = st.sidebar.empty()
                username =placeholder1.text_input("Nom d'utilisateur")
                placeholder2 = st.sidebar.empty()
                password = placeholder2.text_input("Mot de passe",type='password')
                st.session_state.admin=password_enteredV1(username,password)
                placeholder3 = st.sidebar.empty()
                if placeholder3.button("Se connecter"):
                    placeholder1.empty()
                    placeholder2.empty()
                    placeholder3.empty()
                    create_usertable()
                    hashed_pswd = make_hashes(password)
                    result = login_user(username,check_hashes(password,hashed_pswd))
                    if result:
                        placeholder1 = st.sidebar.empty()
                        placeholder1.success("Connexion :{}".format(username))
                        time.sleep(2)
                        placeholder1.empty()
                        st.session_state.con=True
                    else:
                         st.session_state.con=False
                         placeholder1 = st.sidebar.empty()
                         placeholder1.warning("Nom d'utilsateur ou mot de passe incorrecte")
                         time.sleep(2)
                         placeholder1.empty()
                        
        elif choice == "S'inscrire":
            if (st.session_state.con==False) :
                placeholder1 = st.sidebar.empty()
                new_user =placeholder1.text_input("Nom d'utilisateur")
                placeholder2 = st.sidebar.empty()
                new_password = placeholder2.text_input("Mot de passe",type='password')
                placeholder3 = st.sidebar.empty()
                if placeholder3.button("s'inscrire"):
                    placeholder1.empty()
                    placeholder2.empty()
                    placeholder3.empty()
                    create_usertable()
                    user_result = view_all_users()
                    clean_db = pd.DataFrame(user_result,columns=["Username","Password"])
                    if clean_db.shape[0]==0:
                        create_usertable()
                    else:
                        pass
                    if ((new_user in clean_db['Username'].unique())):
                            placeholder1 = st.sidebar.empty()
                            placeholder1.warning("Nom d'utilisateur déja existant")
                            time.sleep(2)
                            placeholder1.empty()
                    else:
                            add_userdata(new_user,make_hashes(new_password))
                            placeholder1 = st.sidebar.empty()
                            placeholder1.success("Félicitations, Votre compte a été bien créé")
                            time.sleep(2)
                            placeholder1.empty()
                            
        else:
            st.session_state.admin=False
            st.session_state.con=False
            placeholder1 = st.sidebar.empty()
            placeholder1.warning("Vous etes déconnectés")
            time.sleep(2)
            placeholder1.empty()
            
set_background('cables.jpg')                            
connexion()
main()