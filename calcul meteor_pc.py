#Imports
import os
import json
import pandas as pd
from datetime import date
import datetime
import numpy as np

#Changement de répertoire
#os.chdir('D:\\Nouveau dossier\\METEOR\\15-07-2022\\')
os.chdir('D:\\Nouveau dossier\\METEOR\\09-09-2022\\')
#os.chdir('D:\\Nouveau dossier\\METEOR\\')

fichier_stock = 'STOCK 21-10-2022.xlsx'
#fichier_etats_meteor = 'Etats_Meteor.xlsx'

d1 = datetime.datetime.now().strftime("%d-%m-%Y %Hh %Mm %Ss")


champs= ["Code",
"N_serie",
"Type_poste",
"Modele",
"Contrat_loueur",
"Ref_loueur",
"Date_debut",
"Date_fin",
"Date_prevue_sortie",
"Etat",
"Blanchiment",
"Grade",
"Localisation",
"Mode_acquis",
"Loueur",
"Gravité",
"Stock",
"Date_inventaire",
"Nom_IP",
"Code_compt",
"Description",
"Groupe_de_gestion",
"Commentaire",
"Utilisateur",
"Mémoire",
"Commentaire_Bien",
"Adresse Physique",
"Affectation",
"Souche",
"ID_element_parc",
"Edition_Modele",
"Nom_complet",
"Modifié_le",
"Date_derniere_modfication_status",
"Ancienne_localisation",
"Ancien_Status",
"Nature_Modele",
"SLA_Brute",
"SLA_ouvrés",
"Palier",
"Alerte_Palier",
"Today",
"Délai_Météor",
'semaine_inventaire',
'delai_inventaire',
'delai_inventaire_ouvre',
'Palier inventaire',
'Alerte Palier inventaire',
"Catégorie",
"Type",
"Departement",
"Ville",
"Site",
"Batiment",
"Etage",
"Bureau",
"Autre_1",
"Autre_2"
]

def creation_dict(fichier):
    with open(fichier,'r', encoding="utf-8") as json_data:
        return json.load(json_data)
        
#Ouverture fichier
df = pd.read_excel(fichier_stock)

dict_delai_meteor = creation_dict("delais_meteor.json")

#Ajout de colonnes
df['SLA_brute'] = ''
df['SLA_ouvrés'] = ''
df['Palier'] = ''
df['Alerte Palier'] = ''
df['Today'] = pd.to_datetime(date.today())
df['Délai Météor'] = ''
df['semaine_inventaire'] = ''
df['delai_inventaire'] = ''
df['delai_inventaire_ouvre'] = ''
df['Palier inventaire'] = ''
df['Alerte Palier inventaire'] = ''
df['Délai Météor'] = df['Etat (Bien)'].map(dict_delai_meteor)
df['SLA_brute'] = df['Today'] - df['Date de dernière modification du champ Status']
df['delai_inventaire'] = df['Today'] - df["Date d'inventaire"]
df['semaine_inventaire'] = 5

#Séparation catégorie
categorie_type = df['Nom complet (Modèle.Sous-modèle de)']
sous_tableau_categorie_type = categorie_type.str.split('/',expand=True)

#Intégration Catégorie et Type
df['Catégorie'] = sous_tableau_categorie_type[3]
df['Type'] = sous_tableau_categorie_type[4]


#Séparation Localisation
pre_localisation = df['Localisation']
#print(pre_localisation)
localisation = pre_localisation.str.split('/',expand=True).fillna('')
#print(localisation)

#Intégration Localisation
df['Departement'] = localisation[1]
df['Ville'] = localisation[2]
df['Site'] = localisation[3]
df['Batiment'] = localisation[4]
df['Etage'] = localisation[5]
df['Bureau'] = localisation[6]
df['Autre_1'] = localisation[7]
df['Autre_2'] = localisation[8]

df.fillna('')
df['Catégorie']

df_imprimantes = df[df['Catégorie']== 'Imprimante'].copy()
df_postes = df[df['Catégorie']== 'PC'].copy()

""" print(df_imprimantes)
print(df_postes)
 """


######## Dataframe pour les postes ########

df_postes['Today'] = pd.to_datetime(df_postes['Today'], errors='coerce')
df_postes['Date de dernière modification du champ Status'].fillna('2017-01-01')
df_postes["Date d\'inventaire"].fillna('2017-01-01')
df_postes['Date de dernière modification du champ Status'] = pd.to_datetime(df_postes['Date de dernière modification du champ Status'], errors='coerce')
df_postes['Date de dernière modification du champ Status'] = df_postes['Date de dernière modification du champ Status'].values.astype('datetime64[D]')
df_postes['Today'] =  df_postes['Today'].values.astype('datetime64[D]')
Todays_Date = np.datetime64('today')

#Calcul des jours ouvrés (suppression jours fériés et des week-ends)
df_postes['SLA_ouvrés'] = np.busday_count(df_postes['Date de dernière modification du champ Status'].values.astype('M8[D]'), Todays_Date,holidays = ["2022-01-01","2022-04-17", "2022-04-18","2022-05-01", "2022-05-08",
    "2022-05-26","2022-06-06", "2022-07-14","2022-08-15", "2022-11-01", "2022-11-11", "2022-12-25"])

#Calcul des jours inventaires ouvrés (suppression jours fériés et des week-ends)
df_postes['delai_inventaire_ouvre'] = np.busday_count(df_postes["Date d'inventaire"].values.astype('M8[D]'), Todays_Date,holidays = ["2022-01-01","2022-04-17", "2022-04-18","2022-05-01", "2022-05-08",
    "2022-05-26","2022-06-06", "2022-07-14","2022-08-15", "2022-11-01", "2022-11-11", "2022-12-25"])


#Calcul du Palier
df_postes.loc[df_postes['SLA_ouvrés'] < df_postes['Délai Météor'],'Palier'] = 'Sans Retard'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*2),'Palier'] = 'Palier 1'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*3),'Palier'] = 'Palier 2'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*3) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*4),'Palier'] = 'Palier 3'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*4) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*5),'Palier'] = 'Palier 4'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*5),'Palier'] = 'Palier 5'
#Calcul Alerte Palier
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*1-2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']),'Alerte Palier'] = 'Alerte Palier 1'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*2-2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*2),'Alerte Palier'] = 'Alerte Palier 2'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*3-2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*3),'Alerte Palier'] = 'Alerte Palier 3'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*4-2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*4),'Alerte Palier'] = 'Alerte Palier 4'
df_postes.loc[(df_postes['SLA_ouvrés'] >= df_postes['Délai Météor']*5-2) & (df_postes['SLA_ouvrés'] < df_postes['Délai Météor']*5),'Alerte Palier'] = 'Alerte Palier 5'
df_postes.loc[df_postes['Alerte Palier']=='','Alerte Palier'] ='Sans Alerte'

#Calcul Palier inventaire
df_postes.loc[df_postes['delai_inventaire_ouvre'] < df_postes['semaine_inventaire'] ,'Palier inventaire'] = 'Moins de 5jours ouvrés'
df_postes.loc[(df_postes['delai_inventaire_ouvre'] >= df_postes['semaine_inventaire']) & (df_postes['delai_inventaire_ouvre'] <= df_postes['semaine_inventaire']*4),'Palier inventaire'] = 'Moins de 20 jours ouvrés'
df_postes.loc[(df_postes['delai_inventaire_ouvre'] > df_postes['semaine_inventaire']*4) & (df_postes['delai_inventaire_ouvre'] <= df_postes['semaine_inventaire']*9),'Palier inventaire'] = 'Plus de 20 jours ouvrés'
df_postes.loc[(df_postes['delai_inventaire_ouvre'] > df_postes['semaine_inventaire']*9) ,'Palier inventaire'] = 'Plus de 45 jours ouvrés'

#Calcul Alerte Palier inventaire
df_postes.loc[(df_postes['delai_inventaire_ouvre'] < df_postes['semaine_inventaire']*4-2) ,'Alerte Palier inventaire'] = 'Sans Alerte'
df_postes.loc[(df_postes['delai_inventaire_ouvre'] >= df_postes['semaine_inventaire']*4-2) & (df_postes['delai_inventaire_ouvre']  <= df_postes['semaine_inventaire']*4) ,'Alerte Palier inventaire'] = 'Alerte Palier Inventaire Plus de 20 jours ouvrés'
df_postes.loc[(df_postes['delai_inventaire_ouvre'] >= df_postes['semaine_inventaire']*9-2) & (df_postes['delai_inventaire_ouvre'] <= df_postes['semaine_inventaire']*9),'Alerte Palier inventaire'] = 'Alerte Palier Inventaire Plus de 45 jours ouvrés'
df_postes.loc[df_postes['Alerte Palier inventaire']=='','Alerte Palier inventaire'] ='Sans Alerte'

df_postes.loc[df_postes['Site']=='','Site'] = 'LOCALISATION A DEFINIR'



df_postes1 = df_postes[(df_postes['Etat (Bien)']=='A restituer')|(df_postes['Etat (Bien)']=='En cours de Blanchiment')|(df_postes['Etat (Bien)']=='En réparation IFG')|(df_postes['Etat (Bien)']=='En réparation SAV')|(df_postes['Etat (Bien)']=='Réservé')]


######## Fin Dataframe pour les postes ########




######## Dataframe pour les imprimantes ########

df_imprimantes['Today'] = pd.to_datetime(df_imprimantes['Today'], errors='coerce')
df_imprimantes['Date de dernière modification du champ Status'].fillna('2017-01-01')
df_imprimantes["Date d\'inventaire"].fillna('2017-01-01')
df_imprimantes['Date de dernière modification du champ Status'] = pd.to_datetime(df_imprimantes['Date de dernière modification du champ Status'], errors='coerce')
df_imprimantes['Date de dernière modification du champ Status'] = df_imprimantes['Date de dernière modification du champ Status'].values.astype('datetime64[D]')
df_imprimantes['Today'] =  df_imprimantes['Today'].values.astype('datetime64[D]')
Todays_Date = np.datetime64('today')

#Calcul des jours ouvrés (suppression jours fériés et des week-ends)
df_imprimantes['SLA_ouvrés'] = np.busday_count(df_imprimantes['Date de dernière modification du champ Status'].values.astype('M8[D]'), Todays_Date,holidays = ["2022-01-01","2022-04-17", "2022-04-18","2022-05-01", "2022-05-08",
    "2022-05-26","2022-06-06", "2022-07-14","2022-08-15", "2022-11-01", "2022-11-11", "2022-12-25"])

#Calcul des jours inventaires ouvrés (suppression jours fériés et des week-ends)
df_imprimantes['delai_inventaire_ouvre'] = np.busday_count(df_imprimantes["Date d'inventaire"].values.astype('M8[D]'), Todays_Date,holidays = ["2022-01-01","2022-04-17", "2022-04-18","2022-05-01", "2022-05-08",
    "2022-05-26","2022-06-06", "2022-07-14","2022-08-15", "2022-11-01", "2022-11-11", "2022-12-25"])


#Calcul du Palier
df_imprimantes.loc[df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor'],'Palier'] = 'Sans Retard'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*2),'Palier'] = 'Palier 1'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*3),'Palier'] = 'Palier 2'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*3) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*4),'Palier'] = 'Palier 3'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*4) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*5),'Palier'] = 'Palier 4'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*5),'Palier'] = 'Palier 5'
#Calcul Alerte Palier
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*1-2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']),'Alerte Palier'] = 'Alerte Palier 1'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*2-2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*2),'Alerte Palier'] = 'Alerte Palier 2'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*3-2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*3),'Alerte Palier'] = 'Alerte Palier 3'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*4-2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*4),'Alerte Palier'] = 'Alerte Palier 4'
df_imprimantes.loc[(df_imprimantes['SLA_ouvrés'] >= df_imprimantes['Délai Météor']*5-2) & (df_imprimantes['SLA_ouvrés'] < df_imprimantes['Délai Météor']*5),'Alerte Palier'] = 'Alerte Palier 5'
df_imprimantes.loc[df_imprimantes['Alerte Palier']=='','Alerte Palier'] ='Sans Alerte'

#Calcul Palier inventaire
df_imprimantes.loc[df_imprimantes['delai_inventaire_ouvre'] < df_imprimantes['semaine_inventaire'] ,'Palier inventaire'] = 'Moins de 5jours ouvrés'
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] >= df_imprimantes['semaine_inventaire']) & (df_imprimantes['delai_inventaire_ouvre'] <= df_imprimantes['semaine_inventaire']*4),'Palier inventaire'] = 'Moins de 20 jours ouvrés'
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] > df_imprimantes['semaine_inventaire']*4) & (df_imprimantes['delai_inventaire_ouvre'] <= df_imprimantes['semaine_inventaire']*9),'Palier inventaire'] = 'Plus de 20 jours ouvrés'
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] > df_imprimantes['semaine_inventaire']*9) ,'Palier inventaire'] = 'Plus de 45 jours ouvrés'

#Calcul Alerte Palier inventaire
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] < df_imprimantes['semaine_inventaire']*4-2) ,'Alerte Palier inventaire'] = 'Sans Alerte'
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] >= df_imprimantes['semaine_inventaire']*4-2) & (df_imprimantes['delai_inventaire_ouvre']  <= df_imprimantes['semaine_inventaire']*4) ,'Alerte Palier inventaire'] = 'Alerte Palier Inventaire Plus de 20 jours ouvrés'
df_imprimantes.loc[(df_imprimantes['delai_inventaire_ouvre'] >= df_imprimantes['semaine_inventaire']*9-2) & (df_imprimantes['delai_inventaire_ouvre'] <= df_imprimantes['semaine_inventaire']*9),'Alerte Palier inventaire'] = 'Alerte Palier Inventaire Plus de 45 jours ouvrés'
df_imprimantes.loc[df_imprimantes['Alerte Palier inventaire']=='','Alerte Palier inventaire'] ='Sans Alerte'

df_imprimantes.loc[df_imprimantes['Site']=='','Site'] = 'LOCALISATION A DEFINIR'


######## Fin Dataframe pour les imprimantes ########


df_postes.columns = champs
df_postes1.columns = champs
df_imprimantes.columns = champs


df_postes2 = df_postes1[[
"Code",
"N_serie",
"Modele",
"Edition_Modele",
"Catégorie",
"Type",
"Date_debut",
"Date_fin",
"Date_prevue_sortie",
"Date_inventaire",
'semaine_inventaire',
'delai_inventaire',
'delai_inventaire_ouvre',
'Palier inventaire',
'Alerte Palier inventaire',
"Date_derniere_modfication_status",
"Today",
"Etat",
"Délai_Météor",
"SLA_ouvrés",
"Palier",
"Alerte_Palier",
"SLA_Brute",
"Ancien_Status",
"Ancienne_localisation",
"Blanchiment",
"Departement",
"Ville",
"Site",
"Batiment",
"Etage",
"Bureau",
"Autre_1",
"Autre_2"
]]



df_postes.to_excel(f"stock_{d1}_Etats_Meteor_Bruts_Postes.xlsx",sheet_name="Etats Météor Bruts",index=False)
df_postes2.to_excel(f"stock_{d1}_Etats_Meteor_Prioritaires_Postes.xlsx",sheet_name="Etats Météor Prioritaires",index=False)
df_imprimantes.to_excel(f"stock_{d1}_Etats_Meteor_Bruts_Imprimantes.xlsx",sheet_name="Etats Météor Imprimantes",index=False)