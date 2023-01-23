##############################################################################
# Created by: Samuel Morency and Paul Blais-Morisset
# Created on: July 2022
# Created for: Global Affairs Canada - Investment Strategy and Analysis
# 
# This module contains the constants used in investmentStockIngestion.py,
# dataframes.py, and bilateralBriefAutomation.py
# 
# If new Trade Agreements are made, or new countries are added to an existing
# Trade Agreement, add them here and the program should handle it automatically
##############################################################################

directory ='//signet/I-org/XED/BED-ELF/Research & Analysis (Investment) - Recherche et analyse (investissement)/Data & economic analysis/__Data/FDI_CDIA_Database/Python/'
toWordDirectory = '//Signet/I-org/XED/BED-ELF/Client Prom & Info Matls - Promotion et info de la clientèle/BEI/Bilateral FDI briefs/Automation project 2022/Word Automation/'

input_name = directory + 'FDI_CDIA_Database_Template.xlsx'
output_name = directory + 'FDI_CDIA_Database_test.xlsx'
input_name_fr = directory + 'FDI_CDIA_Database_Template_Fr.xlsx'
output_name_fr = directory + 'FDI_CDIA_Database_test_Fr.xlsx'
outputToWord = toWordDirectory + 'Generated.xlsm'
inputToWord = toWordDirectory + 'Template.xlsm'

former_countries = ['Union of Soviet Socialist Republics','Czechoslovakia','German Democratic Republic (East)','Yugoslavia', 'Swaziland']
former_countries_fr = ['Union des républiques socialistes soviétiques','Tchécoslovaquie','République démocratique allemande (Est)','Yougoslavie', 'Swaziland'] 

# Dictionary of trade Agreements and the countries associated with each of them
taFull = [
{'TA_name':'CETA*','countries':['Austria', 'Belgium', 'Bulgaria', 'Croatia', 'Cyprus', 'Czech Republic', 'Denmark','Estonia','Finland','France','Germany','Greece','Hungary','Ireland','Italy','Latvia','Lithuania', "Luxembourg",'Malta',"Netherlands",'Poland',"Portugal",'Romania','Slovakia','Slovenia',"Spain","Sweden"]},
{'TA_name':'CPTPP**','countries':['Australia','Japan','Mexico','New Zealand','Peru','Singapore','Vietnam']},
{'TA_name':'EFTA***','countries':['Iceland','Liechtenstein','Norway','Switzerland']},
{'TA_name':'NAFTA****','countries':['Mexico','United States']},
{'TA_name':'CARICOM*****','countries':['Antigua and Barbuda', 'Bahamas', 'Barbados', 'Belize', 'Dominica', 'Grenada', 'Guyana', 'Haiti', 'Jamaica', 'Montserrat', 'Saint Kitts and Nevis', 'Saint Lucia', 'Saint Vincent and the Grenadines', 'Suriname', 'Trinidad and Tobago']},
{'TA_name':'CARICOM****** (associates)','countries':['Anguilla', 'Bermuda', 'British Virgin Islands', 'Cayman Islands', 'Turks and Caicos']},
{'TA_name':'MENA*******','countries':['Algeria','Bahrain','Egypt','Iran','Iraq','Israel','Jordan',"Kuwait",'Lebanon',"Libya",'Morocco','Oman','Qatar','Saudi Arabia','Syria','Tunisia','United Arab Emirates', 'West Bank and Gaza', 'Yemen']},
{'TA_name':'ASEAN********','countries':['Brunei Darussalam', 'Cambodia', 'Indonesia', 'Lao PDR', 'Malaysia', 'Myanmar', 'Philippines', 'Singapore', 'Thailand', 'Vietnam']}
      ]

taFullFr = [
{'TA_name':'AECG*','countries':['Autriche', 'Belgique', 'Bulgarie', 'Croatie', 'Chypre', 'République tchèque', 'Danemark','Estonie','Finlande','France','Allemagne','Grèce','Hongrie','Irlande','Italie','Lettonie','Lituanie', "Luxembourg",'Malte',"Pays-Bas",'Pologne',"Portugal",'Roumanie','Slovaquie','Slovénie',"Espagne","Suède"]},
{'TA_name':'PTPGP**','countries':['Australie','Japon','Mexique','Nouvelle-Zélande','Pérou','Singapour','Vietnam']},
{'TA_name':'AELE***','countries':['Islande','Liechtenstein','Norvège','Suisse']},
{'TA_name':'ALENA / ACEUM****','countries':['Mexique','États-Unis']},
{'TA_name':'CARICOM*****','countries':['Antigua-et-Barbuda', 'Bahamas', 'Barbade', 'Bélize', 'Dominique', 'Grenade', 'Guyana', 'Haïti', 'Jamaïque', 'Montserrat', 'Saint-Kitts-et-Nevis', 'Sainte-Lucie', 'Saint-Vincent-et-les-Grenadines', 'Suriname', 'Trinité-et-Tobago']},
{'TA_name':'CARICOM****** (associés)','countries':['Bermudes', 'Îles Vierges britanniques', 'Îles Caïmans']},
{'TA_name':'MOAN*******','countries':['Algérie','Bahreïn','Cisjordanie et Gaza','Égypte','Iran','Iraq','Israël','Jordanie',"Koweït",'Liban',"Libye",'Maroc','Oman','Qatar','Arabie saoudite','Syrie','Tunisie','Émirats arabes unis','Yémen']},
{'TA_name':'ANASE********','countries':['Brunei Darussalam', 'Cambodge', 'Indonésie', 'Laos', 'Malaisie', 'Myanmar', 'Philippines', 'Singapour', 'Thaïlande', 'Vietnam']}
      ]

# Blurb at the bottom of each FDI/CDIA Excel sheet
formatTAEng = {
    'note': 'Note: The red zero indicate that the data are either not available or confidential',
    'CETA': '* CETA partners include: Austria, Belgium, Bulgaria, Croatia, Cyprus, Czech Republic, Denmark, Estonia, Finland, France, Germany, Greece, Hungary, Ireland, Italy, Latvia, Lithuania, Luxembourg, Malta, Netherlands, Poland, Portugal, Romania, Slovakia, Slovenia, Spain and Sweden.',
    'CPTPP': "** CPTPP partners include: Australia, Japan, Mexico, New Zealand, Peru, Singapore and Vietnam.",
    'EFTA': '*** EFTA partners include: Iceland, Liechtenstein, Norway and Switzerland.',
    'NAFTA': '**** NAFTA / CUSMA partners include Mexico and the United States.',
    'CARICOM': '***** CARICOM members include: Antigua and Barbuda, Bahamas, Barbados, Belize, Dominica, Grenada, Guyana, Haiti, Jamaica, Montserrat, Saint Kitts and Nevis, Saint Lucia, Saint Vincent and the Grenadines, Suriname, and Trinidad and Tobago.',
    'associate': '****** CARICOM associate members include: Anguilla, Bermuda, British Virgin Islands, Cayman Islands, and Turks & Caicos.',
    'MENA': '******* MENA partners include Algeria, Bahrain, Egypt, Iran, Iraq, Israel, Jordan, Kuwait, Lebanon, Libya, Morocco, Oman, Qatar, Saudi Arabia, Syria, Tunisia, United Arab Emirates, West Bank and Gaza, and Yemen. Average annual growth rate was calculated for 2012-2020.',
    'ASEAN': '******** ASEAN members include: Brunei Darussalam, Cambodia, Indonesia, Lao PDR, Malaysia, Myanmar, Philippines, Singapore, Thailand, Vietnam.',
    }

formatTAFr = {
    'note': 'Note : Un zéro en rouge indiquent que les données ne sont pas disponibles ou sont confidentielles.',
    'CETA': '* Les partenaires de l’AECG comprennent l’Autriche, la Belgique, la Bulgarie, la Croatie, Chypre, la République tchèque, le Danemark, l’Estonie, la Finlande, la France, l’Allemagne, la Grèce, la Hongrie, l’Irlande, l’Italie, la Lettonie, la Lituanie, le Luxembourg, Malte, les Pays-Bas, la Pologne, le Portugal, la Roumanie, la Slovaquie, la Slovénie, l’Espagne et la Suède.',
    'CPTPP': "** Les partenaires du PTPGP comprennent l’Australie, le Japon, le Mexique, la Nouvelle-Zélande, le Pérou, Singapour et le Vietnam.",
    'EFTA': '*** Les partenaires de l’AELE comprennent l’Islande, le Liechtenstein, la Norvège et la Suisse.',
    'NAFTA': '**** Les partenaires de l’ALENA / l’ACEUM comprennent le Mexique et les États-Unis.',
    'CARICOM': '***** Les membres de la CARICOM comprennent: Antigua-et-Barbuda, les Bahamas, la Barbade, le Belize, la Dominique, la Grenade, la Guyane, Haïti, la Jamaïque, Montserrat, Saint-Kitts-et-Nevis, Sainte-Lucie, Saint-Vincent-et-les-Grenadines, le Suriname et Trinité-et-Tobago.',
    'associate': '****** Les membres associés de la CARICOM comprennent : Anguilla, les Bermudes, les îles Vierges britanniques, les îles Caïmans et les îles Turks et Caicos.',
    'MENA': '******* Les partenaires du MOAN comprennent l’Algérie, Bahreïn, la Cisjordanie et Gaza, l’Égypte, l’Iran, l’Iraq, Israël, la Jordanie, le Koweït, le Liban, la Libye, le Maroc, Oman, le Qatar, l’Arabie saoudite, la Syrie, la Tunisie, les Émirats arabes unis et le Yémen.',
    'ASEAN': "******** Les partenaires de l'ANASE comprennent: le Brunei Darussalam, le Cambodge, l'Indonésie, la république démocratique populaire du Laos, la Malaisie, le Myanmar, les Philippines, Singapour, la Thaïlande et le Vietnam.",
    }

structureEng = {
    'share': 'Share in ',
    'growth': 'Growth rate ',
    'cagr': 'Average annual growth rate ',
    'all': 'All countries',
    'regRank': 'Regional Rank',
    'other': 'Other ',
    'glRank': 'Global Rank'
    }

structureFr = {
    'share': 'Part en ',
    'growth': 'Taux de croissance ',
    'cagr': 'Taux de croissance annuel moyen ',
    'all': 'Ensemble des pays',
    'regRank': 'Rang régional',
    'other': "Autres pays d'",
    'glRank': 'Rang mondial'
    }

formatEng = {
    'all': 'All countries',
    'loc': list(taFull[0].values())[0],
    'fta': "FTA/Regional Partners",
    'cdia': 'CDIA UIC-All',
    'rest': "Rest of the world"
    }

formatFr = {
    'all': 'Ensemble des pays',
    'loc': list(taFullFr[0].values())[0],
    'fta': "ALE/Partenaires Régionaux",
    'cdia': 'IDEC PIU-Tout',
    'rest': "Reste du monde"
    }
