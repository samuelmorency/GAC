##############################################################################
# Created by: Samuel Morency and Paul Blais-Morisset
# Created on: July 2022
# Created for: Global Affairs Canada - Investment Strategy and Analysis
# 
# This program downloads the StatsCan tables 36-10-0008-01, 36-10-0433-01,
# (fdi/cdia IIC and fdi/cdia UIC, respectively)
# 36-10-0659-01, 36-10-0582-01, and 36-10-0470-01,
# (industry, AMNE, and CMNE, respectively)
# and prepares their dataframes for use in investmentStockIngestion.py
# 
# User must update pandas to run this module
##############################################################################

import pandas as pd
import numpy as np
from stats_can import StatsCan
from datetime import datetime
import copy
import constants as c


sc = StatsCan()
sc.update_tables()
df_iic = sc.table_to_df("36-10-0008-01")
df_uic = sc.table_to_df("36-10-0433-01")
df_industry = sc.table_to_df("36-10-0659-01")
df_AMNE = sc.table_to_df("36-10-0582-01")
df_CMNE = sc.table_to_df("36-10-0470-01")
#sc.downloaded_tables


now = datetime.now()
#retrievedDate = datetime.strftime(now, '%B') + ' ' + datetime.strftime(now,'%Y')
retrievedDate = now.strftime('%B %Y')

monthNum = int(now.strftime('%m'))
yearNum = int(now.strftime('%Y'))

if monthNum == 2:
    releaseDate = 'February ' + str(yearNum)
    releaseDateFr = 'Février ' + str(yearNum)
    nextReleaseDate = 'May ' + str(yearNum)
    nextReleaseDateFr = 'mai ' + str(yearNum)
else:
    releaseDate = 'May ' + str(yearNum)
    releaseDateFr = 'Mai ' + str(yearNum)
    nextReleaseDate = 'February ' + str(yearNum+1)
    nextReleaseDateFr = 'février ' + str(yearNum+1)


# Gets intersection of two lists
def intersection(list1, list2):
    intersect = [value for value in list1 if value in list2]
    return intersect

def minYear(dframe, maxYr):
    if maxYr - 9 in dframe['year'].unique():
        minYr = maxYr - 9
    else:
        minYr = dframe['year'].min()
    return minYr

def prepareDf(df, minYr):
    df.drop(columns=['UOM','UOM_ID','SCALAR_FACTOR','SCALAR_ID','VECTOR','COORDINATE','GEO','DGUID','STATUS','SYMBOL','TERMINATED','DECIMALS'],inplace=True)
    df.drop(df[df.year < minYr].index, inplace=True)
    df.drop(df[df['Countries or regions'].isin(c.former_countries)].index, inplace=True)
    df.drop(df[df['Countries or regions'].isin(c.former_countries_fr)].index, inplace=True)
    df.loc[df['VALUE'].isnull(), 'VALUE'] = 0

def breakdownMNE(dfToCopy,filters,startYear):
    df = dfToCopy.copy(deep=True)
    for fil in filters:
        df = df.loc[df[fil[0]] == fil[1]]
    prepareDf(df,startYear)
    return df

# Generates region dictionaries
def makeRegDict(regs, df, suffix):
    reg = []
    foo = 0
    e = 'é' if regs == regions_fr else 'e'
    name = 'Nom' if regs == regions_fr else 'Clean name'
    for i in regs:
        locals()['reg%s%s' % (suffix, foo)] = df[df['R%sgions' % e] == i].drop(columns='R%sgions' % e)
        locals()['reg%s%s' % (suffix, foo)] =locals()['reg%s%s' % (suffix, foo)][name].values.tolist()
        locals()['dict%s%s' % (suffix, foo)] ={
            'reg_name':i,'countries':locals()['reg%s%s' % (suffix, foo)]
            }
        reg.append(locals()['dict%s%s' % (suffix, foo)])
        foo +=1
    return reg

# Structures a dataframe the same way as the Excel workbook
def structureData(table, firstyear, lastyear, regdict, tradeAgreement, talist, form, regs, hasOther):
    table= pd.pivot_table(table, values="VALUE", index='Countries or regions' , columns='year' )
    # Generate statistics
    table[form['share']+str(lastyear)] = np.nan
    table[form['growth']+str(lastyear-1)+'-'+str(lastyear)] = np.nan
    table[form['cagr']+str(firstyear)+'-'+str(lastyear)] = np.nan
    
    # Remove Canada to exclude from ranking
    if (regdict==reg_UIC or regdict==reg_UIC_fr):
        tempCan = table.loc['Canada', lastyear]
        table.loc['Canada', lastyear] = np.nan
    
    # Subgroup each region and calculate regional rank
    tempList = [table.loc[regs].loc[[form['all']]]]
    otherList = []
    for reg in regdict : 
        tempDf = table.loc[reg['countries']].sort_values(by=['Countries or regions'])
        if hasOther:
            other = tempDf[tempDf.index == form['other'] + reg['reg_name']]
            tempDf = tempDf.drop(form['other'] + reg['reg_name'])
            otherList.append(form['other'] + reg['reg_name'])
        tempDf[form['regRank']]=tempDf[tempDf[lastyear] != 0][lastyear].rank(method='min',ascending=False)
        tempList.append(table.loc[[reg['reg_name']]])
        tempList.append(tempDf)
        if hasOther:
            tempList.append(other)
        
    dataframe = pd.concat(tempList)
 
    # Calculate global rankings and append to dataframe
    gr = dataframe[~dataframe.index.isin(regs)]
    gr = gr[~gr.index.isin(talist)]
    gr = gr[~gr.index.isin(otherList)]
    gr[form['glRank']]=gr[gr[lastyear] != 0][lastyear].rank(method= 'min',ascending=False )
    gr = gr[form['glRank']]
    dataframe= dataframe.merge(gr, how='left', left_index=True, right_index=True)
    
    # Re-insert Canada
    if (regdict==reg_UIC or regdict==reg_UIC_fr):
        dataframe.loc['Canada', lastyear] = tempCan
    
    # Add Trade Agreements to dataframe
    for agreement in tradeAgreement:
        for i in range(lastyear-firstyear+1):
            mySum = 0
            for country in agreement['countries']:
                mySum += dataframe.loc[country,firstyear+i]
            dataframe.at[agreement['TA_name'],firstyear+i] = mySum
    return dataframe

def main():
    #sc.update_tables()
    print('You just ran main() in the dataframes.py module.')

df_iic['REF_DATE'] = df_iic['REF_DATE'].dt.year
df_uic['REF_DATE'] = df_uic['REF_DATE'].dt.year
df_industry['REF_DATE'] = df_industry['REF_DATE'].dt.year
df_AMNE['REF_DATE'] = df_AMNE['REF_DATE'].dt.year
df_CMNE['REF_DATE'] = df_CMNE['REF_DATE'].dt.year

df_iic.rename(columns = {'REF_DATE':'year'}, inplace = True)
df_uic.rename(columns = {'REF_DATE':'year'}, inplace = True)
df_industry.rename(columns = {'REF_DATE':'year'}, inplace = True)
df_AMNE.rename(columns = {'REF_DATE':'year'}, inplace = True)
df_AMNE.rename(columns = {'Ultimate investing country':'Countries or regions'}, inplace = True)
df_CMNE.rename(columns = {'REF_DATE':'year'}, inplace = True)

cname = pd.read_excel(c.directory+'Excel/Country name conversion.xlsx','Country name')
cnames = cname[['Clean name', 'Adjectivals', 'Regions', 'Prefix','Nom','Adjectif', 'Adjectif Plural', 'Féminin', 'Féminin Plural','Régions','Article','Préfixe2','Préfixe3']]
cnames = cnames.set_index('Clean name')

result_iic = pd.merge(left=df_iic, right=cname,left_on = 'Countries or regions' ,right_on= 'Original name',how="left")
result_uic = pd.merge(left=df_uic, right=cname,left_on = 'Countries or regions' ,right_on= 'Original name',how="left")
result_AMNE = pd.merge(left=df_AMNE, right=cname,left_on = 'Countries or regions' ,right_on= 'Original name',how="left")
result_CMNE = pd.merge(left=df_CMNE, right=cname,left_on = 'Countries or regions' ,right_on= 'Original name',how="left")

#New code 4/27/2023
#result_iic['Clean name'].fillna(result_iic['Countries or regions'], inplace=True)
#result_uic['Clean name'].fillna(result_uic['Countries or regions'], inplace=True)
#result_AMNE['Clean name'].fillna(result_AMNE['Countries or regions'], inplace=True)
#result_CMNE['Clean name'].fillna(result_CMNE['Countries or regions'], inplace=True)


result_iic.sort_values(by=['year', 'Canadian and foreign direct investment', 'Countries or regions'], inplace=True)
result_uic.sort_values(by=['year', 'Type of direct investment', 'Countries or regions'], inplace=True)
result_CMNE.sort_values(by=['year', 'Countries or regions'], inplace=True)
result_AMNE.sort_values(by=['year', 'Countries or regions'], inplace=True)
df_iic.sort_values(by=['year', 'Canadian and foreign direct investment', 'Countries or regions'], inplace=True)
df_uic.sort_values(by=['year', 'Type of direct investment', 'Countries or regions'], inplace=True)
df_AMNE.sort_values(by=['year', 'Countries or regions'], inplace=True)
df_CMNE.sort_values(by=['year', 'Countries or regions'], inplace=True)

df_iic['Countries or regions']= result_iic['Clean name'].values
df_uic['Countries or regions']= result_uic['Clean name'].values
df_AMNE['Countries or regions']= result_AMNE['Clean name'].values
df_CMNE['Countries or regions']= result_CMNE['Clean name'].values

df_iic_fr = df_iic.copy(deep=True)
df_uic_fr = df_uic.copy(deep=True)
df_iic_fr['Countries or regions']= result_iic['Nom'].values
df_uic_fr['Countries or regions']= result_uic['Nom'].values

maxyear = df_iic['year'].max()
maxyearIndustry = df_industry['year'].max()
maxyearAMNE = df_AMNE['year'].max()
maxyearCMNE = df_CMNE['year'].max()

minyear = minYear(df_iic, maxyear)
minyearIndustry = minYear(df_industry, maxyearIndustry)
minyearAMNE = minYear(df_AMNE, maxyearAMNE)
minyearCMNE = minYear(df_CMNE, maxyearCMNE)

#New Feb 8 code
df_CMNE['Trade statistics'] = pd.Categorical(df_CMNE['Trade statistics'], categories=df_CMNE['Trade statistics'].unique())
df_CMNE=df_CMNE.groupby(['Countries or regions', 'Trade statistics', 'year'], as_index=False).first()

#df_AMNE['Trade statistics'] = pd.Categorical(df_AMNE['Trade statistics'], categories=df_AMNE['Trade statistics'].unique())
#df_AMNE=df_AMNE.groupby(['Countries or regions', 'Trade statistics'], as_index=False).first()


# Preparing the dataframes
prepareDf(df_iic, minyear)
df_fdi_iic = df_iic.loc[df_iic['Canadian and foreign direct investment'] == 'Foreign direct investment in Canada - total book value']
df_cdia_iic = df_iic.loc[df_iic['Canadian and foreign direct investment'] == 'Canadian direct investment abroad - total book value']
prepareDf(df_uic, minyear)
df_fdi_uic = df_uic.loc[df_uic['Type of direct investment'] == 'Foreign direct investment in Canada by ultimate investor country - total book value']
df_cdia_uic = df_uic.loc[df_uic['Type of direct investment'] == 'Canadian direct investment abroad by ultimate investor country – total book value']
prepareDf(df_iic_fr, minyear)
df_fdi_iic_fr = df_iic_fr.loc[df_iic_fr['Canadian and foreign direct investment'] == 'Foreign direct investment in Canada - total book value']
df_cdia_iic_fr = df_iic_fr.loc[df_iic_fr['Canadian and foreign direct investment'] == 'Canadian direct investment abroad - total book value']
prepareDf(df_uic_fr, minyear)
df_fdi_uic_fr = df_uic_fr.loc[df_uic_fr['Type of direct investment'] == 'Foreign direct investment in Canada by ultimate investor country - total book value']
df_cdia_uic_fr = df_uic_fr.loc[df_uic_fr['Type of direct investment'] == 'Canadian direct investment abroad by ultimate investor country – total book value']

minyear_fdi_uic = minYear(df_fdi_uic, maxyear)
minyear_cdia_uic = minYear(df_cdia_uic, maxyear)

# Breakdown MNEs into variables
df_CMNE_Employees = breakdownMNE(df_CMNE,[['Trade statistics','Number of employees']],minyearCMNE)
df_CMNE_Assets = breakdownMNE(df_CMNE,[['Trade statistics','Total Assets']],minyearCMNE)
df_CMNE_Liabilities = breakdownMNE(df_CMNE,[['Trade statistics','Total Liabilities']],minyearCMNE)
df_CMNE_Sales = breakdownMNE(df_CMNE,[['Trade statistics','Value of sales in dollars']],minyearCMNE)
df_AMNE_Enterprises = breakdownMNE(df_AMNE,[['Descriptive variable','Number of enterprises'],['North American Industry Classification System','All industries']],minyearAMNE)
df_AMNE_Jobs = breakdownMNE(df_AMNE,[['Descriptive variable','Total number of jobs'],['North American Industry Classification System','All industries']],minyearAMNE)
df_AMNE_GDP = breakdownMNE(df_AMNE,[['Descriptive variable','Gross domestic product at basic prices (value added)'],['North American Industry Classification System','All industries']],minyearAMNE)
df_AMNE_Revenues = breakdownMNE(df_AMNE,[['Descriptive variable','Operating revenues'],['North American Industry Classification System','All industries']],minyearAMNE)
df_AMNE_Imports = breakdownMNE(df_AMNE,[['Descriptive variable','Merchandise imports'],['North American Industry Classification System','All industries']],minyearAMNE)
df_AMNE_Exports = breakdownMNE(df_AMNE,[['Descriptive variable','Merchandise exports'],['North American Industry Classification System','All industries']],minyearAMNE)


dfCountriesAndRegions_IIC = result_iic[['Clean name','Regions']].drop_duplicates().dropna(subset=['Regions'],axis='rows')
dfCountriesAndRegions_UIC = result_uic[['Clean name','Regions']].drop_duplicates().dropna(subset=['Regions'],axis='rows')
dfCountriesAndRegions_CMNE = result_CMNE[['Clean name','Regions']].drop_duplicates().dropna(subset=['Regions'],axis='rows')
dfCountriesAndRegions_AMNE = result_AMNE[['Clean name','Regions']].drop_duplicates().dropna(subset=['Regions'],axis='rows')
# in the AMNE database, North America and South and Central america are grouped
# into "Americas". If that every changes, remove the next two lines of code.
dfCountriesAndRegions_AMNE['Regions'] = dfCountriesAndRegions_AMNE['Regions'].replace(['North America'],'Americas')
dfCountriesAndRegions_AMNE['Regions'] = dfCountriesAndRegions_AMNE['Regions'].replace(['South and Central America'],'Americas')
##############################################################################
dfCountriesAndRegions_IIC_fr = result_iic[['Nom','Régions']].drop_duplicates().dropna(subset=['Régions'],axis='rows')
dfCountriesAndRegions_UIC_fr = result_uic[['Nom','Régions']].drop_duplicates().dropna(subset=['Régions'],axis='rows')
dfCountriesAndRegions_IIC.drop(dfCountriesAndRegions_IIC[dfCountriesAndRegions_IIC['Clean name'].isin(c.former_countries)].index, inplace=True)
dfCountriesAndRegions_UIC.drop(dfCountriesAndRegions_UIC[dfCountriesAndRegions_UIC['Clean name'].isin(c.former_countries)].index, inplace=True)
dfCountriesAndRegions_CMNE.drop(dfCountriesAndRegions_CMNE[dfCountriesAndRegions_CMNE['Clean name'].isin(c.former_countries)].index, inplace=True)
dfCountriesAndRegions_AMNE.drop(dfCountriesAndRegions_AMNE[dfCountriesAndRegions_AMNE['Clean name'].isin(c.former_countries)].index, inplace=True)
dfCountriesAndRegions_IIC_fr.drop(dfCountriesAndRegions_IIC_fr[dfCountriesAndRegions_IIC_fr['Nom'].isin(c.former_countries_fr)].index, inplace=True)
dfCountriesAndRegions_UIC_fr.drop(dfCountriesAndRegions_UIC_fr[dfCountriesAndRegions_UIC_fr['Nom'].isin(c.former_countries_fr)].index, inplace=True)

# Generate trade Agreement dictionaries excluding countries not in data
TA = copy.deepcopy(c.taFull)
TA_UIC = copy.deepcopy(c.taFull)
TA_fr = copy.deepcopy(c.taFullFr)
TA_UIC_fr = copy.deepcopy(c.taFullFr)
TA_CMNE = copy.deepcopy(c.taFull)
TA_AMNE = copy.deepcopy(c.taFull)
for ta in TA:
    ta['countries'] = intersection(ta['countries'], df_iic['Countries or regions'].tolist())
for ta in TA_UIC:
    ta['countries'] = intersection(ta['countries'], df_uic['Countries or regions'].tolist())
for ta in TA_fr:
    ta['countries'] = intersection(ta['countries'], df_iic_fr['Countries or regions'].tolist())
for ta in TA_UIC_fr:
    ta['countries'] = intersection(ta['countries'], df_uic_fr['Countries or regions'].tolist())
for ta in TA_CMNE:
    ta['countries'] = intersection(ta['countries'], df_CMNE['Countries or regions'].tolist())
for ta in TA_AMNE:
    ta['countries'] = intersection(ta['countries'], df_AMNE['Countries or regions'].tolist())
TA_list= [a_dict['TA_name'] for a_dict in TA]
TA_list_fr= [a_dict['TA_name'] for a_dict in TA_fr]
regions = dfCountriesAndRegions_IIC['Regions'].drop_duplicates().values.tolist()
regions_CMNE = dfCountriesAndRegions_CMNE['Regions'].drop_duplicates().values.tolist()
regions_AMNE = dfCountriesAndRegions_AMNE['Regions'].drop_duplicates().values.tolist()
regions_fr = dfCountriesAndRegions_IIC_fr['Régions'].drop_duplicates().values.tolist()

# Generate region dictionaries
reg = makeRegDict(regions, dfCountriesAndRegions_IIC, '')
reg_UIC = makeRegDict(regions, dfCountriesAndRegions_UIC, '_UIC_')
reg_CMNE =  makeRegDict(regions_CMNE, dfCountriesAndRegions_CMNE, '_CMNE_')
reg_AMNE =  makeRegDict(regions_AMNE, dfCountriesAndRegions_AMNE, '_AMNE_')
reg_fr = makeRegDict(regions_fr, dfCountriesAndRegions_IIC_fr, '_fr')
reg_UIC_fr = makeRegDict(regions_fr, dfCountriesAndRegions_UIC_fr, '_UIC_fr')
regions.insert(0,"All countries")
regions_CMNE.insert(0,"All countries")
regions_AMNE.insert(0,"All countries")
regions_fr.insert(0,"Ensemble des pays")


if monthNum == 1:
    retrievedDateFr = "janvier " + str(yearNum)
elif monthNum == 2:
    retrievedDateFr = "février " + str(yearNum)
elif monthNum == 3:
    retrievedDateFr = "mars " + str(yearNum)
elif monthNum == 4:
    retrievedDateFr = "avril " + str(yearNum)
elif monthNum == 5:
    retrievedDateFr = "mai " + str(yearNum)
elif monthNum == 6:
    retrievedDateFr = "juin " + str(yearNum)
elif monthNum == 7:
    retrievedDateFr = "juillet " + str(yearNum)
elif monthNum == 8:
    retrievedDateFr = "août " + str(yearNum)
elif monthNum == 9:
    retrievedDateFr = "septembre " + str(yearNum)
elif monthNum == 10:
    retrievedDateFr = "octobre " + str(yearNum)
elif monthNum == 11:
    retrievedDateFr = "novembre " + str(yearNum)
elif monthNum == 12:
    retrievedDateFr = "décembre " + str(yearNum)
else:
    retrievedDateFr = "n.a."

#cnames.drop_duplicates(subset=['Clean name'])
uniqueCnames = cnames[~cnames.index.duplicated(keep='first')]

#df_fdi_iic_t = structureData(df_fdi_iic, minyear, maxyear, reg, TA, TA_list, constants.structureEng, regions, False)
#df_fdi_iic.loc[constants.taFull[0]['TA_name'],firstyear:structure['share']+str(lastyear)]=dataframe.loc[x['countries']].sum(axis=0)

#df_cdia_uic_t = structureData(df_cdia_uic, minyear_cdia_uic, maxyear, reg_UIC, TA_UIC, TA_list, constants.structureEng, regions, True)
#df_cdia_uic_t = pd.concat([df_cdia_uic_t.loc[['All countries','Canada']],df_cdia_uic_t.drop(['All countries','Canada'])],ignore_index=False)

if __name__ == "__main__":
    main()


#df_CMNE_names = df_CMNE['Countries or regions'].unique()
#result_CMNE_names = result_CMNE['Countries or regions'].unique()

#df_CMNE_names_lst = df_CMNE_names.tolist()
#result_CMNE_names_lst = result_CMNE_names.tolist()

