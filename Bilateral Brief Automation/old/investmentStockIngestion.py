##############################################################################
# Created by: Samuel Morency and Paul Blais-Morisset
# Created on: July 2022
# Created for: Global Affairs Canada - Investment Strategy and Analysis
# 
# This program uses the StatsCan tables 36-10-0008-01, 36-10-0433-01,
# (fdi/cdia IIC and fdi/cdia UIC, respectively)
# 36-10-0659-01, 36-10-0582-01, and 36-10-0470-01,
# (industry, AMNE, and CMNE, respectively)
# and automatically generates the FDI_CDIA database and bilat excel worksheets
# 
# User must update pandas to run this module
##############################################################################



countryBilatsToGenerate = ['Australia', 'Brazil', 'Belgium', 'Chile', 'China', 'Denmark', 'France', 'Germany', 'Hong Kong', 'India', 'Ireland', 'Italy', 'Japan', 'Luxembourg', 'Mexico', 'Netherlands', 'Norway', 'Singapore', 'South Korea', 'Spain', 'Sweden', 'Switzerland', 'Taiwan', 'United Arab Emirates', 'United Kingdom', 'United States']
#countryBilatsToGenerate = ['Africa']
reGenerateExcelFile = True
generateBilats = True
bilatsVersionCode = "V15"






if reGenerateExcelFile:
    import pandas as pd
    import xlwings as xw
    from xlwings.utils import rgb_to_int
    import constants
    import dataframes
if generateBilats:
    import sys
    sys.path.append('I:/XED/BED-ELF/Client Prom & Info Matls - Promotion et info de la clientèle/BEI/Bilateral FDI briefs/Automation project 2022/Word Automation')
    import bilateralBriefAutomation


def formatting(dframe, sheet, name_range, formTA, form, regs):
    rows, cols = dframe.shape
    startcol = "A"
    startrow = 4
    lastcol = chr(ord(startcol)+cols)
    lastrow = startrow+rows
    start=startcol+str(startrow)
    sheet.range(start).value=dframe
    temp_range = dframe.index.get_loc(form['loc'])
    
    if sheet.name == form['cdia']:
        sheet.range('A7:Q7').api.Insert()
        sheet.range('A7').value = form['rest']
        sheet.range(chr(ord(startcol)+1)+'7'':'+chr(ord(lastcol)-5)+'7').value = '='+chr(ord(startcol)+1)+'5-'+chr(ord(startcol)+1)+'6'
        rows += 1
        lastrow += 1
        temp_range += 1
    
    #formatting 
    all_table=sheet.range(start+":"+lastcol+str(lastrow))
    all_table.api.Font.Size = 10
    all_table.api.Font.Color = rgb_to_int((16,58,95))
    all_table.column_width = 10
    sheet.range(start+":"+lastcol+str(startrow)).api.Font.Bold=True
    sheet.range(start).column_width = 30
    sheet.range(chr(ord(startcol)+1)+str(startrow+1)+":"+chr(ord(lastcol)-5)+str(lastrow)).number_format="#,##0"
    sheet.range(chr(ord(lastcol)-4)+str(startrow+1)+":"+chr(ord(lastcol)-2)+str(lastrow)).number_format="0.0%"
    sheet.range(chr(ord(lastcol)-4)+str(startrow)+":"+lastcol+str(lastrow)).api.Font.Bold=True
    sheet.range(chr(ord(startcol)+1)+str(startrow+1)+":"+chr(ord(lastcol)-5)+str(lastrow)).api.FormatConditions.Add(1,3, '=0')
    sheet.range(chr(ord(startcol)+1)+str(startrow+1)+":"+chr(ord(lastcol)-5)+str(lastrow)).api.FormatConditions(1).Font.Color=rgb_to_int((255,0,0))
    
    sheet.range(start).expand().name = name_range
    sheet.range(start).expand('right').name=name_range+"_col"
    
    # Write footnote under table
    i = 0
    for taBlurb in formTA.values():
        sheet.range(startcol+str(lastrow+i+3)).value = taBlurb
        i += 1
    
    for j in regs:
        t_range=dframe.index.get_loc(j)
        sheet.range(startcol+str(startrow+1+t_range)+":"+lastcol+str(startrow+1+t_range)).api.Font.Bold=True        
        if j != form['all']: sheet.range(startcol+str(startrow+1+t_range)+":"+lastcol+str(startrow+1+t_range)).api.Borders(9).LineStyle =  1 
   
    sheet.range(startcol+str(startrow+temp_range+1)+":"+lastcol+str(startrow+temp_range+1)).api.Insert()
    rows += 1
    lastrow += 1
    sheet.range(start).expand('down').name=name_range+"_row"
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Borders(9).LineStyle =  1 
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Borders(9).Weight =  2
    sheet.range(startcol+str(startrow+temp_range+1)).value = form['fta']
    sheet.range(startcol+str(startrow+temp_range+1)).api.Font.Bold=True
    
    sheet.range(chr(ord(lastcol)-4)+str(startrow+1)+':'+chr(ord(lastcol)-4)+str(lastrow)).formula = '=IFERROR($'+chr(ord(lastcol)-5)+str(startrow+1)+'/$'+chr(ord(lastcol)-5)+'$'+str(startrow+1)+',"")'
    sheet.range(chr(ord(lastcol)-3)+str(startrow+1)+':'+chr(ord(lastcol)-3)+str(lastrow)).formula = '=IFERROR(($'+chr(ord(lastcol)-5)+str(startrow+1)+'/$'+chr(ord(lastcol)-6)+str(startrow+1)+')-1,"")'
    sheet.range(chr(ord(lastcol)-2)+str(startrow+1)+':'+chr(ord(lastcol)-2)+str(lastrow)).formula = '=IFERROR(RRI('+str((ord(lastcol)-5)-(ord(startcol)+1))+',$'+chr(ord(startcol)+1)+str(startrow+1)+',$'+chr(ord(lastcol)-5)+str(startrow+1)+'),"")'
    sheet.range(chr(ord(startcol)+1)+str(lastrow-8)+':'+chr(ord(lastcol))+str(lastrow-8)).value = ''
    
    
def english(wb):
    df_fdi_iic_t = dataframes.structureData(dataframes.df_fdi_iic, dataframes.minyear, dataframes.maxyear, dataframes.reg, dataframes.TA, dataframes.TA_list, constants.structureEng, dataframes.regions, False)
    df_cdia_iic_t = dataframes.structureData(dataframes.df_cdia_iic, dataframes.minyear, dataframes.maxyear, dataframes.reg, dataframes.TA, dataframes.TA_list, constants.structureEng, dataframes.regions, False)
    df_fdi_uic_t = dataframes.structureData(dataframes.df_fdi_uic, dataframes.minyear_fdi_uic, dataframes.maxyear, dataframes.reg_UIC, dataframes.TA_UIC, dataframes.TA_list, constants.structureEng, dataframes.regions, True)
    df_cdia_uic_t = dataframes.structureData(dataframes.df_cdia_uic, dataframes.minyear_cdia_uic, dataframes.maxyear, dataframes.reg_UIC, dataframes.TA_UIC, dataframes.TA_list, constants.structureEng, dataframes.regions, True)
    
    df_cdia_uic_t = pd.concat([df_cdia_uic_t.loc[['All countries','Canada']],df_cdia_uic_t.drop(['All countries','Canada'])],ignore_index=False)
    
    wb.sheets['FDI IIC-All'].range('A4').value = df_fdi_iic_t
    wb.sheets['CDIA IIC-All'].range('A4').value = df_cdia_iic_t
    print('IIC data updated')
    wb.sheets['FDI UIC-ALL'].range('A4').value = df_fdi_uic_t
    wb.sheets['CDIA UIC-All'].range('A4').value = df_cdia_uic_t
    print('UIC data updated')
    formatting(df_fdi_iic_t,wb.sheets['FDI IIC-All'],'FDI_IIC', constants.formatTAEng, constants.formatEng, dataframes.regions)
    formatting(df_cdia_iic_t,wb.sheets['CDIA IIC-All'],'CDIA_IIC', constants.formatTAEng, constants.formatEng, dataframes.regions)
    print('IIC data formatted')
    formatting(df_fdi_uic_t,wb.sheets['FDI UIC-ALL'],'FDI_UIC', constants.formatTAEng, constants.formatEng, dataframes.regions)
    formatting(df_cdia_uic_t,wb.sheets['CDIA UIC-All'],'CDIA_UIC', constants.formatTAEng, constants.formatEng, dataframes.regions)
    print('UIC data formatted')
    

def french(wb):
    df_fdi_iic_fr_t = dataframes.structureData(dataframes.df_fdi_iic_fr, dataframes.minyear, dataframes.maxyear, dataframes.reg_fr, dataframes.TA_fr, dataframes.TA_list_fr, constants.structureFr, dataframes.regions_fr, False)
    df_cdia_iic_fr_t = dataframes.structureData(dataframes.df_cdia_iic_fr, dataframes.minyear, dataframes.maxyear, dataframes.reg_fr, dataframes.TA_fr, dataframes.TA_list_fr, constants.structureFr, dataframes.regions_fr, False)
    df_fdi_uic_fr_t = dataframes.structureData(dataframes.df_fdi_uic_fr, dataframes.minyear_fdi_uic, dataframes.maxyear, dataframes.reg_UIC_fr, dataframes.TA_UIC_fr, dataframes.TA_list_fr, constants.structureFr, dataframes.regions_fr, True)
    df_cdia_uic_fr_t = dataframes.structureData(dataframes.df_cdia_uic_fr, dataframes.minyear_cdia_uic, dataframes.maxyear, dataframes.reg_UIC_fr, dataframes.TA_UIC_fr, dataframes.TA_list_fr, constants.structureFr, dataframes.regions_fr, True)
    
    df_cdia_uic_fr_t = pd.concat([df_cdia_uic_fr_t.loc[['Ensemble des pays','Canada']],df_cdia_uic_fr_t.drop(['Ensemble des pays','Canada'])],ignore_index=False)
    
    wb.sheets['IDE PII-Tout'].range('A4').value = df_fdi_iic_fr_t
    wb.sheets['IDEC PII-Tout'].range('A4').value = df_cdia_iic_fr_t
    wb.sheets['IDE PIU-Tout'].range('A4').value = df_fdi_uic_fr_t
    wb.sheets['IDEC PIU-Tout'].range('A4').value = df_cdia_uic_fr_t
    formatting(df_fdi_iic_fr_t,wb.sheets['IDE PII-Tout'],'FDI_IIC', constants.formatTAFr, constants.formatFr, dataframes.regions_fr)
    formatting(df_cdia_iic_fr_t,wb.sheets['IDEC PII-Tout'],'CDIA_IIC', constants.formatTAFr, constants.formatFr, dataframes.regions_fr)
    formatting(df_fdi_uic_fr_t,wb.sheets['IDE PIU-Tout'],'FDI_UIC', constants.formatTAFr, constants.formatFr, dataframes.regions_fr)
    formatting(df_cdia_uic_fr_t,wb.sheets['IDEC PIU-Tout'],'CDIA_UIC', constants.formatTAFr, constants.formatFr, dataframes.regions_fr)

def rawSheets(wb, data, dframe, uid, name):    
    wb.sheets[data].range('A1').value = dframe
    wb.sheets[data].range("A1").value = 'UID'
    wb.sheets[data].range("A2").expand('down').value = uid
    wb.sheets[data].range('A1').expand().name = name
    wb.sheets[data].range('A1').expand('right').name = name + '_col'
    wb.sheets[data].range('A1').expand('down').name = name + '_row'

def cmneRank(wb, data, dframe, name):
    dframe = dataframes.structureData(dframe, dataframes.minyearCMNE, dataframes.maxyearCMNE, dataframes.reg_CMNE, dataframes.TA_CMNE, dataframes.TA_list, constants.structureEng, dataframes.regions, False)
    formatting(dframe,wb.sheets[data],name,constants.formatTAEng, constants.formatEng, dataframes.regions)

def amneRank(wb, data, dframe, name):
    dframe = dataframes.structureData(dframe, dataframes.minyearAMNE, dataframes.maxyearAMNE, dataframes.reg_AMNE, dataframes.TA_AMNE, dataframes.TA_list, constants.structureEng, dataframes.regions_AMNE, True)
    formatting(dframe,wb.sheets[data],name,constants.formatTAEng, constants.formatEng, dataframes.regions_AMNE)
    

def bilatWorkbook():
    xw.Book(constants.inputToWord).set_mock_caller()
    wb = xw.Book.caller()
    
    wb.sheets['Countries'].range('A1').value = dataframes.uniqueCnames
    wb.sheets['Countries'].range('A1').expand().name = 'countries'
    wb.sheets['Countries'].range('A1').expand('right').name = 'countries_col'
    wb.sheets['Countries'].range('A1').expand('down').name = 'countries_row'
    print('countries done')
    
    english(wb)
    
    rawSheets(wb, 'Industry Data', dataframes.df_industry, "=B2&E2&F2&G2", 'industry')
    print('Industry data updated')
    rawSheets(wb, 'AMNE Data', dataframes.df_AMNE, "=B2&E2&F2&G2", 'amne')
    print('AMNE data updated')
    
    amneRank(wb, 'AMNE Enterprises', dataframes.df_AMNE_Enterprises, 'amneEnterprises')
    amneRank(wb, 'AMNE Jobs', dataframes.df_AMNE_Jobs, 'amneJobs')
    amneRank(wb, 'AMNE GDP', dataframes.df_AMNE_GDP, 'amneGDP')
    amneRank(wb, 'AMNE Revenues', dataframes.df_AMNE_Revenues, 'amneRevenues')
    amneRank(wb, 'AMNE Imports', dataframes.df_AMNE_Imports, 'amneImports')
    amneRank(wb, 'AMNE Exports', dataframes.df_AMNE_Exports, 'amneExports')
    print('Done splitting AMNE variables')
    #rawSheets(wb, 'CMNE Data', dataframes.df_CMNE, "=B2&E2&F2", 'cmne')
    
    cmneRank(wb, 'CMNE Employees', dataframes.df_CMNE_Employees, 'CMNEemployees')
    cmneRank(wb, 'CMNE Assets', dataframes.df_CMNE_Assets, 'CMNEassets')
    cmneRank(wb, 'CMNE Liabilities', dataframes.df_CMNE_Liabilities, 'CMNEliabilities')
    cmneRank(wb, 'CMNE Sales', dataframes.df_CMNE_Sales, 'CMNEsales')
    print('Done splitting CMNE variables')
    
    wb.sheets['PANEL'].range('B4').value = dataframes.releaseDate
    wb.names['releaseDate'].refers_to = dataframes.releaseDate
    wb.sheets['PANEL'].range('B5').value = dataframes.retrievedDate
    wb.names['retrievedDate'].refers_to = dataframes.retrievedDate
    wb.sheets['PANEL'].range('B6').value = dataframes.nextReleaseDate
    wb.names['nextReleaseDate'].refers_to = dataframes.nextReleaseDate
    wb.sheets['PANEL'].range('B31').value = dataframes.releaseDateFr
    wb.names['releaseDateFr'].refers_to = dataframes.releaseDateFr
    wb.sheets['PANEL'].range('B32').value = dataframes.retrievedDateFr
    wb.names['retrievedDateFr'].refers_to = dataframes.retrievedDateFr
    wb.sheets['PANEL'].range('B33').value = dataframes.nextReleaseDateFr
    wb.names['nextReleaseDateFr'].refers_to = dataframes.nextReleaseDateFr
    
    wb.save(constants.outputToWord)
    wb.close()


def databaseWorkbook():
    xw.Book(constants.directory+'FDI CDIA Database Templates/'+constants.input_name).set_mock_caller()
    wb = xw.Book.caller()
    english(wb)
    wb.save(constants.output_name)
    wb.close()
    
    #xw.Book(constants.directory+'FDI CDIA Database Templates/'+constants.input_name_fr).set_mock_caller()
    #wb=xw.Book.caller()
    #french(wb)
    #wb.save(constants.output_name_fr)
    #wb.close()


def main():
    if reGenerateExcelFile:
        #databaseWorkbook()
        bilatWorkbook()
    if generateBilats:
        for country in countryBilatsToGenerate:
            bilateralBriefAutomation.word(country, bilatsVersionCode)
    print("You just ran main() in the investmentStockIngestion.py module.")
    
if __name__ == "__main__":
    main()
    
