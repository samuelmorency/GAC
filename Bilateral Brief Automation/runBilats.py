##############################################################################
# Created by: Samuel Morency and Paul Blais-Morisset
# Created on: July 2022
# Created for: Global Affairs Canada - Investment Strategy and Analysis
# 
# This program uses the StatsCan tables 36-10-0008-01, 36-10-0433-01,
# (fdi/cdia IIC and fdi/cdia UIC, respectively)
# 36-10-0659-01, 36-10-0582-01, and 36-10-0470-01,
# (industry, AMNE, and CMNE, respectively)
# and automatically generates FDI/CDIA bilateral briefs
##############################################################################


countryBilatsToGenerate = ['Australia', 'Brazil', 'Chile', 'China', 'France', 'Germany', 'Hong Kong', 'India', 'Ireland', 'Italy', 'Japan', 'Luxembourg', 'Mexico', 'Netherlands', 'Norway', 'Singapore', 'South Korea', 'Spain', 'Sweden', 'Switzerland', 'Taiwan', 'United Arab Emirates', 'United Kingdom', 'United States']
#countryBilatsToGenerate = ['Chile']
reGenerateExcelFile = True
generateBilats = False
convertPDFs = False
releaseDate = 'April 2023'
releaseDateFr = 'Avril 2023'
nextReleaseDate = 'January 2024'
nextReleaseDateFr = 'Janvier 2024'









import constants as c

if reGenerateExcelFile:
    import pandas as pd
    import xlwings as xw
    from xlwings.utils import rgb_to_int
    import dataframes as d
if generateBilats:
    import sys
    import bilateralBriefAutomation
    from docx2pdf import convert
    import datetime
    current_time = datetime.datetime.now()
    bilatsVersionCode = current_time.strftime("%Y-%m-%d_%Hh%M")
    newpath = c.directory + "Bilats/" + bilatsVersionCode + "/"
    import os
    import shutil
    if not os.path.exists(newpath):
        os.makedirs(newpath + "English/PDFs/")
        os.makedirs(newpath + "French/PDFs/")
        os.makedirs(newpath + "Charts/")


def formatting(dframe, sheet, name_range, formTA, form, regs):
    rows, cols = dframe.shape
    startcol = "A"
    startrow = 4
    lastcol = chr(ord(startcol)+cols)
    lastrow = startrow+rows
    start=startcol+str(startrow)
    sheet.range(start).value=dframe
    temp_range = dframe.index.get_loc(form['loc'])+1
    
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
    
    if sheet.name == form['cdia']:
        sheet.range('A7:Q7').api.Insert()
        sheet.range('A7').value = form['rest']
        sheet.range(chr(ord(startcol)+1)+'7'':'+chr(ord(lastcol)-5)+'7').value = '='+chr(ord(startcol)+1)+'5-'+chr(ord(startcol)+1)+'6'
        rows += 1
        lastrow += 1
        temp_range += 1
    
    sheet.range(chr(ord(lastcol)-4)+str(startrow+1)+':'+chr(ord(lastcol)-4)+str(lastrow)).formula = '=IFERROR($'+chr(ord(lastcol)-5)+str(startrow+1)+'/$'+chr(ord(lastcol)-5)+'$'+str(startrow+1)+',"")'
    sheet.range(chr(ord(lastcol)-3)+str(startrow+1)+':'+chr(ord(lastcol)-3)+str(lastrow)).formula = '=IFERROR(($'+chr(ord(lastcol)-5)+str(startrow+1)+'/$'+chr(ord(lastcol)-6)+str(startrow+1)+')-1,"")'
    sheet.range(chr(ord(lastcol)-2)+str(startrow+1)+':'+chr(ord(lastcol)-2)+str(lastrow)).formula = '=IFERROR(RRI('+str((ord(lastcol)-5)-(ord(startcol)+1))+',$'+chr(ord(startcol)+1)+str(startrow+1)+',$'+chr(ord(lastcol)-5)+str(startrow+1)+'),"")'
    
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Insert()
    rows += 1
    lastrow += 1
    temp_range += 1
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Insert()
    sheet.range(start).expand('down').name=name_range+"_row"
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Borders(9).LineStyle =  1 
    sheet.range(startcol+str(startrow+temp_range)+":"+lastcol+str(startrow+temp_range)).api.Borders(9).Weight =  2
    sheet.range(startcol+str(startrow+temp_range)).value = form['fta']
    sheet.range(startcol+str(startrow+temp_range)).api.Font.Bold=True
    
    #New code April 11 2023
    sheet.range(start).expand('down').name=name_range+"_row"
    
def english(wb):
    df_fdi_iic_t = d.structureData(d.df_fdi_iic, d.minyear, d.maxyear, d.reg, d.TA, d.TA_list, c.formatEng, d.regions, False)
    df_cdia_iic_t = d.structureData(d.df_cdia_iic, d.minyear, d.maxyear, d.reg, d.TA, d.TA_list, c.formatEng, d.regions, False)
    df_fdi_uic_t = d.structureData(d.df_fdi_uic, d.minyear_fdi_uic, d.maxyear, d.reg_UIC, d.TA_UIC, d.TA_list, c.formatEng, d.regions, True)
    df_cdia_uic_t = d.structureData(d.df_cdia_uic, d.minyear_cdia_uic, d.maxyear, d.reg_UIC, d.TA_UIC, d.TA_list, c.formatEng, d.regions, True)
    
    df_cdia_uic_t = pd.concat([df_cdia_uic_t.loc[['All countries','Canada']],df_cdia_uic_t.drop(['All countries','Canada'])],ignore_index=False)
    
    wb.sheets['FDI IIC-All'].range('A4').value = df_fdi_iic_t
    wb.sheets['CDIA IIC-All'].range('A4').value = df_cdia_iic_t
    print('IIC data updated')
    wb.sheets['FDI UIC-ALL'].range('A4').value = df_fdi_uic_t
    wb.sheets['CDIA UIC-All'].range('A4').value = df_cdia_uic_t
    print('UIC data updated')
    formatting(df_fdi_iic_t,wb.sheets['FDI IIC-All'],'FDI_IIC', c.formatTAEng, c.formatEng, d.regions)
    formatting(df_cdia_iic_t,wb.sheets['CDIA IIC-All'],'CDIA_IIC', c.formatTAEng, c.formatEng, d.regions)
    print('IIC data formatted')
    formatting(df_fdi_uic_t,wb.sheets['FDI UIC-ALL'],'FDI_UIC', c.formatTAEng, c.formatEng, d.regions)
    formatting(df_cdia_uic_t,wb.sheets['CDIA UIC-All'],'CDIA_UIC', c.formatTAEng, c.formatEng, d.regions)
    print('UIC data formatted')
    

def french(wb):
    df_fdi_iic_fr_t = d.structureData(d.df_fdi_iic_fr, d.minyear, d.maxyear, d.reg_fr, d.TA_fr, d.TA_list_fr, c.formatFr, d.regions_fr, False)
    df_cdia_iic_fr_t = d.structureData(d.df_cdia_iic_fr, d.minyear, d.maxyear, d.reg_fr, d.TA_fr, d.TA_list_fr, c.formatFr, d.regions_fr, False)
    df_fdi_uic_fr_t = d.structureData(d.df_fdi_uic_fr, d.minyear_fdi_uic, d.maxyear, d.reg_UIC_fr, d.TA_UIC_fr, d.TA_list_fr, c.formatFr, d.regions_fr, True)
    df_cdia_uic_fr_t = d.structureData(d.df_cdia_uic_fr, d.minyear_cdia_uic, d.maxyear, d.reg_UIC_fr, d.TA_UIC_fr, d.TA_list_fr, c.formatFr, d.regions_fr, True)
    
    df_cdia_uic_fr_t = pd.concat([df_cdia_uic_fr_t.loc[['Ensemble des pays','Canada']],df_cdia_uic_fr_t.drop(['Ensemble des pays','Canada'])],ignore_index=False)
    
    wb.sheets['IDE PII-Tout'].range('A4').value = df_fdi_iic_fr_t
    wb.sheets['IDEC PII-Tout'].range('A4').value = df_cdia_iic_fr_t
    wb.sheets['IDE PIU-Tout'].range('A4').value = df_fdi_uic_fr_t
    wb.sheets['IDEC PIU-Tout'].range('A4').value = df_cdia_uic_fr_t
    formatting(df_fdi_iic_fr_t,wb.sheets['IDE PII-Tout'],'FDI_IIC', c.formatTAFr, c.formatFr, d.regions_fr)
    formatting(df_cdia_iic_fr_t,wb.sheets['IDEC PII-Tout'],'CDIA_IIC', c.formatTAFr, c.formatFr, d.regions_fr)
    formatting(df_fdi_uic_fr_t,wb.sheets['IDE PIU-Tout'],'FDI_UIC', c.formatTAFr, c.formatFr, d.regions_fr)
    formatting(df_cdia_uic_fr_t,wb.sheets['IDEC PIU-Tout'],'CDIA_UIC', c.formatTAFr, c.formatFr, d.regions_fr)

def rawSheets(wb, data, dframe, uid, name):    
    wb.sheets[data].range('A1').value = dframe
    wb.sheets[data].range("A1").value = 'UID'
    wb.sheets[data].range("A2").expand('down').value = uid
    wb.sheets[data].range('A1').expand().name = name
    wb.sheets[data].range('A1').expand('right').name = name + '_col'
    wb.sheets[data].range('A1').expand('down').name = name + '_row'

def cmneRank(wb, data, dframe, name):
    dframe = d.structureData(dframe, d.minyearCMNE, d.maxyearCMNE, d.reg_CMNE, d.TA_CMNE, d.TA_list, c.formatEng, d.regions_CMNE, False)
    formatting(dframe,wb.sheets[data],name,c.formatTAEng, c.formatEng, d.regions)

def amneRank(wb, data, dframe, name):
    dframe = d.structureData(dframe, d.minyearAMNE, d.maxyearAMNE, d.reg_AMNE, d.TA_AMNE, d.TA_list, c.formatEng, d.regions_AMNE, True)
    formatting(dframe,wb.sheets[data],name,c.formatTAEng, c.formatEng, d.regions_AMNE)
    

def bilatWorkbook():
    xw.Book(c.inputToWord).set_mock_caller()
    wb = xw.Book.caller()
    
    wb.sheets['Countries'].range('A1').value = d.uniqueCnames
    wb.sheets['Countries'].range('A1').expand().name = 'countries'
    wb.sheets['Countries'].range('A1').expand('right').name = 'countries_col'
    wb.sheets['Countries'].range('A1').expand('down').name = 'countries_row'
    print('countries done')
    
    english(wb)
    
    rawSheets(wb, 'Industry Data', d.df_industry, "=B2&E2&F2&G2", 'industry')
    print('Industry data updated')
    rawSheets(wb, 'AMNE Data', d.df_AMNE, "=B2&E2&F2&G2", 'amne')
    print('AMNE data updated')
    
    amneRank(wb, 'AMNE Enterprises', d.df_AMNE_Enterprises, 'amneEnterprises')
    amneRank(wb, 'AMNE Jobs', d.df_AMNE_Jobs, 'amneJobs')
    amneRank(wb, 'AMNE GDP', d.df_AMNE_GDP, 'amneGDP')
    amneRank(wb, 'AMNE Revenues', d.df_AMNE_Revenues, 'amneRevenues')
    amneRank(wb, 'AMNE Imports', d.df_AMNE_Imports, 'amneImports')
    amneRank(wb, 'AMNE Exports', d.df_AMNE_Exports, 'amneExports')
    print('Done splitting AMNE variables')
    #rawSheets(wb, 'CMNE Data', dataframes.df_CMNE, "=B2&E2&F2", 'cmne')
    
    cmneRank(wb, 'CMNE Employees', d.df_CMNE_Employees, 'CMNEemployees')
    cmneRank(wb, 'CMNE Assets', d.df_CMNE_Assets, 'CMNEassets')
    cmneRank(wb, 'CMNE Liabilities', d.df_CMNE_Liabilities, 'CMNEliabilities')
    cmneRank(wb, 'CMNE Sales', d.df_CMNE_Sales, 'CMNEsales')
    print('Done splitting CMNE variables')
    
    # Dynamically update the dates in the panel so that if the cell references
    # change it doesn't break.
    for a_cell in wb.sheets['PANEL'].used_range:
        if a_cell.value == 'date':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = releaseDate
            wb.names['releaseDate'].refers_to = releaseDate
            #print('date: row: '+str(a_cell.row))
        elif a_cell.value == 'retrieved':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = d.retrievedDate
            wb.names['retrievedDate'].refers_to = d.retrievedDate
            #print('retrieved: row: '+str(a_cell.row))
        elif a_cell.value == 'nextrelease':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = nextReleaseDate
            wb.names['nextReleaseDate'].refers_to = nextReleaseDate
            #print('nextrelease: row: '+str(a_cell.row))
        elif a_cell.value == 'date_Fr':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = releaseDateFr
            wb.names['releaseDateFr'].refers_to = releaseDateFr
            #print('date_Fr: row: '+str(a_cell.row))
        elif a_cell.value == 'retrieved_Fr':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = d.retrievedDateFr
            wb.names['retrievedDateFr'].refers_to = d.retrievedDateFr
            #print('retrieved_Fr: row: '+str(a_cell.row))
        elif a_cell.value == 'nextrelease_Fr':
            wb.sheets['PANEL'].range('B'+str(a_cell.row)).value = nextReleaseDateFr
            wb.names['nextReleaseDateFr'].refers_to = nextReleaseDateFr
            #print('nextrelease_Fr: row: '+str(a_cell.row))
    
    wb.save(c.outputToWord)
    wb.close()


def databaseWorkbook():
    xw.Book(c.directory+'FDI CDIA Database Templates/'+c.input_name).set_mock_caller()
    wb = xw.Book.caller()
    english(wb)
    wb.save(c.output_name)
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
        shutil.copy(c.outputToWord, newpath+"Automated Analysis.xlsm")
        for country in countryBilatsToGenerate:
            bilateralBriefAutomation.word(country, bilatsVersionCode)
        if convertPDFs:
            print("Converting bilats to PDF")
            print("English: ")
            convert(c.directory +'Bilats/'+bilatsVersionCode+'/English/',
                    c.directory +'Bilats/'+bilatsVersionCode+'/English/PDFs/')
            print("French: ")
            convert(c.directory +'Bilats/'+bilatsVersionCode+'/French/',
                    c.directory +'Bilats/'+bilatsVersionCode+'/French/PDFs/')
    print("You just ran main() in RunBilats.py.")
    
if __name__ == "__main__":
    main()
    
