'''
Created on 17/01/2015

@author: mixibird
'''
import os
import csv
import time
import datetime
#import matplotlib.dates as mdates

rowNum = 0
#data = []
newRow = [0]*20 
newRow_changed = 0
###### CLASSES and FUNCTION ###### 



###### FILES ######     
# open input data file
of = open('Forbrugsoversigt 2013-2014.csv', 'r')
origFile=csv.reader(of,delimiter=';')

# Open output data file    
od = open('outdata.csv','w')
outData=csv.writer(od,delimiter=';')


# Open output file containing rows with errors
f = open('fejl.csv', 'w')
fejl=csv.writer(f,delimiter=';')

r = open('removed.csv', 'w')
removed=csv.writer(r,delimiter=';')
      
##### PROGRAM #####
for row in origFile:
    
    if rowNum == 0:
        header = row
        outData.writerow(header)
        fejl.writerow(header)
        rowNum += 1
        continue
    
    # If first real row (not header), row equals lastPass
    if rowNum == 1:
        lastPass = row
        outData.writerow(row)
        rowNum += 1
        continue
    
    # Splitting row into named variables for easier transparency
    dato_start, dato_slut, antal_pers, antal_overnat, leje, lejrtype, antal_dage, kundegrp, gns_temp, gns_vind, piller_start, piller_slut, piller_ttl, el_start, el_slut, el_ttl, vand_start, vand_slut, vand_ttl, total_omk, total_omk_dag = row
    
    try: 
        datetime.datetime.strptime(dato_start, '%d-%m-%Y')
    
    except:    
        removed.writerow(row)
        continue
       
    # Splitting lastPass row into named variables for easier transparency
    lp_dato_start, lp_dato_slut, lp_antal_pers, lp_antal_overnat, lp_leje, lp_lejrtype, lp_antal_dage, lp_kundegrp, lp_gns_temp, lp_gns_vind, lp_piller_start, lp_piller_slut, lp_piller_ttl, lp_el_start, lp_el_slut, lp_el_ttl, lp_vand_start, lp_vand_slut, lp_vand_ttl, lp_total_omk, lp_total_omk_dag = lastPass
      
    startDato = time.mktime(time.strptime(dato_start, "%d-%m-%Y"))
    lpStartDato = time.mktime(time.strptime(lp_dato_start, "%d-%m-%Y"))
    slutDato = time.mktime(time.strptime(dato_slut, "%d-%m-%Y"))
    lpSlutDato = time.mktime(time.strptime(lp_dato_slut, "%d-%m-%Y"))
        
    if startDato < lpStartDato or slutDato < lpSlutDato:
        print startDato, lpStartDato, slutDato, lpSlutDato
        continue
        
    if startDato > slutDato or el_start > lp_el_slut or vand_start > lp_vand_slut: #If el_start is bigger than el_slut (lastPass), this is internal use of el 
        newRow[0] = lp_dato_slut
        newRow[1] = dato_start
        newRow[13] = lp_el_slut 
        newRow[14] = el_start
        newRow[16] = lp_vand_slut 
        newRow[17] = vand_start
        newRow_changed = 1
       
    if el_start < lp_el_slut: # If el_start is smaller than el_slut (lastPass), this is an error
        fejl.writerow(lastPass)
        fejl.writerow(row)
        continue
          
    if vand_start < lp_vand_slut: # If el_start is smaller than el_slut (lastPass), this is an error. 
        fejl.writerow(lastPass)
        fejl.writerow(row)
        continue
    
    if newRow_changed == 1:
        newRow[5] = "intern"
        outData.writerow(newRow)
        newRow = [0] * 20
        newRow_changed = 0
        
    rowNum += 1             
    outData.writerow(row)
    lastPass = row
      
###### FILES CLOSE ######       
of.close()
od.close()
f.close()
r.close()

# Open output data file    
id = open('outdata.csv','r')
inData=csv.reader(id,delimiter=';')

# Open output data file    
od = open('outdata_calc.csv','w')
outData=csv.writer(od,delimiter=';')

for row in inData:
    dato_start, dato_slut, antal_pers, antal_overnat, leje, lejrtype, antal_dage, kundegrp, gns_temp, gns_vind, piller_start, piller_slut, piller_ttl, el_start, el_slut, el_ttl, vand_start, vand_slut, vand_ttl, total_omk, total_omk_dag = row
    
    elForbrug = int(el_slut) - int(el_start)
    el_ttl = elForbrug
    
    vandForbug = int(vand_slut) - int(vand_start)
    vand_ttl = vandForbug
    
    outData.writerow(row)
    
id.close()
od.close()
    

print("END OF PROGRAM")     
    
        