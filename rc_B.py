

# -*- coding: utf-8 -*-
"""
Created on Mon Apr 12 08:37:32 2021
"""
#Importation des packages
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlsxwriter
import csv
#from pyvirtualdisplay import Display
#Création d’une session chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--verbose')
#Initialisation des variables
h=0
RC=[]
i=0
j=0
ii=0
#display = Display(visible=0, size=(800, 600))
#display.start()
#Ouverture du fichier contenant la liste des fournisseurs 
with open(r'input_B.csv',newline='',errors='ignore') as csvfile2:
    reader2= csv.reader(csvfile2, quotechar='|')
    for row2 in reader2:
        RC.append(row2[0])
    print(row2)
workbook = xlsxwriter.Workbook('rslt_RC_B_new4.xlsx')
worksheet = workbook.add_worksheet("trouvé")
worksheet2 = workbook.add_worksheet("non_trouvé")
worksheet3 = workbook.add_worksheet("Associé")
worksheet4 = workbook.add_worksheet("Bilan")
first=1
i3=0
i4=0
j3=0
j4=0
for nrc in RC:
    try:
        browser = webdriver.Chrome("D:\chromedriver.exe")
        browser.get('https://sidjilcom.cnrc.dz/repertoire-des-commercants-detaille?p_auth=FT1jHkdF&p_p_id=recherchea_portlet_WAR_commercantdetaillee&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view&p_p_col_id=column-2&p_p_col_count=1&_recherchea_portlet_WAR_commercantdetaillee__spage=%2FRecherche%2FRechercheAction.do')
        browser.find_element_by_id("_58_login").send_keys("dr-risque")
        pwd=browser.find_element_by_id("_58_password")
        pwd.send_keys("0000")
        pwd.send_keys(Keys.ENTER)
        prs_physique=browser.find_element_by_xpath("//tbody/tr/td[2]/a")
        prs_physique.click()
        j=0
        i=i+1
        time.sleep(10)
        print(nrc)
        
        nrc2=nrc.replace(" ","")
        browser.find_element_by_id("nrc6").send_keys(nrc2[:3])
        browser.find_element_by_id("nrc7").send_keys(" ب")
        RC=browser.find_element_by_id("nrc8")
        RC.send_keys(nrc2[3::])
        recherche=browser.find_element_by_xpath("//div[@id='id_onglet1']/div/div/a")
        recherche.click()
        time.sleep(5)
        #Récolte des données et enregistrement sur un fichier excel 
        resultat=browser.find_element_by_xpath("//*[contains(@id,'header5')]/tbody//tr")
        data = [item.text for item in resultat.find_elements_by_xpath(".//*[self::td]")]
        print(data)
        numéro=data[0]
        Nom=data[1]
        Prénom=data[2]
        Commune=data[3]
        Etat=data[4]
        
        resultat2=browser.find_element_by_xpath("//*[contains(@id,'header5')]/tbody//tr/td[5]/img")
        resultat2.click()
        time.sleep(5)
        worksheet.write(i,j,nrc)
        j=j+1
        worksheet.write(i,j,Nom)
        j=j+1
        worksheet.write(i,j,Prénom)
        j=j+1
        worksheet.write(i,j,Commune)
        j=j+1
        j=j+1
        worksheet.write(i,j,Etat)
        j=j+1
        time.sleep(10)
        window_after = browser.window_handles[1]
        browser.switch_to.window(window_after)
        associe=browser.find_element_by_xpath("//*[contains(@id,'onglets')]/li[3]")
        associe.click()
        time.sleep(10)
        resultat_associe=browser.find_elements_by_xpath("//*[contains(@id,'header1')]/tbody/tr")
        for d in resultat_associe:
            data3 = [item.text for item in d.find_elements_by_xpath(".//*[self::td]")]
            i3=i3+1
            j3=0
            worksheet3.write(i3,j3,nrc)
            j3=j3+1
            worksheet3.write(i3,j3,Nom)
            j3=j3+1
            worksheet3.write(i3,j3,Prénom)
            j3=j3+1
            worksheet3.write(i3,j3,Commune)
            j3=j3+1
            j3=j3+1
            worksheet3.write(i3,j3,Etat)
            j3=j3+1
            worksheet3.write(i3,j3,data3[0])
            j3=j3+1
            worksheet3.write(i3,j3,data3[1])
            j3=j3+1
            worksheet3.write(i3,j3,data3[2])
            j3=j3+1
            worksheet3.write(i3,j3,data3[3])
            j3=j3+1
            worksheet3.write(i3,j3,data3[4])
            j3=j3+1
            worksheet3.write(i3,j3,data3[5])
            j3=j3+1
            worksheet3.write(i3,j3,data3[6])
            j3=j3+1
        window_after = browser.window_handles[0]
        browser.switch_to.window(window_after)
        boal=browser.find_element_by_xpath("//*[contains(@id,'header5')]/tbody//tr/td[7]/img")
        boal.click()
        time.sleep(10)
        fermer=browser.find_element_by_xpath("//*[@id='div_sec']/center/table/tbody/tr[1]/td[3]/a/u")
        fermer.click()
        bilan=browser.find_elements_by_xpath("//*[@id='header']/tbody/tr")
        taille=len(bilan)
        for t in range (taille-1,0,-1):
            i4=i4+1
            
            année=browser.find_element_by_xpath("//*[@id='header']/tbody/tr["+str(t)+"]/td[1]").text
            print(année)
            bilan2=browser.find_element_by_xpath("//*[contains(@id,'header')]/tbody/tr["+str(t)+"]/td[2]/img")
            bilan2.click()
            time.sleep(20)
            window_after = browser.window_handles[1]
            browser.switch_to.window(window_after)
            #id_onglet3
            recherche2=browser.find_element_by_name("#id_onglet3")
            recherche2.click()
            time.sleep(5)
            j4=0
            worksheet4.write(i4,j4,nrc)
            j4=j4+1
            stock=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[3]/div/div/table[1]/tbody/tr[19]/td[5]")
            stock2=stock.text
            worksheet4.write(i4,j4,année+ " :"+stock2)
            j4=j4+1
            creance=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[3]/div/div/table[1]/tbody/tr[20]/td[5]")
            creance2=creance.text
            worksheet4.write(i4,j4,année+ " :"+creance2)
            j4=j4+1
            clients=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[3]/div/div/table[1]/tbody/tr[21]/td[5]")
            clients2=clients.text
            worksheet4.write(i4,j4,année+ " :"+clients2)
            j4=j4+1
            autres=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[3]/div/div/table[1]/tbody/tr[22]/td[5]")
            autres2=autres.text
            worksheet4.write(i4,j4,année+ " :"+autres2)
            j4=j4+1
            trésorerie=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[3]/div/div/table[1]/tbody/tr[27]/td[5]")
            trésorerie2=trésorerie.text
            worksheet4.write(i4,j4,année+ " :"+trésorerie2)
            j4=j4+1
            #id_onglet3
            recherche2=browser.find_element_by_name("#id_onglet4")
            recherche2.click()
            time.sleep(5)
            emprunt=browser.find_element_by_xpath("//*[@id='header2']/tbody/tr[13]/td[3]")
            emprunt2=emprunt.text
            worksheet4.write(i4,j4,année+ " :"+emprunt2)
            j4=j4+1
            fournisseur=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[4]/div/div/table[1]/tbody/tr[19]/td[3]")
            fournisseur2=fournisseur.text
            worksheet4.write(i4,j4,année+ " :"+fournisseur2)
            j4=j4+1
            dettes=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[4]/div/div/table[1]/tbody/tr[22]/td[3]")
            dettes2=dettes.text
            worksheet4.write(i4,j4,année+ " :"+dettes2)
            j4=j4+1
            tréso=browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div[1]/div/div[1]/section/div/div/div/form/center/fieldset/div/div/div[4]/div/div/table[1]/tbody/tr[22]/td[3]")
            tréso2=tréso.text
            worksheet4.write(i4,j4,année+ " :"+tréso2)
            j4=j4+1
            
            recherche2=browser.find_element_by_name("#id_onglet5")
            recherche2.click()
            time.sleep(5)
            resultat5=browser.find_element_by_xpath("//*[contains(@id,'heade3r')]/tbody//tr/td[3]")
            CA=resultat5.text
            worksheet4.write(i4,j4,année+ " :"+CA)
            j4=j4+1
            production=browser.find_element_by_xpath("//*[@id='heade3r']/tbody/tr[2]/td[3]")
            production2=production.text
            worksheet4.write(i4,j4,année+ " :"+production2)
            j4=j4+1
            
            immo=browser.find_element_by_xpath("//*[@id='heade3r']/tbody/tr[3]/td[3]")
            immo2=immo.text
            worksheet4.write(i4,j4,année+ " :"+immo2)
            j4=j4+1
            
            achats=browser.find_element_by_xpath("//*[@id='heade3r']/tbody/tr[6]/td[3]")
            achats2=achats.text
            worksheet4.write(i4,j4,année+ " :"+achats2)
            j4=j4+1
            window_after = browser.window_handles[0]
            browser.switch_to.window(window_after)
    except Exception as e :
        print(e)
        worksheet2.write(ii,0,nrc)
        ii=ii+1
    browser.quit()
workbook.close()
#for ele in resultat:
 #   print(ele.text)
