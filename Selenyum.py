import win32com.client as win32
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as pg
import datetime

cred = "my_creds"

driver = webdriver.Chrome()
driver.get("http://mywebsite.com")
driver.find_element(By.NAME, "login").send_keys(cred)
driver.find_element(By.NAME, "password").send_keys(cred)
driver.find_element(By.CSS_SELECTOR, '.bouton.ClassSubmit').click() #https://stackoverflow.com/questions/21350605/python-selenium-click-on-button
driver.get("http://mywebsite/subpage.php")
driver.find_element(By.XPATH, '//input[@onclick="appliquer_tableau(\'Semaine\');"]').click()
driver.find_element(By.XPATH, '//input[@onclick="$(\'#printSheet\').val(\'Semaine\');$(\'#enregistrement\').click();"]').click()
sleep(4)
loc_download = pg.locateOnScreen(r"C:\my\path\to\image\download.png", confidence=0.7, grayscale=True)#ici l'url complete où est le screen shot
pg.click(loc_download)
sleep(1)
loc_save = pg.locateOnScreen(r"C:\my\path\to\image\enregistrer.png", confidence=0.7)#a changer
pg.click(loc_save)
sleep(1)

nom = "Prenom Nom"
fonction = "Métier" #les 3 a changer
num = "Numéro de tel "
todayz_date = datetime.date.today()
week = datetime.date(todayz_date.year, todayz_date.month, todayz_date.day).isocalendar().week

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'prenom.nom'
mail.Subject = f'Horaires semaine {week}'
mail.HTMLBody =f"""
    Bonjour <Prénom de a personne>,<br>
    Je mets ci-joint mes horaires de la semaine<br>
    Merci 😊<br>
    <br>
    <br>
        <div style="color:#C00000;font-family:Calibri (Corps);font-size:13px;"><strong>{nom}</strong>
    <br>
    nom du service</div>
        <div style="color:black;font-family:Calibri (Corps);font-size:13px;font-style:italic;">{fonction}</div>
    <br>
        <div style="color:#C00000;font-family:Calibri (Corps);font-size:13px;">{num}</div>  
        <div style="text-decoration:underline;font-family:Calibri (Corps);font-size:13px; color:black;"><strong>entreprise</strong></div>
    <br>
        <div style="color:#7F7F7F; font-style:italic;font-family:Calibri (Corps);font-size:11px;">Réduisons notre signature pour réduire notre empreinte</div>
    """

# To attach a file to the email (optional):
attachment  = f"C:\\path\\to\\downloaded\\file\\NOM-{week}.{todayz_date.year}.pdf" # a changer
mail.Attachments.Add(attachment)

mail.Send()
driver.quit()