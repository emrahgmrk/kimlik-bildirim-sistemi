from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import openpyxl

excel_path = 'veriler.xlsx'  # Excel dosyanızın yolunu buraya girin
wb = openpyxl.load_workbook(excel_path)
sheet = wb.active

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

driver.get('https://vatandas.jandarma.gov.tr/KBS_tesis_2/')

time.sleep(3)

input("Başlamak için F11 tuşuna basın ve ardından Enter'a basın...")

calisanlar_link = driver.find_element(By.XPATH, "//a[@href='frmTssPrsLst.aspx']")
calisanlar_link.click()

time.sleep(2)

yeni_ekle_button = driver.find_element(By.ID, '_ctl0_headercontent_btnYeni')
yeni_ekle_button.click()

# Excel'deki her satır için işlemleri yap
for row in sheet.iter_rows(min_row=2, values_only=True):
    if not row[0]:
        print("Boş A sütunu bulundu, program durduruluyor.")
        break

    print(f"Satır: {row}")

    tc_input = driver.find_element(By.ID, '_ctl0_Content_txtTC')
    tc_input.clear()
    tc_input.send_keys(row[0])

    ara_button = driver.find_element(By.ID, '_ctl0_Content_btnAra')
    ara_button.click()

    time.sleep(7)

    select_container = driver.find_element(By.ID, 'select2-_ctl0_Content_drpPerTur-container')
    select_container.click()

    select_option = driver.find_element(By.XPATH, "//li[text()='Sürekli Personel (SGK Bildirimi Yapılacak) (M6)']")
    select_option.click()

    grs_tar_input = driver.find_element(By.ID, '_ctl0_Content_txtGrsTar')
    grs_tar_input.click()  # Takvimi açmak için tıklayın
    time.sleep(1)  # Takvimin açılması için kısa bir bekleme
    driver.execute_script("arguments[0].value = '';", grs_tar_input)  # Varsayılan tarihi temizle
    grs_tar_input.send_keys(row[1])  # Excel'deki değeri gir
    grs_tar_input.send_keys("\n")  # Takvimi kapatmak için Enter tuşuna bas

    print(f"Giriş Tarihi: {row[1]}")

    gorevi_input = driver.find_element(By.ID, '_ctl0_Content_txtGorevi')
    gorevi_input.clear()
    print(f"Görevi: {row[2]}")
    gorevi_input.send_keys(row[2])

    # Brüt Maaş alanını bul ve Excel'deki değeri gir
    brut_maas_input = driver.find_element(By.ID, '_ctl0_Content_txtBrutMaas')
    brut_maas_input.clear()
    print(f"Brüt Maaş: {row[3]}")
    brut_maas_input.send_keys(row[3])

    sgk_no_input = driver.find_element(By.ID, '_ctl0_Content_txtSgkNo')
    sgk_no_input.clear()
    print(f"SGK No: {row[4]}")
    sgk_no_input.send_keys(row[4])

    adres_input = driver.find_element(By.ID, '_ctl0_Content_txtAdres')
    adres_input.clear()
    print(f"Adres: {row[5]}")
    adres_input.send_keys(row[5])

    barinma_adres_input = driver.find_element(By.ID, '_ctl0_Content_txtBarınmaAdres')
    barinma_adres_input.clear()
    print(f"Barınma Adresi: {row[6]}")
    barinma_adres_input.send_keys(row[6])

    kaydet_button = driver.find_element(By.ID, '_ctl0_Content_btnKaydet')
    kaydet_button.click()

    time.sleep(5)

    tc_input = driver.find_element(By.ID, '_ctl0_Content_txtTC')
    tc_input.clear()

    time.sleep(2)


