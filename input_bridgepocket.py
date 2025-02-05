from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Path ke ChromeDriver
CHROME_DRIVER_PATH = "/path/to/chromedriver"  # Sesuaikan dengan lokasi chromedriver

# Load data dari spreadsheet
df = pd.read_excel("data_tim_bridge.xlsx")  # Sesuaikan nama file Excel

# Setup browser
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(CHROME_DRIVER_PATH), options=options)
driver.get("https://bridgepocket.net/#/tournament/pps25/teams")

# Tunggu hingga halaman termuat
wait = WebDriverWait(driver, 10)

def add_team(team_name, members):
    # Klik tombol plus (+) untuk tambah tim
    plus_button = wait.until(EC.element_to_be_clickable((By.ID, "plus_team")))
    plus_button.click()

    # Isi nama tim
    team_name_input = wait.until(EC.presence_of_element_located((By.NAME, "team_name")))
    team_name_input.send_keys(team_name)

    # Isi anggota tim
    for i, member in enumerate(members):
        member_input = driver.find_element(By.NAME, f"member_{i+1}")
        member_input.send_keys(member)

    # Simpan tim
    save_button = driver.find_element(By.NAME, "save_team")
    save_button.click()
    time.sleep(1)  # Tunggu sejenak sebelum menambahkan tim berikutnya

# Loop untuk input data dari Excel
for _, row in df.iterrows():
    team_name = row['Team Name']
    members = [row[f'Member {i+1}'] for i in range(1, 5) if not pd.isna(row[f'Member {i+1}'])]
    add_team(team_name, members)

# Selesai
print("Data tim berhasil ditambahkan.")
driver.quit()
