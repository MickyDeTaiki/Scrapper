from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import os
import time
import re

# Liste der PLZ-Werte
plz_list = ["01097", "01099", "01108", "01109", "01127", "01129", "01139"]

def scrape_website(website):
    print("Launching Chrome browser...")

    chrome_driver_path = "./chromedriver.exe"
    options = ChromeOptions()
    driver = Chrome(service=Service(chrome_driver_path), options=options)

    driver.get(website)
    time.sleep(3)
    print("Navigated to the website")

    for plz in plz_list:  # Schleife über alle PLZ-Werte
        print(f"Scraping for PLZ: {plz}")

        # Eingabefeld für PLZ finden und Wert setzen
        plz_input = driver.find_element(By.ID, "searchForm:txtPostal")
        plz_input.clear()  # Falls bereits ein Wert vorhanden ist, wird er gelöscht
        plz_input.send_keys(plz)  # PLZ-Wert eingeben
        time.sleep(2)  # Kurze Wartezeit für Stabilität

        # Click the "Suche starten!" button
        driver.find_element(By.ID, "searchForm:cmdSearch").click()
        time.sleep(3)

        # Wait for results to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "resultCard"))
        )

        # Initialize the output file
        file_path = "entries.xlsx"
        if os.path.exists(file_path):
            os.remove(file_path)  # Ensure a fresh file is created

        # Scrape until no more pages
        current_page = 1
        offset = 0  # Keeps track of the cumulative offset across pages

        while True:
            print(f"Scraping page {current_page} for PLZ {plz}...")

            # Handle potential modal blocking
            try:
                modal = driver.find_element(By.ID, "resultForm:exceptionDialogViewExpired:exceptionDialogViewExpired_modal")
                driver.execute_script("arguments[0].click();", modal)
                print("Closed blocking modal.")
                time.sleep(2)
            except:
                pass  # No modal found; continue scraping

            # Process each result card on the page
            result_cards = driver.find_elements(By.CLASS_NAME, "resultCard")
            for local_index in range(len(result_cards)):
                global_index = local_index + offset  # Adjust the index for cumulative offset
                try:
                    # Click on the "Info" link of the result card
                    info_button = driver.find_element(By.ID, f"resultForm:dlResultList:{global_index}:j_idt209")
                    driver.execute_script("arguments[0].scrollIntoView(true);", info_button)
                    info_button.click()
                    print(f"Clicked on Info link of result card {global_index + 1}")

                    # Wait for the "Detailansicht" to open
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "ResultDetailDialog_title"))
                    )

                    # Extract the HTML of the "Detailansicht"
                    detail_html = driver.page_source
                    entry = extract_details_from_detailansicht(detail_html)
                    print(f"Extracted details for result card {global_index + 1}")

                    # Append the entry to the Excel file
                    append_to_excel(entry, file_path)

                    # Close the "Detailansicht"
                    close_button = driver.find_element(By.XPATH, "//div[@id='ResultDetailDialog']//a[contains(@class, 'ui-dialog-titlebar-close')]")
                    close_button.click()

                    # Wait for the modal to close before proceeding
                    WebDriverWait(driver, 10).until(
                        EC.invisibility_of_element_located((By.ID, "ResultDetailDialog_title"))
                    )
                    print(f"Closed Detailansicht for result card {global_index + 1}")
                    time.sleep(2)  # Ensure the dialog closes
                except Exception as e:
                    print(f"Error processing result card {global_index + 1}: {str(e)}")

            # Update the offset to account for the current page's result cards
            offset += len(result_cards)

            # Click the "Next Page" button
            try:
                next_button = driver.find_element(By.XPATH, "//a[contains(@class, 'ui-paginator-next')]")
                if "ui-state-disabled" in next_button.get_attribute("class"):
                    print(f"No more pages available for PLZ {plz}.")
                    break  # **Beende die Schleife für diese PLZ und starte mit der nächsten**
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                next_button.click()
                print(f"Navigated to page {current_page + 1} for PLZ {plz}")

                # Introduce a 60-second wait after the 13th page
                if current_page == 13:
                    print("60-second wait before scraping the 14th page...")
                    time.sleep(60)  # Wait for 60 seconds

                current_page += 1

                # Wait for the next page to load
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "resultCard"))
                )
                time.sleep(3)  # Ensure the page fully loads
            except Exception as e:
                print(f"Error navigating to the next page for PLZ {plz}: {str(e)}")
                break

    driver.quit()
    print("Browser closed.")

# Extract details from the "Detailansicht" page
def extract_details_from_detailansicht(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    data = {}

    # Helper function to extract text safely
    def safe_extract(element):
        return element.get_text(strip=True) if element else "N/A"

    # Extract relevant fields
    data["Anrede"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt266:textEntry"))
    data["Berufsbezeichnung"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt273:textEntry"))
    data["Fachantwaltsbezeichnung(en)"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt282:textEntry"))
    if "Fachantwaltsbezeichnung(en)" in data:
        data["Fachantwaltsbezeichnung(en)"] = re.sub(r'(\w)([A-Z])', r'\1, \2', data["Fachantwaltsbezeichnung(en)"])
    data["Interesse an Pflichtverteidigungen"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt289:textEntry"))
    data["Vorname, Name"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt307:textEntry"))
    data["Datum der Zulassung"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt321:textEntry"))
    data["Datum der ersten Zulassung"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt330:textEntry"))
    data["Kammerzugehörigkeit"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt339:textEntry"))
    data["Name der Kanzlei"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt346:textEntry"))
    address_div = soup.find(id="resultDetailForm:tabPersonal:j_idt353:textEntry")
    data["Anschrift der Kanzlei"] = " ".join([div.get_text(strip=True) for div in address_div.find_all("div", class_="cssColResultDetailTextLine")]) if address_div else "N/A"
    data["Telefon"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt368:textEntry"))
    data["Mobilfunknummer"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt375:textEntry"))
    data["Telefax"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt382:textEntry"))
    data["E-Mail"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt389:textEntry"))
    data["Internetadresse"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt397:textEntry"))
    data["beA SAFE-ID"] = safe_extract(soup.find(id="resultDetailForm:tabPersonal:j_idt404:textEntry"))

    return data

# Append a single entry to an Excel file
def append_to_excel(entry, file_path):
    df = pd.DataFrame([entry])
    if not os.path.exists(file_path):
        df.to_excel(file_path, index=False)
    else:
        existing_df = pd.read_excel(file_path)
        updated_df = pd.concat([existing_df, df], ignore_index=True)
        updated_df.to_excel(file_path, index=False)
    print(f"Entry appended to {file_path}")

def process_excel_file(file_path):
    print("Processing Excel file to split names and titles...")
    df = pd.read_excel(file_path)

    # Definition der Listen für Titel A und Titel B
    titles_a = [
        "Dr.", "Dr. jur.", "Dr.jur.", "Dr. iur.", "Dr. rer. nat.", "Dr. rer. nat. Dipl.-Phys.",
        "Dr. rer. pol.", "Dr. jur. utr.", "Dr. Prof.", "Prof.", "Prof. Dr.", "Prof. Dr. iur.",
        "Prof. Dr. jur.", "Maître en Droit", "Mag. jur. Dr. jur.", "Dipl.-Kffr.", "Dipl.-Ing.",
        "Dipl.-Ing (FH).", "Dipl.-Finanzwirt (FH)", "Dipl.-Jur.", "JUDr. (Bratislava)",
        "Dipl.-Biol.", "Master of Laws", "M.M. Master in Mediation", "Wirtschaftsjurist (Univ.)",
        "Dipl.-Psych.", "LL.M.Eur.", "B.Sc.", "Dipl.-Verwaltungswirt (FH)", "Dipl.-Betriebswirt (FH)",
        "Wirtschaftsjur. (Uni Bayreuth)", "Dipl.-Jur. (Univ.)", "MLE", "LL.M.",
        "D.E.S.U. (Strasbourg)", "Univ.-Prof.", "RA", "FA", "Me", "Avv.", "Dott.", "Abg.",
        "Adv.", "Mr.", "QC", "KC", "SC", "Sir", "Dame", "OTM", "VT", "Adw.", "Radca prawny",
        "RP", "JUDr.", "Mgr.", "Δικ.", "M.A.", "Ass. Mag.iur.", "Dipl. BWin, Dipl. Jur.", "Ph.D. (UIBE)",
        "Dipl. Jur.", "JR Dr.", "Mag. iur.", "Dr.jur.,Dipl.-Betriebswirt(FH)", "Dr. (IT)", "Wirtschaftsjur. (Uni Bayreuth)",
        "Dipl.-Betriebsw.", "Maître en droit (Paris)", "M.A. phil.", "Ass. Jur.", "Dr. iur. utr.", "Ass. iur., Dipl.- jur. (Univ.)",
        "Dr. Dr.", "Dipl.-Ing. (FH)", "Dipl.Mag.-Jur.", "LL.M.oec.", "Europajurist", "Dipl. iur, Ass. Iur", "Ass. iur., Dipl.-iur.",
        "DEA", "Dipl.-Kfm.", "Ass. iur., Dipl.- jur. (Univ.)", "Prof. (em.) Dr.", "Ass. Dipl.-Jur. (Univ.)", "Dipl.-jur.", "Ass. iur., Dipl.- jur. (Univ.)",
        "Prof. Dr. Dr.", "Ass. iur., Dipl.- jur. (Univ.)", "Ass. iur., Dipl.-iur.", "Dipl.-Jur. Univ.", "Dipl.-Kfm.", "Ass. iur.",
        "Dipl.-Ing. (FH)","Dipl.-Jur. Univ.", "MBA", "Ass. jur.", "Dipl. iur, Ass. Iur", "Assessor jur.", "Dipl. iur, Ass. iur"    
             
    ]

    titles_b = [
        "LL.M.", "LL.M. IP Law", "LL.M.Eur.", "LL.M. (University of London)", "LL.M. (University of San Diego)", 
        "LL.M. (University of the Pacific)", "LL.M. (University of Houston)", "LL.M. (University of Essex)", 
        "LL.M. (University of Auckland)", "LL.M. (Victoria University of Wellington)", "LL.M (Maastricht)", 
        "LL.M. (Düsseldorf)", "LL.M. (Cape Town)", "LL.B.", "M.A.", "LL.M. M.A.", "LL.M.oec.", "LL.M. Lic.droit Maitrise en Droit", 
        "Lic.droit", "MLE", "M. mel.", "Mag. Jur.", "MBA", "MGlobL", "LL.M.Eur. Integ.", "Maître en droit", "D.E.S.", "MM", 
        "LL.M. (VUW)", "LL.M. (UQ)", "LL.M (Columbia)", "LL.M. (Christchurch)", "LL.M. (USA)", "LL.M. (Bristol)", "B.A.", 
        "LL.M. (King's College, London)", "M.B.A.", "M.C.L.", "LL.M. (Wirtschaftsrecht)", "M.L.E.", "LL.M. in European Law (Paris II - Assas)", 
        "Attorney-at-law (New York)", "J.D.", "Dr. jur.", "Ph.D.", "Dr. en droit", "Dott.", "Dr. hab.", "Dr. iur.", "DEA", "Máster en Derecho", 
        "Mag. iur.", "Mr.", "RA", "FA", "Avv.", "Abg.", "Adw.", "Radca Prawny", "Adv.", "QC/KC", "Barrister", "Solicitor", "Avocat(e)", "Advokat",
        "Attorney at Law", "LL.M. IP", "LL.M. Gewerblicher Rechtsschutz, Máster de Estudios Europeos", "Ass.Jur., Dipl.-iur. (Univ.)", "LL.M. oec.",
        "Attorney-at-law (New York), LL.M.", "DES (Liege- Belgium)", "Licenciado en Derecho, LL.B.,LL.M.", "M.B.L.", "J.S.M",
        "JD (KU)", "EMBA", "LL.M. Taxation", "Maître en droit (Paris)", "M.R.F", "Lawyer", "Licencié en droit", "Maître en Droit",
        "LL.M. (Harvard)", "LL.M (UK)", "LL.M.(Glasgow)", "LL.M. (New York)", "LL.M (Georgetown)", "LL.M. (Univ. Edinburgh)",
        "LL.M. Gewerblicher Rechtsschutz", "LL.M. (Miami)", "LL.M. (Cardiff)", "LL.M. (Berkeley)", "M.A. (London)",
        "LL.M. (London)", "Maitre en droit", "LL.M. (UCT)", "Lic. en droit", "LL.M. Computer Law (London)", "LL.M. (Auckland)",
        "LL.M.(Glasgow)", "LL.M. (Medienrecht)", "LL.M. (Sydney)", "LL.M. (UCT)", "LL.M. (Univ. of Aberdeen)", "LL.M. (Sydney)",
        "LL.M (Georgetown)", "L.L.M.Eur.", "M.C.J.", "LL.M. (Cardiff)", "J.S.M.", "Dipl.-Jur.", "LL.M (University of San Diego)",
        "LL.M. Eur.", "LL.M. (Columbia)", "LL.M (Edinburgh)"      
         
    ]

    def split_name(name):
        # Hilfsfunktionen zur genauen Titelermittlung
        def find_title(title_list, name, position="start"):
            matched_titles = []
            for title in title_list:
                pattern = r'^' + re.escape(title) + r'(\s|$)' if position == "start" else r'(\s|$)' + re.escape(title) + r'$'
                if re.search(pattern, name):
                    matched_titles.append(title)
            # Sortierung zur Sicherstellung der Reihenfolge, falls mehrere Titel
            matched_titles.sort(key=len, reverse=True)
            return matched_titles[0] if matched_titles else "Kein"

        # Titel A am Anfang des Namens
        title_a = find_title(titles_a, name, position="start")
        if title_a != "Kein":
            name = re.sub(r'^' + re.escape(title_a) + r'\s*', "", name)

        # Titel B am Ende des Namens
        title_b = find_title(titles_b, name, position="end")
        if title_b != "Kein":
            name = re.sub(r'\s*' + re.escape(title_b) + r'$', "", name)

        # Bereinigung des verbleibenden Namens
        parts = name.strip().split()
        first_name = parts[0] if len(parts) >= 1 else ""
        last_name = " ".join(parts[1:]) if len(parts) > 1 else ""

        return title_a, first_name, last_name, title_b

    # Anwendung der Spaltung auf die Spalte "Vorname, Name"
    if "Vorname, Name" in df.columns:
        title_first_last = df["Vorname, Name"].apply(lambda x: pd.Series(split_name(x), index=["Titel A", "Vorname", "Nachname", "Titel B"]))
        # Einfügen der neuen Spalten direkt nach "Vorname, Name"
        df = pd.concat([df.iloc[:, :1], title_first_last, df.iloc[:, 1:]], axis=1)
        # Entfernen der Originalspalte "Vorname, Name"
        df.drop(columns=["Vorname, Name"], inplace=True)
    else:
        print("Spalte 'Vorname, Name' nicht gefunden. Überspringe Verarbeitung.")

    # Formatierung der Spalte "Fachantwaltsbezeichnung(en)"
    if "Fachantwaltsbezeichnung(en)" in df.columns:
        df["Fachantwaltsbezeichnung(en)"] = df["Fachantwaltsbezeichnung(en)"].str.replace(r'(\w)([A-Z])', r'\1, \2', regex=True)

    # Speichern der bearbeiteten Datei
    df.to_excel(file_path, index=False)
    print("Excel file processed and updated successfully.")