from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import re
from team_abbreviations import team_dict

def sanitize_filename(text, max_length=50):
    """Sanitize text to be safe for filenames and limit its length."""
    text = re.sub(r'[\\/*?:"<>|]', "", text)  # Remove potentially problematic characters
    return text[:max_length]  # Truncate to max_length

index = 0

# Specify your ChromeDriver path
chrome_driver_path = "chromedriver.exe"

# Setup ChromeDriver
service = Service(executable_path=chrome_driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

try:
    # Navigate to the webpage
    driver.get("https://www.rib.gg/series/75248")
    WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "table.MuiTable-root")))

    # Create a new Excel workbook
    wb = Workbook()
    ws_dict = {}

    # Creating a sheet for match information
    wb.worksheets[0].title = "Match Data"

    map_elements = driver.find_elements(By.CSS_SELECTOR, "div.MuiBox-root.css-dgd30b")
    map_names = [re.sub(r'\s*VOD\s*', '', element.text).strip() for element in map_elements if element.text]
    for map_name in map_names:
        ws = wb.create_sheet(title=map_name)
        ws_dict[map_name] = ws
        headers = ['Team', 'Player', 'Agent', "Rating", 'ACS', 'K', 'D', 'A', 'KD+/-', 'KD', 'ADR', 'FK', 'FD', 'FKFD+/-', 'Clutches', 'KAST%', 'HS%']
        ws.append(headers)
    
    try:
        # Common patterns for event names and match titles
        page_title = driver.title
        headings = driver.find_elements(By.CSS_SELECTOR, "h1, h2")
        event_name = sanitize_filename(headings[0].text if headings else 'Event Name Not Found')
        match_title = sanitize_filename(headings[1].text if len(headings) > 1 else 'Match Title Not Found')
    except IndexError:
        event_name = match_title = "Not Found"

    # Construct filenames
    excel_filename = f"{event_name}_{match_title}_stats.xlsx"

    # Filling out Match Data Page

    event_name = headings[0].text
    wb.worksheets[0].append([event_name])
    match_title = headings[1].text
    wb.worksheets[0].append([match_title])

    teams = match_title.split(" vs ")
    i = 0
    for team in teams:
        abbreviation = team_dict.get(team[i], "Unknown Team")
        
        i = i+1
    
    match_details = driver.find_elements(By.CSS_SELECTOR, ".MuiBox-root.css-nwxytv")

    # Extract text from each element
    element_raw = [detail.text for detail in match_details]
    element = [line for item in element_raw for line in item.split('\n')]


    # Print or process the extracted texts
    team1_mapcnt = int(element[0])
    team2_mapcnt = int(element[2])
    wb.worksheets[0].append([str(element[0] + " " + element[1] + " " + element[2])])
    wb.worksheets[0].append([element[3]])
    wb.worksheets[0].append([element[4]])
    wb.worksheets[0].append(['\n'])
    

    pick_ban = driver.find_elements(By.CSS_SELECTOR, ".MuiTypography-root.MuiTypography-main.css-j00j6r")
    pick_ban_raw = [select.text for select in pick_ban]
    map_veto = [line for item in pick_ban_raw for line in item.split('; ')]

    for veto in map_veto:
        wb.worksheets[0].append([veto])

    
    # Find the table by its CSS class and extract rows

    # Process each data row
    wait = WebDriverWait(driver, 20)
    for map_name in map_names:
        map_name_to_click = map_names[index]
        print(f"extracting content for {map_name_to_click} ...")
        clickable_element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(@class, 'css-dgd30b') and contains(text(), '{map_name_to_click}')]")))
        clickable_element.click()

        table = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "table.MuiTable-root")))
        # Now interact with the table
        rows = table.find_elements(By.CSS_SELECTOR, "tr.MuiTableRow-root")
        for row in rows:
            if "header" in row.get_attribute('class'):
                th_element = row.find_element(By.CSS_SELECTOR, "a")
                current_team = th_element.text if th_element else 'Unknown Team'
            else:
                cols = row.find_elements(By.CSS_SELECTOR, "td")
                if len(cols) > 10:  # Ensure there are enough columns
                    data = [current_team]
                    data.append(cols[0].text.strip())  # Player name

                    # Extract agent names using 'alt' attribute from images within the agent column
                    agent = 'No Agent Info'
                    if len(cols) > 1:
                        images = cols[0].find_elements(By.CSS_SELECTOR, "img[alt]")  # This ensures only images with 'alt' attributes are considered
                        agent_names = set(img.get_attribute('alt') for img in images if "/assets/agents/" in img.get_attribute('src'))
                        agent = ' '.join(agent_names)

                    data.append(agent)

                    # Add ACS from the next column
                    rating = cols[1].text.strip()
                    data.append(rating)

                    acs = cols[2].text.strip()
                    data.append(acs)

                    # K, D, A are expected in a single cell as numbers separated by space
                    kda_text = cols[3].text.strip()
                    data.extend(kda_text.split())  # Split into K, D, A assuming three numbers are present

                    # Process the rest of the data
                    for col in cols[4:]:  # Starting after K, D, A
                        text = col.text.strip()
                        data.append(text)

                    wb.worksheets[index+1].append(data)
        index = index + 1
finally:
    # Close the browser
    driver.quit()
    wb.save(filename=excel_filename)
    print(f"Data extraction completed and saved to '{excel_filename}'.")

