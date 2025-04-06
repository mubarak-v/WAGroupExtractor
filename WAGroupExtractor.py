from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys
import pandas as pd
from colorama import init, Fore, Style
import time

def Ani_function(porpuse):
    init()
    print(Fore.GREEN + Style.BRIGHT + "Initializing " + porpuse +"...")
    time.sleep(3)
    print(Fore.RED + "Accessing secure "+porpuse + "... 🔓")



# Setup Chrome
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
input("📷 Scan the QR code, then press ENTER here when logged in...")

# Select group chat
Ani_function("Whatsapp")
group_name = input("Enter the WhatsApp group name: ").strip()

try:
    sidebar_chat = driver.find_element(By.XPATH, f'//span[@title="{group_name}"]')
    sidebar_chat.click()
    print(f"✅ Opened group: {group_name}")
except Exception as e:
    print(f'❌ Failed to open group chat\nError: {e}')
    driver.quit()
    exit()

# Open group info
Ani_function("Find Whatsapp Group ")
try:
    menu_buttons = driver.find_elements(By.XPATH, '//span[@data-icon="menu"]')
    if len(menu_buttons) >= 2:
        menu_buttons[1].click()
        time.sleep(1)
    else:
        print("❌ Couldn’t find group menu button")
except Exception as e:
    print(f"❌ Step 3 Failed: Couldn't click menu\nError: {e}")

try:
    group_info_button = driver.find_element(By.XPATH, '//div[@aria-label="Group info"]')
    group_info_button.click()
    time.sleep(2)
except Exception as e:
    print(f'❌ Couldn’t click "Group info"\nError: {e}')

# Click "View all" if present
time.sleep(1)
try:
    view_all_button = driver.find_element(By.XPATH, '//div[contains(text(), "View all")]')
    view_all_button.click()
    print('✅ Clicked "View all" to show all members')
    time.sleep(2)
except Exception:
    pass  # Continue even if "View all" not found

# Scroll and extract phone numbers
all_phone_numbers = set()
scroll_step = 500       # fast scroll
scroll_delay = 0.3      # short delay
batch_size = 8
same_scroll_count = 0
max_same_scroll = 6

def extract_numbers():
    spans = driver.find_elements(By.XPATH, '//span[@title]')
    fresh_numbers = set()

    for s in spans:
        title = s.get_attribute("title").strip()
        if (title.startswith('+') or title.replace(" ", "").isdigit()) and title not in all_phone_numbers:
            fresh_numbers.add(title)

    all_phone_numbers.update(fresh_numbers)

    if fresh_numbers:
        for i, number in enumerate(fresh_numbers, start=len(all_phone_numbers) - len(fresh_numbers) + 1):
            print(f"{i}. {number}")

    return list(fresh_numbers)

try:
    scroll_box = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.x1n2onr6.x1n2onr6.xyw6214.x78zum5.x1r8uery.x1iyjqo2.xdt5ytf.x6ikm8r.x1odjw0f.x1hc1fzr.x1tkvqr7'))
    )

    last_height = 0

    while True:
        new_batch = []

        while len(new_batch) < batch_size:
            driver.execute_script("arguments[0].scrollBy(0, arguments[1]);", scroll_box, scroll_step)
            time.sleep(scroll_delay)

            found_numbers = extract_numbers()
            for num in found_numbers:
                if num not in new_batch:
                    new_batch.append(num)

            current_height = driver.execute_script('return arguments[0].scrollHeight', scroll_box)
            if current_height == last_height:
                same_scroll_count += 1
            else:
                same_scroll_count = 0
                last_height = current_height

            if same_scroll_count >= max_same_scroll:
                break

        if not new_batch:
            print("\n🚫 No more new phone numbers found.")
            choice = input("🔄 Press [1] to scroll more, [2] to save and exit: ").strip()
            if choice == "1":
                same_scroll_count = 0
                continue
            elif choice == "2":
                break
            else:
                print("❗ Invalid option, saving by default.")
                break

except Exception as e:
    print(f"❌ Failed during scroll or extraction\nError: {e}")

# Save to Excel
contact_rows = []
for i, number in enumerate(all_phone_numbers, start=1):
    contact_rows.append({
            "First Name": f"smazio {i}",
            "Middle Name": "",
            "Last Name": "",
            "Phonetic First Name": "",
            "Phonetic Middle Name": "",
            "Phonetic Last Name": "",
            "Name Prefix": "",
            "Name Suffix": "",
            "Nickname": "",
            "File As": "",
            "Organization Name": "",
            "Organization Title": "",
            "Organization Department": "",
            "Birthday": "",
            "Notes": "",
            "Photo": "",
            "Labels": "* myContacts",
            "Phone 1 - Label": "Mobile",
            "Phone 1 - Value": number
        })

df = pd.DataFrame(contact_rows)
file_name = f"{group_name.replace(' ', '_')}_contacts.xlsx"
df.to_excel(file_name, index=False)

print(f"\n✅ Saved {len(all_phone_numbers)} contacts to '{file_name}'")

time.sleep(3)
driver.quit()
