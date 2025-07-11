from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re

# --- CONFIGURE THESE ---
RETRO_URL = "https://apollo.healthcare.siemens.com/tfs/IKM.TPC.Projects/syngo.net/_apps/hub/ms-devlabs.team-retrospectives.home#teamId=5e86f89b-b28c-4750-b5ff-fdc2a1630ec9&boardId=281cce01-1903-40cf-ba44-ee06af3d2da9"
USERNAME = "your_username"  # If needed for login automation
PASSWORD = "your_password"  # If needed for login automation

# --- SELENIUM SETUP ---
options = Options()
options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=options)

def login_if_needed(driver):
    # If your org uses SSO or a login page, automate login here
    # Example for Azure AD login (customize as needed):
    # driver.find_element(By.ID, 'i0116').send_keys(USERNAME)
    # driver.find_element(By.ID, 'idSIButton9').click()
    # time.sleep(2)
    # driver.find_element(By.ID, 'i0118').send_keys(PASSWORD)
    # driver.find_element(By.ID, 'idSIButton9').click()
    # time.sleep(2)
    pass

def extract_date_from_text(text):
    # Try to find a date in the format 'June 27, 2025' or '2025-06-27'
    date_patterns = [
        r"([A-Z][a-z]+ \d{1,2}, \d{4})",  # e.g., June 27, 2025
        r"(\d{4}-\d{2}-\d{2})"           # e.g., 2025-06-27
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return None

def scrape_retrospective_cards():
    driver.get(RETRO_URL)
    print("Launching browser to scrape feedback cards (you may need to log in manually)...")
    time.sleep(30)  # Wait for manual login or page load
    login_if_needed(driver)
    time.sleep(5)  # Wait for board to render

    # Switch to the iframe containing the board
    try:
        iframe = driver.find_element(By.CSS_SELECTOR, 'iframe.external-content--iframe')
        driver.switch_to.frame(iframe)
        print("Switched to iframe.")
    except Exception as e:
        print(f"Could not switch to iframe: {e}")
        driver.quit()
        return

    columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
    print(f"Found {len(columns)} columns.")
    for idx, col in enumerate(columns):
        try:
            col_title = col.find_element(By.CSS_SELECTOR, '.feedback-column-title h2').text
        except Exception as e:
            print(f"Column {idx+1}: Could not find title. Error: {e}")
            col_title = f"Column {idx+1}"
        print(f"\n=== {col_title} ===")
        try:
            card_container = col.find_element(By.CSS_SELECTOR, '.feedback-column-content')
            cards = card_container.find_elements(By.CSS_SELECTOR, 'div.feedbackItem')
            print(f"  Found {len(cards)} cards in this column.")
            seen = set()
            for card in cards:
                feedback = None
                # Try textarea first
                try:
                    textarea = card.find_element(By.CSS_SELECTOR, 'textarea.ms-TextField-field')
                    feedback = textarea.get_attribute('value').strip()
                except Exception:
                    pass
                # If not found or empty, try visible text
                if not feedback:
                    try:
                        feedback = card.text.strip()
                    except Exception:
                        continue
                if not feedback or len(feedback) < 5:
                    continue
                if feedback in seen:
                    continue
                seen.add(feedback)
                print(f"- {feedback}")
        except Exception as e:
            print(f"Could not find card container in column {col_title}. Error: {e}")
            continue

    print("\nDone. Close the browser window to exit.")
    time.sleep(30)
    driver.quit()

def add_new_feedback(column_title, feedback_text):
    driver.get(RETRO_URL)
    print(f"Opening board to add feedback to '{column_title}'...")
    time.sleep(30)  # Wait for manual login or page load
    login_if_needed(driver)
    time.sleep(5)
    try:
        iframe = driver.find_element(By.CSS_SELECTOR, 'iframe.external-content--iframe')
        driver.switch_to.frame(iframe)
        print("Switched to iframe.")
    except Exception as e:
        print(f"Could not switch to iframe: {e}")
        driver.quit()
        return
    columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
    for col in columns:
        try:
            col_title = col.find_element(By.CSS_SELECTOR, '.feedback-column-title h2').text.strip()
            if col_title.lower() == column_title.lower():
                print(f"Found column: {col_title}")
                try:
                    add_btn = col.find_element(By.CSS_SELECTOR, 'button[aria-label="Add new feedback"]')
                    add_btn.click()
                    time.sleep(1)
                    textarea = col.find_element(By.CSS_SELECTOR, 'textarea.ms-TextField-field')
                    textarea.clear()
                    textarea.send_keys(feedback_text)
                    time.sleep(1)
                    # Try to find the save/submit button
                    try:
                        save_btn = col.find_element(By.CSS_SELECTOR, 'button.ms-Button--primary')
                        save_btn.click()
                        print("Feedback added!")
                        time.sleep(3)
                        return
                    except Exception:
                        # Try clicking the first enabled button after textarea
                        buttons = col.find_elements(By.CSS_SELECTOR, 'button')
                        for btn in buttons:
                            if btn.is_enabled():
                                btn.click()
                                print("Feedback added (fallback button)!")
                                time.sleep(3)
                                return
                        print("Could not find a save/submit button to click.")
                        return
                except Exception as e:
                    print(f"Could not add feedback: {e}")
                    return
        except Exception:
            continue
    print(f"Column '{column_title}' not found.")
    driver.quit()

def get_all_feedback_cards():
    driver.get(RETRO_URL)
    time.sleep(30)
    login_if_needed(driver)
    time.sleep(5)
    try:
        iframe = driver.find_element(By.CSS_SELECTOR, 'iframe.external-content--iframe')
        driver.switch_to.frame(iframe)
    except Exception:
        driver.quit()
        return []
    cards_list = []
    columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
    for col in columns:
        try:
            col_title = col.find_element(By.CSS_SELECTOR, '.feedback-column-title h2').text
        except Exception:
            col_title = "Unknown"
        try:
            card_container = col.find_element(By.CSS_SELECTOR, '.feedback-column-content')
            cards = card_container.find_elements(By.CSS_SELECTOR, 'div.feedbackItem')
            for card in cards:
                feedback = None
                try:
                    textarea = card.find_element(By.CSS_SELECTOR, 'textarea.ms-TextField-field')
                    feedback = textarea.get_attribute('value').strip()
                except Exception:
                    pass
                if not feedback:
                    try:
                        feedback = card.text.strip()
                    except Exception:
                        continue
                if not feedback or len(feedback) < 5:
                    continue
                date_found = extract_date_from_text(feedback)
                cards_list.append({'column': col_title, 'text': feedback, 'date': date_found})
        except Exception:
            continue
    driver.quit()
    return cards_list

if __name__ == "__main__":
    print("""
===================================

Options:
1. List retrospective boards
2. View retrospective items
3. Create new retrospective item
4. Update retrospective item
5. Ask question about retrospectives
6. Exit
7. View feedback cards from Retrospectives board (extension)
8. Scrape feedback cards from Retrospectives board (web scraping)
9. Add new feedback card (web scraping)
10. Ask any question about retrospectives (FAQ search)
""")
    option = input("Choose an option (1-10): ").strip()
    if option == '8':
        scrape_retrospective_cards()
    elif option == '9':
        col = input("Enter column title (e.g., What went well): ").strip()
        text = input("Enter feedback text: ").strip()
        add_new_feedback(col, text)
    elif option == '10':
        print("Scraping all feedback cards for FAQ search...")
        all_cards = get_all_feedback_cards()
        print(f"Loaded {len(all_cards)} feedback cards. Type 'exit' to return to the menu.")
        while True:
            question = input("Ask your question: ").strip()
            if question.lower() == 'exit':
                break
            matches = [c for c in all_cards if question.lower() in c['text'].lower() or (c.get('date') and question.lower() in c['date'].lower())]
            if matches:
                print(f"Found {len(matches)} relevant feedback cards:")
                for m in matches:
                    date_str = f" | Date: {m['date']}" if m.get('date') else ''
                    print(f"[{m['column']}] {m['text']}{date_str}")
            else:
                print("No relevant feedback found.")
    else:
        print("Option not implemented in this script.")
