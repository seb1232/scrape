from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re
from collections import Counter
from datetime import datetime
import json

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
        r"(\d{4}-\d{2}-\d{2})"  # e.g., 2025-06-27
    ]
    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return None


def scrape_retrospective_cards():
    driver.get(RETRO_URL)
    print(
        "Launching browser to scrape feedback cards (you may need to log in manually)..."
    )
    time.sleep(30)  # Wait for manual login or page load
    login_if_needed(driver)
    time.sleep(5)  # Wait for board to render

    # Switch to the iframe containing the board
    try:
        iframe = driver.find_element(By.CSS_SELECTOR,
                                     'iframe.external-content--iframe')
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
            col_title = col.find_element(By.CSS_SELECTOR,
                                         '.feedback-column-title h2').text
        except Exception as e:
            print(f"Column {idx+1}: Could not find title. Error: {e}")
            col_title = f"Column {idx+1}"
        print(f"\n=== {col_title} ===")
        try:
            card_container = col.find_element(By.CSS_SELECTOR,
                                              '.feedback-column-content')
            cards = card_container.find_elements(By.CSS_SELECTOR,
                                                 'div.feedbackItem')
            print(f"  Found {len(cards)} cards in this column.")
            seen = set()
            for card in cards:
                feedback = None
                # Try textarea first
                try:
                    textarea = card.find_element(
                        By.CSS_SELECTOR, 'textarea.ms-TextField-field')
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
            print(
                f"Could not find card container in column {col_title}. Error: {e}"
            )
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
        iframe = driver.find_element(By.CSS_SELECTOR,
                                     'iframe.external-content--iframe')
        driver.switch_to.frame(iframe)
        print("Switched to iframe.")
    except Exception as e:
        print(f"Could not switch to iframe: {e}")
        driver.quit()
        return
    columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
    for col in columns:
        try:
            col_title = col.find_element(
                By.CSS_SELECTOR, '.feedback-column-title h2').text.strip()
            if col_title.lower() == column_title.lower():
                print(f"Found column: {col_title}")
                try:
                    add_btn = col.find_element(
                        By.CSS_SELECTOR,
                        'button[aria-label="Add new feedback"]')
                    add_btn.click()
                    time.sleep(1)
                    textarea = col.find_element(By.CSS_SELECTOR,
                                                'textarea.ms-TextField-field')
                    textarea.clear()
                    textarea.send_keys(feedback_text)
                    time.sleep(1)
                    # Try to find the save/submit button
                    try:
                        save_btn = col.find_element(
                            By.CSS_SELECTOR, 'button.ms-Button--primary')
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


def analyze_feedback_question(cards, question):
    """Analyze feedback cards and provide intelligent answers to user questions."""
    question_lower = question.lower()

    # Handle special commands
    if question_lower == 'summary':
        return generate_summary(cards)
    elif question_lower == 'trends':
        return analyze_trends(cards)
    elif question_lower == 'stats':
        return generate_statistics(cards)

    # Question classification and routing
    if any(word in question_lower for word in ['how many', 'count', 'number']):
        return handle_count_question(cards, question_lower)
    elif any(word in question_lower for word in ['most', 'common', 'frequent']):
        return handle_frequency_question(cards, question_lower)
    elif any(word in question_lower for word in ['when', 'date', 'time']):
        return handle_date_question(cards, question_lower)
    elif any(word in question_lower for word in ['what went well', 'positive', 'good']):
        return handle_positive_feedback(cards, question_lower)
    elif any(word in question_lower for word in ['improve', 'problem', 'issue', 'bad']):
        return handle_improvement_feedback(cards, question_lower)
    else:
        return handle_general_search(cards, question_lower)


def generate_summary(cards):
    """Generate an overall summary of all feedback."""
    if not cards:
        return "No feedback cards available."

    columns = {}
    for card in cards:
        col = card['column']
        if col not in columns:
            columns[col] = []
        columns[col].append(card['text'])

    summary = f"📊 FEEDBACK SUMMARY ({len(cards)} total items)\n"
    summary += "=" * 50 + "\n"

    for col, items in columns.items():
        summary += f"\n🔹 {col} ({len(items)} items):\n"
        # Show top 3 items or all if less than 3
        for i, item in enumerate(items[:3]):
            preview = item[:80] + "..." if len(item) > 80 else item
            summary += f"  • {preview}\n"
        if len(items) > 3:
            summary += f"  ... and {len(items) - 3} more items\n"

    return summary


def analyze_trends(cards):
    """Analyze trends and patterns in feedback."""
    if not cards:
        return "No feedback cards available for trend analysis."

    # Analyze by column
    column_counts = Counter(card['column'] for card in cards)

    # Find common keywords
    all_text = ' '.join(card['text'].lower() for card in cards)
    words = re.findall(r'\b\w+\b', all_text)
    # Filter out common words
    stop_words = {'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'is', 'was', 'are', 'were', 'be', 'been', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'can', 'a', 'an'}
    filtered_words = [word for word in words if len(word) > 3 and word not in stop_words]
    common_words = Counter(filtered_words).most_common(10)

    trends = "📈 TRENDS & PATTERNS\n"
    trends += "=" * 30 + "\n"
    trends += "\n🔹 Feedback Distribution:\n"
    for col, count in column_counts.most_common():
        percentage = (count / len(cards)) * 100
        trends += f"  • {col}: {count} items ({percentage:.1f}%)\n"

    trends += "\n🔹 Most Common Topics:\n"
    for word, count in common_words:
        trends += f"  • '{word}': mentioned {count} times\n"

    return trends


def generate_statistics(cards):
    """Generate detailed statistics about feedback."""
    if not cards:
        return "No feedback cards available for statistics."

    # Basic stats
    total_cards = len(cards)
    columns = set(card['column'] for card in cards)
    avg_length = sum(len(card['text']) for card in cards) / total_cards

    # Date analysis
    dated_cards = [card for card in cards if card.get('date')]

    stats = "📊 DETAILED STATISTICS\n"
    stats += "=" * 30 + "\n"
    stats += f"Total feedback items: {total_cards}\n"
    stats += f"Number of columns: {len(columns)}\n"
    stats += f"Average feedback length: {avg_length:.1f} characters\n"
    stats += f"Items with dates: {len(dated_cards)}\n"

    # Column breakdown
    stats += "\n🔹 Column Breakdown:\n"
    for col in columns:
        col_cards = [card for card in cards if card['column'] == col]
        col_avg_length = sum(len(card['text']) for card in col_cards) / len(col_cards)
        stats += f"  • {col}: {len(col_cards)} items (avg: {col_avg_length:.1f} chars)\n"

    return stats


def handle_count_question(cards, question):
    """Handle questions about counting items."""
    if 'column' in question or 'category' in question:
        columns = Counter(card['column'] for card in cards)
        result = f"Found {len(columns)} columns with feedback:\n"
        for col, count in columns.most_common():
            result += f"  • {col}: {count} items\n"
        return result
    else:
        return f"Total feedback items: {len(cards)}"


def handle_frequency_question(cards, question):
    """Handle questions about most common/frequent items."""
    if 'word' in question or 'topic' in question:
        all_text = ' '.join(card['text'].lower() for card in cards)
        words = re.findall(r'\b\w+\b', all_text)
        stop_words = {'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by'}
        filtered_words = [word for word in words if len(word) > 3 and word not in stop_words]
        common_words = Counter(filtered_words).most_common(5)

        result = "Most frequently mentioned topics:\n"
        for word, count in common_words:
            result += f"  • '{word}': {count} times\n"
        return result
    else:
        columns = Counter(card['column'] for card in cards)
        most_common = columns.most_common(1)[0]
        return f"Most common feedback category: '{most_common[0]}' with {most_common[1]} items"


def handle_date_question(cards, question):
    """Handle questions about dates and timing."""
    dated_cards = [card for card in cards if card.get('date')]
    if not dated_cards:
        return "No feedback items with dates found."

    # Extract specific date from question if present
    date_patterns = [
        r"([A-Z][a-z]+ \d{1,2},? \d{4})",  # e.g., June 27, 2025 or June 27 2025
        r"(\d{4}-\d{2}-\d{2})",  # e.g., 2025-06-27
        r"(\d{1,2}/\d{1,2}/\d{4})"  # e.g., 06/27/2025
    ]
    
    target_date = None
    for pattern in date_patterns:
        match = re.search(pattern, question)
        if match:
            target_date = match.group(1)
            break
    
    if target_date:
        # Filter cards that match the specific date
        matching_cards = [card for card in dated_cards if target_date.lower() in card['date'].lower()]
        if matching_cards:
            result = f"📅 Feedback for {target_date} ({len(matching_cards)} items):\n"
            result += "=" * 50 + "\n"
            for card in matching_cards:
                result += f"\n🔹 [{card['column']}]\n"
                result += f"   Date: {card['date']}\n"
                result += f"   Text: {card['text']}\n"
            return result
        else:
            return f"No feedback found for {target_date}."
    
    # General date overview
    result = f"Found {len(dated_cards)} feedback items with dates:\n"
    # Group by date
    date_groups = {}
    for card in dated_cards:
        date_key = card['date']
        if date_key not in date_groups:
            date_groups[date_key] = []
        date_groups[date_key].append(card)
    
    # Sort dates and show most recent first
    sorted_dates = sorted(date_groups.keys(), reverse=True)
    for date_key in sorted_dates[:5]:  # Show last 5 dates
        cards_for_date = date_groups[date_key]
        result += f"\n📅 {date_key} ({len(cards_for_date)} items):\n"
        for card in cards_for_date:
            result += f"  • [{card['column']}] {card['text'][:80]}...\n"

    return result


def handle_positive_feedback(cards, question):
    """Handle questions about positive feedback."""
    positive_columns = ['What went well', 'Positives', 'Good', 'Wins', 'Successes']
    positive_cards = [card for card in cards if any(col.lower() in card['column'].lower() for col in positive_columns)]

    if not positive_cards:
        return "No clearly positive feedback columns found."

    result = f"Found {len(positive_cards)} positive feedback items:\n"
    for card in positive_cards[:5]:  # Show first 5
        result += f"  • [{card['column']}] {card['text'][:80]}...\n"

    return result


def handle_improvement_feedback(cards, question):
    """Handle questions about areas for improvement."""
    improvement_columns = ['What could be improved', 'Issues', 'Problems', 'Concerns', 'Improvements']
    improvement_cards = [card for card in cards if any(col.lower() in card['column'].lower() for col in improvement_columns)]

    if not improvement_cards:
        return "No clearly improvement-focused feedback columns found."

    result = f"Found {len(improvement_cards)} improvement feedback items:\n"
    for card in improvement_cards[:5]:  # Show first 5
        result += f"  • [{card['column']}] {card['text'][:80]}...\n"

    return result


def handle_general_search(cards, question):
    """Handle general search queries with better matching."""
    # Check if this is a date-specific question
    date_patterns = [
        r"([A-Z][a-z]+ \d{1,2},? \d{4})",  # e.g., June 27, 2025
        r"(\d{4}-\d{2}-\d{2})",  # e.g., 2025-06-27
        r"(\d{1,2}/\d{1,2}/\d{4})"  # e.g., 06/27/2025
    ]
    
    target_date = None
    for pattern in date_patterns:
        match = re.search(pattern, question)
        if match:
            target_date = match.group(1)
            break
    
    if target_date:
        return handle_date_question(cards, question)
    
    # Extract key terms from question
    question_words = re.findall(r'\b\w+\b', question.lower())
    question_words = [word for word in question_words if len(word) > 2]

    # Score cards based on relevance
    scored_cards = []
    for card in cards:
        score = 0
        text_lower = card['text'].lower()

        # Exact phrase match (highest score)
        if question.lower() in text_lower:
            score += 10

        # Individual word matches
        for word in question_words:
            if word in text_lower:
                score += 1

        # Enhanced date matching
        if card.get('date'):
            card_date_lower = card['date'].lower()
            # Exact date match gets high score
            for word in question_words:
                if word in card_date_lower:
                    score += 5

        if score > 0:
            scored_cards.append((score, card))

    # Sort by relevance score
    scored_cards.sort(key=lambda x: x[0], reverse=True)

    if not scored_cards:
        return f"No relevant feedback found for: '{question}'"

    # Get only the highest scoring items
    if scored_cards:
        highest_score = scored_cards[0][0]
        highest_scored_cards = [(score, card) for score, card in scored_cards if score == highest_score]
    else:
        highest_scored_cards = []

    result = f"🔍 Search Results for '{question}' ({len(highest_scored_cards)} highest matches):\n"
    result += "=" * 50 + "\n"
    
    for i, (score, card) in enumerate(highest_scored_cards, 1):
        date_str = f" | 📅 {card['date']}" if card.get('date') else ''
        result += f"\n#{i} [{card['column']}] (Score: {score})\n"
        result += f"    {card['text'][:120]}...{date_str}\n"

    return result


def get_all_feedback_cards():
    driver.get(RETRO_URL)
    time.sleep(30)
    login_if_needed(driver)
    time.sleep(5)
    try:
        iframe = driver.find_element(By.CSS_SELECTOR,
                                     'iframe.external-content--iframe')
        driver.switch_to.frame(iframe)
    except Exception:
        driver.quit()
        return []
    cards_list = []
    columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
    for col in columns:
        try:
            col_title = col.find_element(By.CSS_SELECTOR,
                                         '.feedback-column-title h2').text
        except Exception:
            col_title = "Unknown"
        try:
            card_container = col.find_element(By.CSS_SELECTOR,
                                              '.feedback-column-content')
            cards = card_container.find_elements(By.CSS_SELECTOR,
                                                 'div.feedbackItem')
            for card in cards:
                feedback = None
                try:
                    textarea = card.find_element(
                        By.CSS_SELECTOR, 'textarea.ms-TextField-field')
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
                cards_list.append({
                    'column': col_title,
                    'text': feedback,
                    'date': date_found
                })
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
        print("Scraping all feedback cards for intelligent Q&A...")
        all_cards = get_all_feedback_cards()
        if not all_cards:
            print("No feedback cards found.")

        print(f"Loaded {len(all_cards)} feedback cards.")
        print("Available commands:")
        print("- Ask questions about the feedback")
        print("- 'summary' - Get overall summary")
        print("- 'trends' - Analyze trends and patterns") 
        print("- 'stats' - Get statistics")
        print("- 'exit' - Return to menu")

        while True:
            question = input("\nAsk your question: ").strip()
            if question.lower() == 'exit':
                break

            answer = analyze_feedback_question(all_cards, question)
            print(f"\n{answer}")
    else:
        print("Option not implemented in this script.")