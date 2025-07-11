from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import re
from collections import Counter
from datetime import datetime
import json
import ollama
import pyautogui

print("""
===================================
IMPORTANT: Do NOT close the Chrome browser window while this script is running.
If you close the browser, the script will lose connection and you must restart it.
===================================
""")

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


def select_board(driver):
    """Allow the user to select a retrospective board from the dropdown."""
    import time
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains
    print("\nLocating board selector...")
    try:
        # Wait for the board selector to be present
        board_selector = driver.find_element(By.CSS_SELECTOR, 'div.selector-button')
        board_selector.click()
        time.sleep(1)
        # Find all board options in the dropdown
        options = driver.find_elements(By.CSS_SELECTOR, 'div.ms-ContextualMenu-itemText')
        if not options:
            print("No board options found. Are you logged in and is the board loaded?")
            return None
        print("\nAvailable boards:")
        for idx, opt in enumerate(options, 1):
            print(f"{idx}. {opt.text}")
        choice = input("Select a board by number: ").strip()
        try:
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(options):
                options[choice_idx].click()
                print(f"Selected board: {options[choice_idx].text}")
                time.sleep(2)
                return options[choice_idx].text
            else:
                print("Invalid selection.")
                return None
        except Exception:
            print("Invalid input.")
            return None
    except Exception as e:
        print(f"Error selecting board: {e}")
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

    # NEW: Let user select the board
    select_board(driver)

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

    print("\nDone. You may now continue using the script. Do NOT close the browser window until you are finished.")
    # driver.quit()  # Removed to keep browser open


def add_new_feedback(column_title, feedback_text):
    global driver
    # Do NOT re-initialize or quit the driver; just use the existing session
    try:
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
            return f"Could not access the board iframe: {e}"
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
                            'button[aria-label=\"Add new feedback\"]')
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
                            return "Could not find save button"
                    except Exception as e:
                        print(f"Could not add feedback: {e}")
                        return f"Error adding feedback: {e}"
            except Exception:
                continue
        return f"Column '{column_title}' not found."
    except Exception as e:
        # Detect Selenium session loss and print a user-friendly message
        if 'Max retries exceeded with url' in str(e) or 'Failed to establish a new connection' in str(e):
            print("\n[ERROR] The connection to the browser was lost. This usually happens if the browser window was closed. Please restart the script to continue.\n")
            return "Selenium session lost. Please restart the script."
        print(f"Error in add_new_feedback: {e}")
        return f"Error accessing the board: {e}"


def setup_ollama_model():
    """Setup and pull the Ollama model for analysis."""
    try:
        # Pull a lightweight model suitable for text analysis
        print("Setting up Ollama model (this may take a few minutes)...")
        ollama.pull(
            'deepseek-coder:6.7b')  # Using smaller model for faster responses
        print("Ollama model ready!")
        return True
    except Exception as e:
        print(f"Error setting up Ollama: {e}")
        return False


def analyze_feedback_with_ollama(cards, question):
    """Use Ollama to analyze feedback cards and answer questions intelligently."""
    try:
        # Check if this is a request to add new feedback
        if any(phrase in question.lower() for phrase in ['add feedback', 'create feedback', 'new feedback', 'add card', 'create card']):
            return handle_ai_feedback_creation(question)

        # Enhanced: If question asks for cards with a vote range (e.g., 3-5 votes, between 2 and 4 votes)
        import re
        range_match = re.search(r'(?:between|from)?\s*(\d+)\s*(?:-|to|and)\s*(\d+)\s*vote', question.lower())
        if range_match:
            vmin = int(range_match.group(1))
            vmax = int(range_match.group(2))
            filtered = [c for c in cards if vmin <= c.get('votes', 0) <= vmax]
            if not filtered:
                return f"No feedback cards with votes between {vmin} and {vmax} found."
            result = f"Feedback cards with votes between {vmin} and {vmax} (Board | Tab | Column | Votes | Text):\n"
            for c in filtered:
                result += f"- {c.get('board', '?')} | {c.get('tab', '?')} | {c.get('column', '?')} | Votes: {c.get('votes', 0)} | {c['text'][:100]}\n"
            return result
        # Existing: If question asks for cards with N votes
        vote_match = re.search(r'(?:with|of|=|at least)?\s*(\d+)\s*vote', question.lower())
        if vote_match and not range_match:
            vote_count = int(vote_match.group(1))
            filtered = [c for c in cards if c.get('votes', 0) == vote_count]
            if not filtered:
                return f"No feedback cards with exactly {vote_count} vote(s) found."
            result = f"Feedback cards with exactly {vote_count} vote(s) (Board | Tab | Column | Votes | Text):\n"
            for c in filtered:
                result += f"- {c.get('board', '?')} | {c.get('tab', '?')} | {c.get('column', '?')} | Votes: {c.get('votes', 0)} | {c['text'][:100]}\n"
            return result
        # Fallback: at least 1 vote
        if 'at least 1 vote' in question.lower() or 'with votes' in question.lower() or 'votes' in question.lower():
            filtered = [c for c in cards if c.get('votes', 0) >= 1]
            if not filtered:
                return "No feedback cards with at least 1 vote found."
            result = "Feedback cards with at least 1 vote (Board | Tab | Column | Votes | Text):\n"
            for c in filtered:
                result += f"- {c.get('board', '?')} | {c.get('tab', '?')} | {c.get('column', '?')} | Votes: {c.get('votes', 0)} | {c['text'][:100]}\n"
            return result

        # Prepare context from all feedback cards
        context = "Feedback Data from Team Retrospectives:\n\n"
        for i, card in enumerate(cards, 1):
            date_info = f" (Date: {card['date']})" if card.get('date') else ""
            votes_info = f" (Votes: {card['votes']})" if card.get('votes', 0) > 0 else ""
            context += f"{i}. [{card.get('board', '?')}] [{card.get('tab', '?')}] [{card['column']}] {card['text']}{date_info}{votes_info}\n"

        # Create the prompt for Ollama
        prompt = f"""Based on the following team retrospective feedback data, please answer the user's question.\n\n{context}\nUser Question: {question}\n\nPlease provide a comprehensive answer based on the feedback data above. If the question asks for specific feedback items, include relevant quotes. If asking for analysis or trends, provide insights based on the data."""

        # Query Ollama
        response = ollama.chat(model='deepseek-coder:6.7b',
                               messages=[{
                                   'role': 'user',
                                   'content': prompt
                               }])

        return f"🤖 AI Analysis:\n{response['message']['content']}"

    except Exception as e:
        print(f"Ollama error: {e}")
        # Fallback to original analysis
        return analyze_feedback_question_fallback(cards, question)


def analyze_feedback_question(cards, question):
    """Main analysis function that tries Ollama first, then falls back to rule-based analysis."""
    # Try Ollama first
    try:
        return analyze_feedback_with_ollama(cards, question)
    except:
        # Fallback to original rule-based system
        return analyze_feedback_question_fallback(cards, question)


def handle_ai_feedback_creation(question):
    """Handle AI requests to create new feedback cards."""
    try:
        # Use Ollama to parse the feedback creation request
        prompt = f"""The user wants to add new feedback to a retrospective board. Please analyze their request and extract:
1. The column/category where the feedback should go (like \"What went well\", \"What could be improved\", \"Action items\", etc.)
2. The feedback text content

User request: {question}

Please respond in this exact format:
COLUMN: [column name]
FEEDBACK: [feedback text]

If you cannot determine the column or feedback from the request, ask for clarification."""

        response = ollama.chat(model='deepseek-coder:6.7b',
                               messages=[{
                                   'role': 'user',
                                   'content': prompt
                               }])

        ai_response = response['message']['content']

        # Parse the AI response
        lines = ai_response.split('\n')
        column = None
        feedback = None

        for line in lines:
            if line.startswith('COLUMN:'):
                column = line.replace('COLUMN:', '').strip()
            elif line.startswith('FEEDBACK:'):
                feedback = line.replace('FEEDBACK:', '').strip()

        if column and feedback:
            # Confirm with user before adding
            print(f"\n🤖 AI wants to add feedback:")
            print(f"   Column: {column}")
            print(f"   Feedback: {feedback}")

            confirm = input("\nProceed with adding this feedback? (y/n): ").strip().lower()
            if confirm == 'y' or confirm == 'yes':
                # Call the existing add_new_feedback function and check result
                result = add_new_feedback(column, feedback)
                if result is None:
                    return f"✅ Feedback added successfully to '{column}' column!"
                elif isinstance(result, str) and (result.lower().startswith('error') or 'could not' in result.lower()):
                    return f"❌ Error adding feedback: {result}"
                else:
                    return f"✅ Feedback added successfully to '{column}' column!"
            else:
                return "❌ Feedback creation cancelled."
        else:
            return f"🤖 I need more information. Please specify:\n{ai_response}"

    except Exception as e:
        return f"❌ Error creating feedback: {e}\n\nPlease try again with a clearer request like: 'Add feedback to What went well: The team worked great together'"

def analyze_feedback_question_fallback(cards, question):
    """Fallback rule-based analysis (original function)."""
    question_lower = question.lower()

    # Handle special commands
    if question_lower == 'summary':
        return generate_summary(cards)
    elif question_lower == 'trends':
        return analyze_trends(cards)
    elif question_lower == 'stats':
        return generate_statistics(cards)
    elif any(phrase in question_lower for phrase in ['add feedback', 'create feedback', 'new feedback', 'add card', 'create card']):
        return handle_ai_feedback_creation(question)

    # Question classification and routing
    if any(word in question_lower for word in ['how many', 'count', 'number']):
        return handle_count_question(cards, question_lower)
    elif any(word in question_lower
             for word in ['most', 'common', 'frequent']):
        return handle_frequency_question(cards, question_lower)
    elif any(word in question_lower for word in ['when', 'date', 'time']):
        return handle_date_question(cards, question_lower)
    elif any(word in question_lower
             for word in ['what went well', 'positive', 'good']):
        return handle_positive_feedback(cards, question_lower)
    elif any(word in question_lower
             for word in ['improve', 'problem', 'issue', 'bad']):
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
    stop_words = {
        'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with',
        'by', 'is', 'was', 'are', 'were', 'be', 'been', 'have', 'has', 'had',
        'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may',
        'might', 'can', 'a', 'an'
    }
    filtered_words = [
        word for word in words if len(word) > 3 and word not in stop_words
    ]
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
        col_avg_length = sum(len(card['text'])
                             for card in col_cards) / len(col_cards)
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
        stop_words = {
            'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of',
            'with', 'by'
        }
        filtered_words = [
            word for word in words if len(word) > 3 and word not in stop_words
        ]
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
        matching_cards = [
            card for card in dated_cards
            if target_date.lower() in card['date'].lower()
        ]
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
    positive_columns = [
        'What went well', 'Positives', 'Good', 'Wins', 'Successes'
    ]
    positive_cards = [
        card for card in cards if any(col.lower() in card['column'].lower()
                                      for col in positive_columns)
    ]

    if not positive_cards:
        return "No clearly positive feedback columns found."

    result = f"Found {len(positive_cards)} positive feedback items:\n"
    for card in positive_cards[:5]:  # Show first 5
        result += f"  • [{card['column']}] {card['text'][:80]}...\n"

    return result


def handle_improvement_feedback(cards, question):
    """Handle questions about areas for improvement."""
    improvement_columns = [
        'What could be improved', 'Issues', 'Problems', 'Concerns',
        'Improvements'
    ]
    improvement_cards = [
        card for card in cards if any(col.lower() in card['column'].lower()
                                      for col in improvement_columns)
    ]

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
        highest_scored_cards = [(score, card) for score, card in scored_cards
                                if score == highest_score]
    else:
        highest_scored_cards = []

    result = f"🔍 Search Results for '{question}' ({len(highest_scored_cards)} highest matches):\n"
    result += "=" * 50 + "\n"

    for i, (score, card) in enumerate(highest_scored_cards, 1):
        date_str = f" | 📅 {card['date']}" if card.get('date') else ''
        result += f"\n#{i} [{card['column']}] (Score: {score})\n"
        result += f"    {card['text'][:120]}...{date_str}\n"

    return result


def create_training_context(cards):
    """Create a comprehensive training context for Ollama."""
    context = "TEAM RETROSPECTIVE FEEDBACK DATABASE\n"
    context += "=" * 50 + "\n\n"

    # Group by columns
    columns = {}
    for card in cards:
        col = card['column']
        if col not in columns:
            columns[col] = []
        columns[col].append(card)

    for col_name, col_cards in columns.items():
        context += f"\n{col_name.upper()} ({len(col_cards)} items):\n"
        context += "-" * 30 + "\n"
        for i, card in enumerate(col_cards, 1):
            date_info = f" [Date: {card['date']}]" if card.get('date') else ""
            context += f"{i}. {card['text']}{date_info}\n"

    # Add summary statistics
    context += f"\n\nSUMMARY STATISTICS:\n"
    context += f"- Total feedback items: {len(cards)}\n"
    context += f"- Number of categories: {len(columns)}\n"

    # Add key insights
    all_text = ' '.join(card['text'].lower() for card in cards)
    words = re.findall(r'\b\w+\b', all_text)
    stop_words = {
        'the', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with',
        'by', 'is', 'was', 'are', 'were'
    }
    filtered_words = [
        word for word in words if len(word) > 3 and word not in stop_words
    ]
    common_words = Counter(filtered_words).most_common(5)

    context += f"- Most mentioned topics: {', '.join([f'{word}({count})' for word, count in common_words])}\n"

    return context


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
    # Get current board name
    try:
        current_board = driver.find_element(By.CSS_SELECTOR, 'div.feedback-board-container-header div.selector-button span').text.strip()
    except Exception:
        current_board = "Unknown"
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
                # Extract vote count
                try:
                    vote_elem = card.find_element(By.CSS_SELECTOR, 'span.vote-count, span.votes, span[class*="vote"]')
                    votes = int(re.sub(r'\D', '', vote_elem.text)) if vote_elem.text.strip() else 0
                except Exception:
                    votes = 0
                cards_list.append({
                    'board': current_board,
                    'column': col_title,
                    'text': feedback,
                    'date': date_found,
                    'votes': votes
                })
        except Exception:
            continue
    # driver.quit()  # Removed to keep browser open
    return cards_list


def scrape_all_boards_and_train_ai():
    """Scrape feedback from all boards and feed to AI."""
    driver.get(RETRO_URL)
    print("Launching browser to scrape all boards (you may need to log in manually)...")
    time.sleep(30)
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

    try:
        scraped_boards = set()
        all_cards = []
        while True:
            # Get the currently selected board
            try:
                current_board = driver.find_element(By.CSS_SELECTOR, 'div.feedback-board-container-header div.selector-button span').text.strip()
            except Exception:
                current_board = None
            if not current_board or current_board in scraped_boards:
                break
            print(f"Scraping currently selected board: {current_board}")
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
                        # Extract vote count
                        try:
                            vote_elem = card.find_element(By.CSS_SELECTOR, 'span.vote-count, span.votes, span[class*="vote"]')
                            votes = int(re.sub(r'\D', '', vote_elem.text)) if vote_elem.text.strip() else 0
                        except Exception:
                            votes = 0
                        all_cards.append({
                            'board': current_board,
                            'column': col_title,
                            'text': feedback,
                            'date': date_found,
                            'votes': votes
                        })
                except Exception:
                    continue
            scraped_boards.add(current_board)
            # Open dropdown and select the first available board not yet scraped
            board_selector = driver.find_element(By.CSS_SELECTOR, 'div.feedback-board-container-header div.selector-button')
            board_selector.click()
            time.sleep(1)
            board_items = driver.find_elements(By.CSS_SELECTOR, 'div.selector-list-item')
            next_board_found = False
            for item in board_items:
                try:
                    board_name = item.find_element(By.CSS_SELECTOR, 'div.selector-list-item-text').text.strip()
                    if board_name not in scraped_boards:
                        item.click()
                        print(f"Selected board: {board_name}")
                        time.sleep(3)
                        next_board_found = True
                        break
                except Exception:
                    continue
            if not next_board_found:
                break
        print(f"\nScraped {len(all_cards)} feedback cards from all boards.")
        # Feed to AI
        print("\nTraining AI with all feedback data from all boards...")
        training_context = ""
        boards = {}
        for card in all_cards:
            board = card['board']
            if board not in boards:
                boards[board] = []
            boards[board].append(card)
        for board, cards in boards.items():
            training_context += f"\n\nBOARD: {board}\n" + "="*40 + "\n"
            for c in cards:
                date_info = f" [Date: {c['date']}]" if c.get('date') else ""
                votes_info = f" [Votes: {c['votes']}]" if c.get('votes', 0) > 0 else ""
                training_context += f"[{c.get('column', '?')}] {c['text']}{date_info}{votes_info}\n"
        # Start interactive chat session
        print("\n" + "="*50)
        print("🤖 AI CHAT SESSION READY")
        print("="*50)
        print("The AI has been trained with feedback from all boards.")
        print("You can now ask questions, request analysis, or add new feedback.")
        print("\nAvailable commands:")
        print("- Ask any question about the feedback (AI-powered)")
        print("- Add new feedback: 'Add feedback to [column]: [text]'")
        print("- 'summary' - Get overall summary")
        print("- 'trends' - Analyze trends and patterns")
        print("- 'stats' - Get statistics")
        print("- 'exit' - Return to main menu")
        while True:
            question = input("\n🤖 Ask your question: ").strip()
            if question.lower() == 'exit':
                break
            answer = analyze_feedback_with_ollama(all_cards, question)
            print(f"\n{answer}")
    except Exception as e:
        print(f"Error scraping all boards: {e}")


def click_tab_by_name(tab_name):
    """Click a tab (Collect, Vote, Act, etc.) by its visible name using Selenium."""
    try:
        tab_buttons = driver.find_elements(By.CSS_SELECTOR, 'p.stage-text')
        for btn in tab_buttons:
            if btn.text.strip().lower() == tab_name.lower():
                btn.click()
                time.sleep(3)
                return True
        print(f"Tab '{tab_name}' not found!")
        return False
    except Exception as e:
        print(f"Error clicking tab '{tab_name}': {e}")
        return False


def scrape_all_boards_all_tabs_and_train_ai():
    """Scrape all boards, all tabs (Collect, Vote, Act), and train AI."""
    driver.get(RETRO_URL)
    print("Launching browser to scrape all boards and all tabs (Collect, Vote, Act)...")
    time.sleep(30)
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
    all_cards = []
    scraped_boards = set()
    while True:
        # Get current board name
        try:
            current_board = driver.find_element(By.CSS_SELECTOR, 'div.feedback-board-container-header div.selector-button span').text.strip()
        except Exception:
            current_board = None
        if not current_board or current_board in scraped_boards:
            break
        print(f"Scraping board: {current_board}")
        for tab_name in ['Collect', 'Vote', 'Act']:
            print(f"Switching to tab: {tab_name}")
            click_tab_by_name(tab_name)
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
                        # Extract vote count
                        try:
                            vote_elem = card.find_element(By.CSS_SELECTOR, 'span.vote-count, span.votes, span[class*="vote"]')
                            votes = int(re.sub(r'\D', '', vote_elem.text)) if vote_elem.text.strip() else 0
                        except Exception:
                            votes = 0
                        all_cards.append({
                            'board': current_board,
                            'tab': tab_name,
                            'column': col_title,
                            'text': feedback,
                            'date': date_found,
                            'votes': votes
                        })
                except Exception:
                    continue
        scraped_boards.add(current_board)
        # Open dropdown and select the next board
        try:
            board_selector = driver.find_element(By.CSS_SELECTOR, 'div.feedback-board-container-header div.selector-button')
            board_selector.click()
            time.sleep(1)
            board_items = driver.find_elements(By.CSS_SELECTOR, 'div.selector-list-item')
            next_board_found = False
            for item in board_items:
                try:
                    board_name = item.find_element(By.CSS_SELECTOR, 'div.selector-list-item-text').text.strip()
                    if board_name not in scraped_boards:
                        item.click()
                        print(f"Selected board: {board_name}")
                        time.sleep(3)
                        next_board_found = True
                        break
                except Exception:
                    continue
            if not next_board_found:
                break
        except Exception:
            break
    print(f"\nScraped {len(all_cards)} feedback cards from all boards and tabs.")
    # Feed to AI
    print("\nTraining AI with all feedback data from all boards and tabs...")
    training_context = ""
    boards = {}
    for card in all_cards:
        board = card['board']
        if board not in boards:
            boards[board] = []
        boards[board].append(card)
    for board, cards in boards.items():
        training_context += f"\n\nBOARD: {board}\n" + "="*40 + "\n"
        for c in cards:
            date_info = f" [Date: {c['date']}]" if c.get('date') else ""
            votes_info = f" [Votes: {c['votes']}]" if c.get('votes', 0) > 0 else ""
            training_context += f"[{c.get('column', '?')}] {c['text']}{date_info}{votes_info}\n"
    try:
        ollama.chat(
            model='deepseek-coder:6.7b',
            messages=[{
                'role': 'user',
                'content': f"Please learn and remember this team retrospective data for future questions: {training_context}"
            }])
        print("✅ AI trained with all feedback data from all boards and tabs!")
        # Start interactive chat session
        print("\n" + "="*50)
        print("🤖 AI CHAT SESSION READY")
        print("="*50)
        print("The AI has been trained with feedback from all boards and tabs.")
        print("You can now ask questions, request analysis, or add new feedback.")
        print("\nAvailable commands:")
        print("- Ask any question about the feedback (AI-powered)")
        print("- Add new feedback: 'Add feedback to [column]: [text]'")
        print("- 'summary' - Get overall summary")
        print("- 'trends' - Analyze trends and patterns")
        print("- 'stats' - Get statistics")
        print("- 'exit' - Return to main menu")
        while True:
            question = input("\n🤖 Ask your question: ").strip()
            if question.lower() == 'exit':
                break
            answer = analyze_feedback_with_ollama(all_cards, question)
            print(f"\n{answer}")
    except Exception as e:
        print(f"⚠️ Training failed: {e}")


def scrape_feedback_from_tab(tab_name):
    """Switch to the specified tab (Collect, Vote, Act) and scrape feedback cards."""
    tab_name = tab_name.capitalize()
    print(f"Switching to '{tab_name}' tab...")
    try:
        # Find the tab button by its text and click it
        tab_buttons = driver.find_elements(By.CSS_SELECTOR, 'div[role="tab"]')
        found = False
        for btn in tab_buttons:
            if btn.text.strip().lower() == tab_name.lower():
                btn.click()
                found = True
                time.sleep(3)  # Wait for tab content to load
                break
        if not found:
            print(f"Tab '{tab_name}' not found!")
            return []
        # Now scrape feedback cards as usual
        cards = []
        columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
        for col in columns:
            try:
                col_title = col.find_element(By.CSS_SELECTOR, '.feedback-column-title h2').text
            except Exception:
                col_title = "Unknown"
            try:
                card_container = col.find_element(By.CSS_SELECTOR, '.feedback-column-content')
                feedback_cards = card_container.find_elements(By.CSS_SELECTOR, 'div.feedbackItem')
                for card in feedback_cards:
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
                    cards.append({
                        'tab': tab_name,
                        'column': col_title,
                        'text': feedback,
                        'date': date_found
                    })
            except Exception:
                continue
        print(f"Scraped {len(cards)} cards from '{tab_name}' tab.")
        return cards
    except Exception as e:
        print(f"Error scraping '{tab_name}' tab: {e}")
        return []


def scrape_tabs_with_autogui(tab_coords):
    """
    tab_coords: list of (tab_name, x, y) tuples, e.g. [ ('Collect', 200, 150), ('Vote', 300, 150), ('Act', 400, 150) ]
    For each tab, click at (x, y) using pyautogui, then scrape feedback cards.
    """
    all_cards = []
    for tab_name, x, y in tab_coords:
        print(f"Clicking on '{tab_name}' tab at ({x}, {y})...")
        pyautogui.click(x, y)
        time.sleep(3)  # Wait for tab content to load
        # Scrape feedback cards as usual
        cards = []
        columns = driver.find_elements(By.CSS_SELECTOR, 'div.feedback-column')
        for col in columns:
            try:
                col_title = col.find_element(By.CSS_SELECTOR, '.feedback-column-title h2').text
            except Exception:
                col_title = "Unknown"
            try:
                card_container = col.find_element(By.CSS_SELECTOR, '.feedback-column-content')
                feedback_cards = card_container.find_elements(By.CSS_SELECTOR, 'div.feedbackItem')
                for card in feedback_cards:
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
                    cards.append({
                        'tab': tab_name,
                        'column': col_title,
                        'text': feedback,
                        'date': date_found
                    })
            except Exception:
                continue
        print(f"Scraped {len(cards)} cards from '{tab_name}' tab.")
        all_cards.extend(cards)
    print(f"\nTotal feedback cards scraped from all tabs: {len(all_cards)}")
    return all_cards


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
11. Scrape all boards and all tabs and train AI
""")
    option = input("Choose an option (1-11): ").strip()
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

        # Setup Ollama
        print("Initializing AI analysis system...")
        ollama_ready = setup_ollama_model()
        if ollama_ready:
            print(
                "✅ AI analysis ready! You can now ask complex questions about your feedback."
            )
        else:
            print(
                "⚠️ AI analysis unavailable, using fallback analysis system.")

        print("\nAvailable commands:")
        print("- Ask any question about the feedback (AI-powered)")
        print("- Add new feedback: 'Add feedback to [column]: [text]'")
        print("- 'summary' - Get overall summary")
        print("- 'trends' - Analyze trends and patterns")
        print("- 'stats' - Get statistics")
        print("- 'train' - Feed all data to AI for better context")
        print("- 'exit' - Return to menu")

        while True:
            question = input("\nAsk your question: ").strip()
            if question.lower() == 'exit':
                break
            elif question.lower() == 'train':
                print("Training AI with current feedback data...")
                # Create a comprehensive training prompt
                training_context = create_training_context(all_cards)
                try:
                    ollama.chat(
                        model='deepseek-coder:6.7b',
                        messages=[{
                            'role':
                            'user',
                            'content':
                            f"Please learn and remember this team retrospective data for future questions: {training_context}"
                        }])
                    print("✅ AI trained with current feedback data!")
                except:
                    print(
                        "⚠️ Training failed, but you can still ask questions.")
                continue

            answer = analyze_feedback_question(all_cards, question)
            print(f"\n{answer}")
    elif option == '11':
        scrape_all_boards_all_tabs_and_train_ai()
    else:
        print("Option not implemented in this script.")
