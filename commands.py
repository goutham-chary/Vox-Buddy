from tasks import *
import spacy

nlp = spacy.load("en_core_web_sm")

def process_command_func(command):
    doc = nlp(command)

    if any(token.text.lower() in ["hello", "hi", "hai", "hay", "hey"] for token in doc):
        return greet()
    elif any(token.text.lower() in ["name", "your name", "who are you", "what is your name"] for token in doc):
        return "My name is Vox Buddy. How can i assist you today"
    
    elif any(token.text.lower() in ["thanks", "thank", "thank you"] for token in doc):
        return "You're welcome! If you have any more questions or if there's anything else I can help you with, feel free to ask."

    elif "location" in command.lower():
        return get_current_location()

    elif any(token.text.lower() == "calendar" for token in doc):
        return manage_calendar(command)
    
    elif any(token.text.lower() == "google" for token in doc):
        return open_google()

    elif "news" in command.lower():
        return fetch_news(command)
    
    elif any(token.text.lower() == "chrome" for token in doc):
        return open_chrome()

    elif any(token.text.lower() == "open" for token in doc) and any(
        platform.text.lower() in ["instagram", "twitter", "facebook"] for platform in doc):
        platform_to_open = next(
            platform.text.lower()
            for platform in doc
            if platform.text.lower() in ["instagram", "twitter", "facebook"]
        )
        return open_social_media(platform_to_open)

    elif any(token.text.lower() == "joke" for token in doc):
        return tell_joke()
    
    elif "open project ppt" in command.lower():
        return open_ppt()
    elif "tell me the time" in command.lower() or "time" in command.lower():
        return tell_time()

    elif "generate random number" in command.lower() or "random number" in command.lower():
        start, end = extract_range(command)
        generate_random_number(start, end)
        return "Generating a random number."

    elif "calculate" in command.lower():
        calculation = extract_calculation(command)
        result = calculate_long_calculation(calculation)
        return f'the result is {result}'

    elif "send email" in command.lower():
        send_email_via_gmail()

    elif "mail" in command.lower() and any(action in command.lower() for action in ["open", "", "check"]):
        return open_mail()

    elif "send whatsapp message" in command.lower() or "whatsapp message" in command.lower():
        send_whatsapp_message()
        return "Sending WhatsApp message."
    
    elif 'switch tab' in command.lower():
        pyautogui.hotkey('ctrl', 'tab')

    elif 'close tab' in command.lower():
        pyautogui.hotkey('ctrl', 'w')

    elif "write a note" in command.lower():
        write_note()
        return "Writing to a note."

    elif "create a note" in command.lower():
        create_note()
        return "Creating a new note."

    elif "open a note" in command.lower():
        open_note()
        return "Opening a note."

    elif "delete a note" in command.lower():
        delete_note()
        return "Deleting a note."

    elif "create a word file" in command.lower():
        create_word_file()
        return "Creating a new Word file."

    elif "open a word file" in command.lower():
        open_word_file()
        return "Opening a Word file."

    elif "write a word file" in command.lower():
        write_to_word_file()
        return "Writing to a Word file."

    elif "delete a word file" in command.lower():
        delete_word_file()
        return "Deleting a Word file."

    elif "search in youtube" in command.lower():
        query = command.lower().replace("search in youtube about", "").strip()
        return search_in_youtube(query)

    elif "search" in command.lower() and any(platform in command.lower() for platform in ["amazon", "flipkart"]):
        platform_to_search = next(platform for platform in ["amazon", "flipkart"] if platform in command.lower())
        product_to_search = command.lower().replace(f"search in {platform_to_search}", "").strip()
        return search_product(platform_to_search, product_to_search)

    elif "search" in command.lower():
        return search_wikipedia()

    elif "play" in command.lower():
        result = play_song()
        return result

    elif "tell date" in command.lower() or "date" in command.lower():
        result = tell_date_month_year_day("date")
        return result

    elif "tell month" in command.lower() or "month" in command.lower():
        result = tell_date_month_year_day("month")
        return result

    elif "tell year" in command.lower() or "year" in command.lower():
        result = tell_date_month_year_day("year")
        return result

    elif "tell day" in command.lower() or "day" in command.lower():
        result = tell_date_month_year_day("day")
        return result

    elif "add" in command.lower():
        operands = extract_operands(command)
        perform_arithmetic_operation("addition", *operands)
        return "Performing addition."

    elif "subtract" in command.lower():
        operands = extract_operands(command)
        perform_arithmetic_operation("subtraction", *operands)
        return "Performing subtraction."

    elif "multiply" in command.lower():
        operands = extract_operands(command)
        perform_arithmetic_operation("multiplication", *operands)
        return "Performing multiplication."

    elif "divide" in command.lower():
        operands = extract_operands(command)
        perform_arithmetic_operation("division", *operands)
        return "Performing division."

    elif "restart" in command.lower():
        return restart()

    elif "shutdown" in command.lower() or "shut down" in command.lower():
        return shutdown()

    elif "lock screen" in command.lower() or "lock" in command.lower():
        return lock_screen()

    elif "volume up" in command.lower():
        return control_volume("up")

    elif "open" in command.lower():
        app_name = extract_app_name(command)
        if app_name:
            return open_application(app_name)
        else:
            return "Could not identify the application name."

    elif "volume down" in command.lower():
        return control_volume("down")

    elif "full screen" in command.lower():
        return toggle_full_screen()

    elif "normal screen" in command.lower():
        return toggle_full_screen()

    elif "close" in command.lower():
        return close_current_task()

    elif "tell me the weather in" in command.lower() or "weather" in command.lower():
        return tell_weather(api_key='542ddfbe9d69543b92cbc5bf36ce28d2', city="Gandimaisamma")

    elif "exit" in command.lower() or "quit" in command.lower():
        return "existing voice assistant" and exit_voice_assistant()

    else:
        return "I'm still learning. Ask me something else."
