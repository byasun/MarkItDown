import os
import re
import datetime
from src.config import LOG_FILE

def ensure_directory_exists(directory):
    """Cria a pasta se não existir."""
    if not os.path.exists(directory):
        os.makedirs(directory)

def log_message(message, error=False):
    """Registra mensagens no log."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_type = "ERROR" if error else "INFO"
    log_entry = f"[{timestamp}] [{log_type}] {message}\n"
    
    print(log_entry, end="")  # Exibe no terminal
    
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        log_file.write(log_entry)

def clean_text(text):
    """Remove marcações como '<!-- Slide number: X -->' e espaços extras."""
    text = re.sub(r'<!-- Slide number: \d+ -->', '', text)  # Remove marcações de slide
    return text.strip()  # Remove espaços extras no início e fim