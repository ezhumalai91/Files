import os
import time
import re
import sys
import logging
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.controls.uia_controls import ButtonWrapper

# === Config Logging ===
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# === User Inputs ===
XML_FILE = input("Enter path to the XML file: ").strip()
OUTPUT_DIR = os.path.join(os.path.dirname(XML_FILE), "eps_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

MATHTYPE_PATH = r"C:\Program Files (x86)\MathType\MathType.exe"
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")
REPORT_FILE = os.path.join(DESKTOP_PATH, "EquationFontReport.txt")
MATHML_TXT_PATH = os.path.join(DESKTOP_PATH, "mathml.txt")

# === Global counter ===
eqn_counter = 1

# === Helper Functions ===

def extract_mathml_blocks(xml_path):
    with open(xml_path, 'r', encoding='utf-8') as f:
        xml_content = f.read()
    matches = re.findall(r'<math[^>]*altimg="([^"]+)"[^>]*>(.*?)</math>', xml_content, re.DOTALL)
    blocks = [(f'<math altimg="{alt}">{content}</math>', alt) for alt, content in matches]
    return blocks

def write_mathml_to_file(blocks, filepath):
    with open(filepath, 'w', encoding='utf-8') as f:
        for block, _ in blocks:
            f.write(block.strip().replace('\n', '') + '\n')
    logging.info(f"📄 Saved all MathML blocks to: {filepath}")

def read_eps_preferences(report_file):
    """Reads a report file and returns a dict of {eps_filename: preference_name}"""
    mapping = {}
    if not os.path.exists(report_file):
        logging.warning(f"Preference report file not found: {report_file}")
        return mapping

    with open(report_file, 'r', encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split('\t')
            if len(parts) >= 2:
                eps_name = parts[0].strip()
                pref_name = parts[1].strip()
                mapping[eps_name] = pref_name
    return mapping

def connect_or_start_mathtype():
    try:
        app = Application().connect(path="MathType.exe")
        logging.info("Connected to running MathType.")
    except Exception:
        app = Application().start(MATHTYPE_PATH)
        logging.info("Started new instance of MathType.")
        time.sleep(3)
    return app

def get_main_window(app, timeout=10):
    """Try to get MathType main window within timeout, else None"""
    start = time.time()
    while time.time() - start < timeout:
        for w in app.windows():
            if "MathType" in w.window_text() and w.class_name() == "EQNWINCLASS":
                return w
        time.sleep(0.5)
    return None

def paste_mathml_to_mathtype(mathml):
    send_keys("^a{BACKSPACE}")
    time.sleep(0.3)
    send_keys("^v")
    time.sleep(1)

def apply_preferences(app, preference_name):
    main_win = get_main_window(app)
    if not main_win:
        logging.warning("MathType window not found during preference application.")
        return

    main_win.set_focus()
    send_keys("^a")  # Select all
    time.sleep(0.5)

    try:
        main_win.menu_select(f"Preferences->{preference_name}")
        logging.info(f"Applied preference: {preference_name}")
        time.sleep(1)

        # Handle optional dialog
        try:
            pref_dlg = app.window(title_re="Load Equation Preferences from File")
            if pref_dlg.exists(timeout=3):
                pref_dlg["OK"].click()
                logging.info("Clicked OK in preferences dialog.")
        except Exception:
            pass

    except Exception as e:
        logging.warning(f"Could not apply preference '{preference_name}': {e}")

def save_as_eps(app):
    global eqn_counter

    main_win = get_main_window(app)
    if not main_win:
        raise RuntimeError("MathType main window not found during EPS save.")

    main_win.menu_select("File->Save As")
    time.sleep(1)

    save_dlg = app.window(title_re=".*Save.*")
    save_dlg.wait("visible", timeout=10)

    eps_name = f"Eqn{eqn_counter}.eps"
    eps_path = os.path.join(OUTPUT_DIR, eps_name)
    eqn_counter += 1

    save_dlg.Edit.set_edit_text(eps_path)
    time.sleep(0.5)
    save_dlg["Save"].click()
    time.sleep(1)

    try:
        confirm = app.window(title_re=".*Confirm Save.*")
        if confirm.exists(timeout=2):
            confirm["Yes"].click()
    except Exception:
        pass

    logging.info(f"✅ Saved EPS: {eps_path}")
    return eps_name

def close_current_window(app):
    try:
        main_win = get_main_window(app)
        if main_win:
            main_win.set_focus()
            send_keys("^w")
            time.sleep(0.5)
            logging.info("Closed current MathType window.")
    except Exception:
        logging.warning("Could not close MathType window.")

import time
import logging
from pywinauto.keyboard import send_keys

def dismiss_error_popup(app, timeout=5):
    try:
        all_windows = app.windows()
        for w in all_windows:
            title = w.window_text()
            lower_title = title.lower()
            # Check for popup windows with typical keywords
            if any(keyword in lower_title for keyword in ("error", "math", "warning", "confirm", "save", "server")):
                logging.info(f"Popup detected with title: '{title}'")

                # Print control identifiers for debugging
                try:
                    w.print_control_identifiers()
                except Exception as e:
                    logging.warning(f"Failed to print control identifiers: {e}")

                # Try to set focus and send ESC to close popup
                try:
                    w.set_focus()
                    send_keys("{ESC}")
                    logging.info(f"Sent ESC key to popup '{title}'")
                    time.sleep(1)  # Give time for popup to close
                    return True
                except Exception as e_key:
                    logging.warning(f"Failed to send keys to popup '{title}': {e_key}")
                    continue

        logging.info("No popup dialog detected to dismiss.")
        return False
    except Exception as e:
        logging.warning(f"Error while trying to detect/dismiss popup: {e}")
        return False



# === Main Processing ===

def process_mathml_blocks_from_file(txt_path, eps_pref_map):
    import pyperclip

    app = connect_or_start_mathtype()
    main_win = get_main_window(app)
    main_win.set_focus()
    global eqn_counter

    with open(txt_path, 'r', encoding='utf-8') as f:
        for i, line in enumerate(f, 1):
            block = line.strip()
            if not block:
                continue

            logging.info(f"\n--- Processing equation {i} ---")

            try:
                pyperclip.copy(block)
                paste_mathml_to_mathtype(block)

                eps_filename = f"Eqn{eqn_counter}.eps"
                preference = eps_pref_map.get(eps_filename)

                if preference:
                    apply_preferences(app, preference)
                else:
                    logging.warning(f"No preference found for {eps_filename}, saving without applying preference.")

                save_as_eps(app)
                close_current_window(app)

            except Exception as e:
                logging.error(f"❌ Error processing equation {i}: {e}")

                # Attempt to dismiss popup error dialog if it exists
                dismissed = dismiss_error_popup(app)
                if dismissed:
                    logging.info("Popup dismissed, moving to next equation.")
                else:
                    logging.warning("No popup found to dismiss; moving on.")

                # Save a failed empty EPS file for this equation
                failed_eps_name = f"Eqn{eqn_counter}_failed.eps"
                failed_eps_path = os.path.join(OUTPUT_DIR, failed_eps_name)
                with open(failed_eps_path, 'w', encoding='utf-8') as f:
                    f.write('')
                logging.warning(f"Saved failed EPS placeholder: {failed_eps_path}")
                eqn_counter += 1

            time.sleep(1)

    logging.info("\n✅ Done processing all MathML blocks from file.")

# === Entry Point ===

if __name__ == "__main__":
    try:
        import pyperclip
    except ImportError:
        print("❌ Please install pyperclip first:\n   pip install pyperclip")
        sys.exit(1)

    if not os.path.exists(XML_FILE):
        print(f"❌ XML file does not exist: {XML_FILE}")
        sys.exit(1)

    if not os.path.exists(MATHTYPE_PATH):
        print(f"❌ MathType not found at path: {MATHTYPE_PATH}")
        sys.exit(1)

    mathml_blocks = extract_mathml_blocks(XML_FILE)
    if not mathml_blocks:
        print("❌ No <math> blocks with altimg found in the XML.")
        sys.exit(1)

    logging.info(f"🔍 Found {len(mathml_blocks)} math blocks.")

    write_mathml_to_file(mathml_blocks, MATHML_TXT_PATH)
    eps_pref_map = read_eps_preferences(REPORT_FILE)
    process_mathml_blocks_from_file(MATHML_TXT_PATH, eps_pref_map)
