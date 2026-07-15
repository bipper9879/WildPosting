import os
import sys
import ctypes
import re
import pandas as pd
from rapidfuzz import process, fuzz
import tkinter as tk
from tkinter import messagebox, simpledialog

USER_HOME = os.path.expanduser("~")
BASE_LOCAL_PATH = os.path.join(USER_HOME, "OneDrive", "Documents", "posters")

def show_popup(title, message, is_error=True):
    """Triggers native Windows pop-up dialog boxes with clickable inputs."""
    if is_error:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10 | 0x00)
    else:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x00)

def strip_leading_route_number(text):
    """
    Checks if a folder starts with a 1 or 2-digit number less than 35.
    Strips it and any following symbols, returning the clean street name.
    """
    if not text: return ""
    
    match = re.match(r'^(\d{1,2})[\s\.\-_/]*', text.strip())
    if match:
        route_num = int(match.group(1))
        if route_num < 35:
            return re.sub(r'^\d{1,2}[\s\.\-_/]*', '', text.strip())
            
    return text.strip()

def clean_text_for_match(text):
    """Deep cleans text for the final fuzzy matching fallback."""
    if not text: return ""
    clean = text.lower().strip()
    clean = re.sub(r'\b(and|btw|between|st|street|ave|avenue|rd|road|blvd)\b', '', clean)
    clean = re.sub(r'[^a-z0-9]', '', clean)
    return clean

def run_fast_folder_fixer():
    root = tk.Tk()
    root.withdraw()

    city_input = simpledialog.askstring("Select Market", "Enter Market Code to fix:\n(Type: PHL, DC, or BOS)", parent=root)
    if not city_input: return
    city_folder = {"PHL": "philly", "DC": "dc", "BOS": "boston"}.get(city_input.strip().upper())
    
    if not city_folder:
        messagebox.showerror("Error", "Invalid Market Code.")
        return

    city_root_dir = os.path.join(BASE_LOCAL_PATH, city_folder)
    master_folder_name = f"{city_folder.capitalize()}_Master"
    city_master_dir = os.path.join(city_root_dir, master_folder_name)
    manual_fix_excel_path = os.path.join(city_root_dir, "needs_manual_fixing.xlsx")

    master_folders = [f for f in os.listdir(city_master_dir) if os.path.isdir(os.path.join(city_master_dir, f))]
    if not master_folders:
        messagebox.showerror("Error", "No authoritative folders found inside Master!")
        return

    master_exact_lowercase = {f.lower().strip(): f for f in master_folders}
    master_fuzzy_lookup = {clean_text_for_match(f): f for f in master_folders}
    cleaned_master_keys = list(master_fuzzy_lookup.keys())

    print("🚀 Initializing Direct-Strip Auto-Rename Engine...")
    
    auto_rename_count = 0
    manual_review_rows = []
    
    all_items = os.listdir(city_root_dir)
    for item in all_items:
        dated_folder_path = os.path.join(city_root_dir, item)
        if item == master_folder_name or not os.path.isdir(dated_folder_path):
            continue
            
        subfolders = os.listdir(dated_folder_path)
        for sub in subfolders:
            current_sub_path = os.path.join(dated_folder_path, sub)
            if not os.path.isdir(current_sub_path): continue
            
            if sub.lower().strip() in master_exact_lowercase:
                continue
                
            stripped_name = strip_leading_route_number(sub)
            
            if stripped_name != sub:
                if stripped_name.lower().strip() in master_exact_lowercase:
                    true_name = master_exact_lowercase[stripped_name.lower().strip()]
                    new_path = os.path.join(dated_folder_path, true_name)
                    
                    if not os.path.exists(new_path):
                        os.rename(current_sub_path, new_path)
                        auto_rename_count += 1
                        continue

            clean_sub = clean_text_for_match(stripped_name)
            best_match_key, score, _ = process.extractOne(clean_sub, cleaned_master_keys, scorer=fuzz.token_sort_ratio)
            
            proposed_name = master_fuzzy_lookup[best_match_key] if score >= 50 else "COULD NOT GUESS"
            
            manual_review_rows.append({
                "Current Broken Name": sub,
                "Stripped Name Attempt": stripped_name,
                "Proposed Match": proposed_name,
                "Match Confidence %": round(score, 1),
                "Dated Folder": item,
                "Folder Path": current_sub_path
            })

    summary_msg = f"Fast Fix Complete!\n\n✅ Instantly stripped route numbers and renamed {auto_rename_count} folders on your hard drive!"
    
    if manual_review_rows:
        review_df = pd.DataFrame(manual_review_rows).sort_values(by=["Match Confidence %"], ascending=True)
        review_df.to_excel(manual_fix_excel_path, index=False, engine='openpyxl')
        summary_msg += f"\n\n⚠️ {len(manual_review_rows)} folders still don't match. They were saved to a tiny cleanup sheet named 'needs_manual_fixing.xlsx' inside your folder."
        
    show_popup("Renamer Done", summary_msg, is_error=False)

if __name__ == "__main__":
    run_fast_folder_fixer()
