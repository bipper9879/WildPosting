import os
import sys
import ctypes
import shutil  # Using shutil.copy2 to preserve file timestamps
import tkinter as tk
from tkinter import messagebox, simpledialog

# Dynamically target your OneDrive layout path
USER_HOME = os.path.expanduser("~")
BASE_ONEDRIVE_PATH = os.path.join(USER_HOME, "OneDrive", "Documents", "posters")

def run_folder_name_sweep():
    root = tk.Tk()
    root.withdraw()

    # 1. Ask which market you want to scan right now
    city_input = simpledialog.askstring(
        "Select Market", 
        "Enter Market Code to scan:\n(Type: PHL, DC, or BOS)", 
        parent=root
    )
    if not city_input: return
    city_code = city_input.strip().upper()
    
    city_map = {"PHL": "philly", "DC": "dc", "BOS": "boston"}
    if city_code not in city_map:
        messagebox.showerror("Error", "Invalid Market Code. Use PHL, DC, or BOS.")
        return
        
    city_folder = city_map[city_code]
    master_folder_name = f"{city_folder.capitalize()}_Master"
    
    # 2. Establish our folder directories
    city_root_dir = os.path.join(BASE_ONEDRIVE_PATH, city_folder)
    city_master_dir = os.path.join(city_root_dir, master_folder_name)
    
    if not os.path.exists(city_master_dir):
        messagebox.showerror("Error", f"Authoritative Master folder not found at:\n{city_master_dir}\nPlease create it first.")
        return

    # 3. Index your authoritative master folder names (Forces lowercase comparison to handle human case typos)
    authoritative_folders = {f.lower().strip(): f for f in os.listdir(city_master_dir) if os.path.isdir(os.path.join(city_master_dir, f))}

    print(f"🚀 Scanning all backlog folders inside {city_root_dir}...")
    print(f"📋 Authoritative Master list loaded with {len(authoritative_folders)} target locations.")
    
    copied_files_count = 0
    mismatch_list = set() # Uses a set to prevent duplicate lines in your mismatch list

    # 4. Grab all top-level items in the city directory (looks at your dated folders)
    all_items = os.listdir(city_root_dir)
    
    for item in all_items:
        dated_folder_path = os.path.join(city_root_dir, item)
        
        # Skip the Authoritative Master folder and the Excel workbook file
        if item == master_folder_name or not os.path.isdir(dated_folder_path):
            continue
            
        # Look inside this specific dated campaign folder for location subfolders
        subfolders_in_dated = os.listdir(dated_folder_path)
        
        for sub_folder in subfolders_in_dated:
            current_sub_path = os.path.join(dated_folder_path, sub_folder)
            
            # Ensure it's a folder, not a loose file
            if not os.path.isdir(current_sub_path):
                continue
                
            clean_sub_name = sub_folder.lower().strip()
            
            # 5. Check if this subfolder name exists in your Authoritative Master directory
            if clean_sub_name in authoritative_folders:
                # Grab the correct case-sensitive folder name from our master list
                true_master_name = authoritative_folders[clean_sub_name]
                destination_dir = os.path.join(city_master_dir, true_master_name)
                
                # Copy every photo inside this matching folder over to the master tree
                for file_name in os.listdir(current_sub_path):
                    if file_name.lower().endswith(('.jpg', '.jpeg', '.png')):
                        src_file = os.path.join(current_sub_path, file_name)
                        dest_file = os.path.join(destination_dir, file_name)
                        
                        # Handle duplicate photo name collisions safely
                        if os.path.exists(dest_file):
                            base, ext = os.path.splitext(file_name)
                            dest_file = os.path.join(destination_dir, f"{base}_copy{ext}")
                            
                        # CRITICAL FIX: Changed from shutil.move to shutil.copy2
                        shutil.copy2(src_file, dest_file)
                        copied_files_count += 1
                        
                print(f"✅ Matched & Copied: Folder '{sub_folder}' inside [{item}] safely duplicated to Master.")
            else:
                # 6. Flag as a mismatch if the name doesn't exist in Philly_Master
                mismatch_list.add(f"• Dated Folder: [{item}]  -->  Broken Subfolder Name: '{sub_folder}'")

    # --- PHASE 5: REPORT SUMMARY ALERTS ---
    print(f"\n🏁 Sweep run finished.")
    
    if mismatch_list:
        # Build the user notification string
        report_msg = f"Folder Name Copy Sweep Complete!\n\nSuccessfully copied {copied_files_count} photos.\n\n⚠️ Found {len(mismatch_list)} mismatched folder names that do not match Philly_Master.\n\nReview your PowerShell terminal window right now for the exact list to fix!"
        
        print("\n❌ MISMATCHED FOLDERS FOUND (Fix these names inside your dated folders to match Philly_Master):")
        print("\n".join(sorted(mismatch_list)))
        
        show_popup("Sweep Complete (With Actions Needed)", report_msg, is_error=False)
    else:
        show_popup("Sweep Complete!", f"Perfect run! Successfully copied {copied_files_count} photos. Zero mismatches found.", is_error=False)

if __name__ == "__main__":
    run_folder_name_sweep()
