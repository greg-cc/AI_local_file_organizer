# comboaisort27.py
import os
import fitz  # PyMuPDF for PDF
from docx import Document
import pandas as pd
from pptx import Presentation
from openpyxl import load_workbook
from transformers import pipeline
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import shutil
import json
import colorama
from colorama import Fore, Style
import re
import sys

# Initialize colorama to auto-reset styles after each print
colorama.init(autoreset=True)

# Check for GPU availability and set the device
if torch.cuda.is_available():
    device = "cuda"
    print("Device set to use GPU.")
else:
    device = "cpu"
    print("GPU not found. Device set to use CPU.")

# --- Configuration ---
# These are now set via user input or loaded from a file.

# Define the categories for categorization
CATEGORIES = [
    "Health",
    "History",
    "Weather",
    "do not combine general categories with categories containing targeted things within it (delete this)"
]

# --- Load and Edit Categories Function ---
def load_and_edit_categories():
    """
    Loads categories from a file, prompts the user to edit them, and saves the final list.
    """
    category_file = "categories.json"
    default_categories = [
        "Health",
        "History",
        "Weather",
        "do not combine general categories with categories containing targeted things within it (delete this)"
    ]
    categories = default_categories
    
    try:
        if os.path.exists(category_file):
            with open(category_file, 'r', encoding='utf-8') as f:
                categories = json.load(f)
            print("Step 1.2: Loaded categories from previous session.")
    except Exception as e:
        print(f"Error loading categories file: {e}. Using default categories.")

    print("\n--- Current Categories ---")
    for i, cat in enumerate(categories):
        print(f"[{i+1}] {cat}")
    print("--------------------------")
    
    edit_choice = input("Do you want to edit these categories? (y/n) [default: no]: ").strip().lower()
    if edit_choice in ['y', 'yes']:
        instructions = "\nEditing categories. Enter 'add <new_category>', 'remove <number>', 'edit <number> <new_category>', 'list', or 'done' to finish."
        print(instructions)
        while True:
            command = input("> ").strip().lower()
            if command == 'done':
                break
            
            parts = command.split()
            if not parts:
                continue

            action = parts[0]
            if action == 'add' and len(parts) >= 2:
                new_cat = " ".join(parts[1:])
                categories.append(new_cat)
                print(f"Added: {new_cat}")
                # Omitted print statements for brevity
            elif action == 'remove' and len(parts) == 2 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    removed_cat = categories.pop(idx)
                    print(f"Removed: {removed_cat}")
                else:
                    print("Invalid category number.")
            elif action == 'edit' and len(parts) >= 3 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    new_cat = " ".join(parts[2:])
                    old_cat = categories[idx]
                    categories[idx] = new_cat
                    print(f"Edited: '{old_cat}' to '{new_cat}'")
                else:
                    print("Invalid category number.")
            elif action == 'list':
                print("\n--- Current Categories ---")
                for i, cat in enumerate(categories):
                    print(f"[{i+1}] {cat}")
                print("--------------------------")
            else:
                print("Invalid command. Please try again.")
    
    with open(category_file, 'w', encoding='utf-8') as f:
        json.dump(categories, f, indent=4)
    print("\nStep 1.3: Categories saved for next session.")
    
    return categories

# --- Load Models ---
print("\nStep 1: Loading summarization and classification models...")
summarizer = pipeline("summarization", model="t5-small", max_length=1024, device=device, framework="pt")
try:
    classifier = pipeline("zero-shot-classification", model="MoritzLaurer/xtremedistil-l6-h256-zeroshot-v1.1-all-33", device=device, multi_label=True)
    print("Step 1.1: Classification model loaded successfully.")
except Exception as e:
    print(f"Error loading classification model: {e}")
    classifier = None

# --- File Reading Functions ---
def read_chunks_from_file(file_path, start_chunk, end_chunk, max_chunk_length):
    ext = os.path.splitext(file_path)[1].lower()
    words = []
    total_words_needed = end_chunk * max_chunk_length
    try:
        if ext == ".pdf":
            with fitz.open(file_path) as doc:
                for page in doc:
                    words.extend(page.get_text().split())
                    if len(words) >= total_words_needed: break
        elif ext == ".docx":
            doc = Document(file_path)
            for para in doc.paragraphs:
                words.extend(para.text.split())
                if len(words) >= total_words_needed: break
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as f: words = f.read().split()
        elif ext == ".xlsx":
            sheets = pd.ExcelFile(file_path).sheet_names
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet)
                words.extend(df.to_string(index=False).split())
                if len(words) >= total_words_needed: break
        elif ext == ".pptx":
            presentation = Presentation(file_path)
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        words.extend(shape.text.split())
                        if len(words) >= total_words_needed: break
                if len(words) >= total_words_needed: break
    except Exception as e:
        print(f"Error reading {ext} file: {e}")
        return ""
    start_index = (start_chunk - 1) * max_chunk_length
    end_index = end_chunk * max_chunk_length
    return " ".join(words[start_index:end_index])

# --- Summarization Function ---
def summarize_text(text, max_chunk_length, max_summary_length, min_summary_length):
    if not text.strip(): return ["No text to summarize."]
    words = text.split()
    chunks = [" ".join(words[i:i + max_chunk_length]) for i in range(0, len(words), max_chunk_length)]
    summaries = []
    for i, chunk in enumerate(chunks):
        try:
            summary = summarizer(chunk, max_length=max_summary_length, min_length=min_summary_length, truncation=True)
            summaries.append(summary[0]['summary_text'])
        except Exception as e:
            print(f"Error during summarization of chunk {i + 1}: {e}")
            summaries.append("Summary failed for this chunk.")
    return ". ".join(summaries).split(". ")
    
# --- Categorization Function ---
def categorize_summary(summary_text, categories):
    if classifier is None: return {'labels': ['Other'], 'scores': [1.0]}
    try:
        return classifier(summary_text, candidate_labels=categories, multi_label=True)
    except Exception as e:
        print(f"Error during AI categorization: {e}")
        return {'labels': ['Other'], 'scores': [1.0]}

# --- File Tagging Function ---
def tag_file_properties(file_path, categories_to_tag):
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if not os.path.exists(file_path):
            print(f"Tagging skipped: File not found at {file_path}")
            return
        if ext == ".pdf":
            doc = fitz.open(file_path)
            metadata, keywords = doc.metadata, metadata.get("keywords", "")
            current_keywords = {kw.strip() for kw in keywords.split(';') if kw.strip()}
            for cat in categories_to_tag: current_keywords.add(cat)
            metadata["keywords"] = "; ".join(sorted(list(current_keywords)))
            doc.set_metadata(metadata)
            doc.save(file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
            doc.close()
        elif ext in [".docx", ".pptx", ".xlsx"]:
            # Logic for docx, pptx, xlsx
            obj = None
            if ext == ".docx": obj = Document(file_path)
            elif ext == ".pptx": obj = Presentation(file_path)
            elif ext == ".xlsx": obj = load_workbook(file_path)
            
            keywords = obj.core_properties.keywords if ext != ".xlsx" else obj.properties.keywords
            keywords = keywords or ""
            current_keywords = {kw.strip() for kw in keywords.split(';') if kw.strip()}
            for cat in categories_to_tag: current_keywords.add(cat)
            
            if ext != ".xlsx": obj.core_properties.keywords = "; ".join(sorted(list(current_keywords)))
            else: obj.properties.keywords = "; ".join(sorted(list(current_keywords)))
            obj.save(file_path)
        else:
            print(f"Tagging not supported for {ext} files.")
            return
        print(f"Tagged '{os.path.basename(file_path)}' with: {', '.join(categories_to_tag)}")
    except Exception as e:
        print(f"Error tagging file {os.path.basename(file_path)}: {e}")

# --- Shortcut Creation Function ---
def create_shortcut(source_path, shortcut_path):
    if sys.platform == "win32":
        shortcut_path += ".url"
        if os.path.exists(shortcut_path): return
        with open(shortcut_path, "w") as f:
            f.write("[InternetShortcut]\n")
            f.write(f"URL=file:///{os.path.abspath(source_path)}\n")
    else:
        if os.path.lexists(shortcut_path): return
        os.symlink(source_path, shortcut_path)

# --- Pre-processing Function (RESTORED) ---
def preprocess_by_filename(file_list, categories, file_management_settings, file_action):
    print("\n--- Starting Pre-processing Step (Filename Scan) ---")
    processed_files = set()
    remaining_files = list(file_list)
    
    for file_path in file_list:
        file_name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
        cleaned_file_name = file_name_no_ext.replace('_', ' ').replace('-', ' ')
        words_in_name = {word.lower() for word in cleaned_file_name.split()}
        
        found_category = None
        for cat in categories: 
            category_words = set(cat.lower().split())
            if category_words.issubset(words_in_name):
                found_category = cat
                break 
        
        if found_category and found_category in file_management_settings:
            settings = file_management_settings[found_category]
            prefix = settings['prefix']
            destination_folder = settings['destination']
            os.makedirs(destination_folder, exist_ok=True)

            new_file_name = f"{prefix}_{file_name_no_ext}" if prefix else file_name_no_ext
            destination_path = os.path.join(destination_folder, new_file_name)
            
            try:
                if file_action == 'shortcut':
                    create_shortcut(file_path, destination_path)
                    print(f"Pre-processed shortcut for: {os.path.basename(file_path)} -> {destination_folder}")
                else:
                    destination_path_with_ext = destination_path + os.path.splitext(file_path)[1]
                    if file_action == 'move':
                        shutil.move(file_path, destination_path_with_ext)
                        print(f"Pre-processed and moved: {os.path.basename(file_path)} -> {destination_folder}")
                    elif file_action == 'copy':
                        shutil.copy2(file_path, destination_path_with_ext)
                        print(f"Pre-processed and copied: {os.path.basename(file_path)} -> {destination_folder}")

                processed_files.add(file_path)
                remaining_files.remove(file_path)
            except Exception as e:
                print(f"Error pre-processing {os.path.basename(file_path)}: {e}")

    print(f"--- Pre-processing Complete: {len(processed_files)} files processed. ---")
    return remaining_files

# --- Main Processing Logic ---
def process_files_in_folder(file_list, categories, pdf_padding, pdf_chunks, generic_padding, generic_chunks, file_management_settings, 

confidence_threshold, secondary_confidence_threshold, pdf_chunk_length, generic_chunk_length, max_summary_length, min_summary_length, tagging_enabled, 

file_action, output_folder_path):
    if not file_list: print("\nNo remaining files to process."); return []
    print(f"\n--- Starting Main AI Processing for {len(file_list)} files... ---")
    all_summaries = []
    for i, file_path in enumerate(file_list):
        print(f"\n--- AI Processing file {i + 1}/{len(file_list)}: {os.path.basename(file_path)} ---")
        
        is_pdf = os.path.splitext(file_path)[1].lower() == '.pdf'
        padding, chunks_to_process, max_chunk_length = (pdf_padding, pdf_chunks, pdf_chunk_length) if is_pdf else (generic_padding, generic_chunks, 

generic_chunk_length)
        
        start_chunk, end_chunk = padding + 1, padding + chunks_to_process
        text = read_chunks_from_file(file_path, start_chunk, end_chunk, max_chunk_length)

        if text.strip():
            bullet_points = summarize_text(text, max_chunk_length, max_summary_length, min_summary_length)
            
            if bullet_points and bullet_points != ["No text to summarize."]:
                cleaned_file_name = os.path.splitext(os.path.basename(file_path))[0].replace('_', ' ').replace('-', ' ')
                text_for_classifier = ". ".join([cleaned_file_name] + bullet_points)
                results = categorize_summary(text_for_classifier, categories)
                
                primary_category, secondary_categories, all_valid_categories = "Other", [], []
                if results['scores'][0] * 100 >= confidence_threshold:
                    primary_category = results['labels'][0]
                    all_valid_categories.append(primary_category)
                    for label, score in zip(results['labels'][1:], results['scores'][1:]):
                        if score * 100 >= secondary_confidence_threshold:
                            secondary_categories.append(label)
                            all_valid_categories.append(label)
                
                if tagging_enabled and all_valid_categories:
                    tag_file_properties(file_path, all_valid_categories)
                
                print(f"Primary Category: {Fore.GREEN}{Style.BRIGHT}{primary_category} ({results['scores'][0]*100:.2f}%)")
                if secondary_categories: print(f"Secondary Categories: {Fore.YELLOW}{', '.join(secondary_categories)}")
                
                all_summaries.append({'file_path': os.path.abspath(file_path), 'primary_category': primary_category, 'all_categories': 

all_valid_categories, 'summary': "\n".join(bullet_points)})

                if file_action != 'none' and all_valid_categories:
                    for category in all_valid_categories:
                        if category in file_management_settings:
                            settings = file_management_settings[category]
                            prefix, destination_folder = settings['prefix'], settings['destination']
                            os.makedirs(destination_folder, exist_ok=True)

                            new_file_name = f"{prefix}_{cleaned_file_name}" if prefix else cleaned_file_name
                            destination_path = os.path.join(destination_folder, new_file_name)
                            
                            try:
                                if file_action == 'shortcut':
                                    create_shortcut(file_path, destination_path)
                                    print(f"Created shortcut for '{os.path.basename(file_path)}' in -> {destination_folder}")
                                elif category == primary_category:
                                    destination_path_with_ext = destination_path + os.path.splitext(file_path)[1]
                                    if file_action == 'move' and os.path.exists(file_path):
                                        shutil.move(file_path, destination_path_with_ext)
                                        if tagging_enabled: tag_file_properties(destination_path_with_ext, all_valid_categories)
                                    elif file_action == 'copy':
                                        shutil.copy2(file_path, destination_path_with_ext)
                                        if tagging_enabled: tag_file_properties(destination_path_with_ext, all_valid_categories)
                            except Exception as e:
                                print(f"Error during file operation: {e}")
    return all_summaries

# --- Main Execution ---
if __name__ == "__main__":
    try:
        CATEGORIES = load_and_edit_categories()
        settings_file = "settings.json"
        defaults = {
            'pdf_chunk_length': 500, 'generic_chunk_length': 250, 'max_summary_length': 150, 'min_summary_length': 40,
            'pdf_padding': 0, 'pdf_chunks': 1, 'generic_padding': 0, 'generic_chunks': 1,
            'confidence_threshold': 50, 'secondary_confidence_threshold': 20, 'tagging_enabled_choice': 'n',
            'add_prefix_choice': 'n', 'global_prefix_choice': 'n', 'global_prefix': '', 'special_char': '',
            'selected_file_types': ['.pdf', '.docx', '.txt', '.xlsx', '.pptx'], 'file_action_choice': 'none',
            'last_list_file': '', 'last_destination_folder': '', 'source_choice': 'f',
            'preprocess_choice': 'n' # New
        }

        if os.path.exists(settings_file):
            if input("Found saved settings. Load them? (y/n) [default: y]: ").strip().lower() not in ['n', 'no']:
                try:
                    with open(settings_file, 'r') as f: defaults.update(json.load(f))
                    print("Settings loaded successfully.")
                except Exception as e: print(f"Error loading settings: {e}.")
        
        # Simplified user input flow
        source_choice = input(f"Process files from [f]older or [l]ist file? [default: {defaults['source_choice']}]: ").strip().lower() or defaults

['source_choice']
        initial_file_list, output_folder_path, list_file_path = [], "", defaults.get('last_list_file', '')
        last_destination_folder = defaults.get('last_destination_folder', '')

        if source_choice == 'f':
            folder_path = input("Enter the folder path to scan: ").strip()
            dest_choice = input("Send organized files to [s]ubdirectory or [d]efined location?: ").strip().lower()
            output_folder_path = folder_path if dest_choice == 's' else (input(f"Enter destination folder [default: {last_destination_folder}]: ").strip() 

or last_destination_folder)
            if not os.path.isdir(output_folder_path): os.makedirs(output_folder_path, exist_ok=True)
            scan_subdirectories = input("Scan subdirectories? (y/n) [default: n]: ").strip().lower() in ['y', 'yes']
            
            # Populate file list
            file_walk = os.walk(folder_path) if scan_subdirectories else [(os.path.dirname(folder_path), [], os.listdir(folder_path))]
            for root, _, files in file_walk:
                for file in files:
                    if os.path.splitext(file)[1].lower() in defaults['selected_file_types']: initial_file_list.append(os.path.join(root, file))

        elif source_choice == 'l':
            list_file_path = input(f"Enter path to the list file [default: {list_file_path}]: ").strip() or list_file_path
            output_folder_path = input(f"Enter destination folder [default: {last_destination_folder}]: ").strip() or last_destination_folder
            if not os.path.isdir(output_folder_path): os.makedirs(output_folder_path, exist_ok=True)
            with open(list_file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    file_path = line.strip()
                    if os.path.isfile(file_path) and os.path.splitext(file_path)[1].lower() in defaults['selected_file_types']: initial_file_list.append

(file_path)

        action_map = {'m': 'move', 'c': 'copy', 's': 'shortcut', 'd': 'none'}
        action_input = input(f"Action: [M]ove, [C]opy, create [S]hortcut, or [D]o nothing? [default: {defaults['file_action_choice']}]: ").strip().lower()
        file_action_to_use = action_map.get(action_input, defaults['file_action_choice'] if not action_input else 'none')

        # NEW: Pre-processing choice
        preprocess_choice = input(f"Pre-process files based on filename before full AI analysis? (y/n) [default: {defaults['preprocess_choice']}]: 

").strip().lower() or defaults['preprocess_choice']
        
        # Build settings to save
        settings_to_save = defaults.copy()
        settings_to_save.update({
            'source_choice': source_choice,
            'last_list_file': list_file_path if source_choice == 'l' else defaults['last_list_file'],
            'last_destination_folder': output_folder_path,
            'file_action_choice': file_action_to_use,
            'preprocess_choice': preprocess_choice
        })
        with open(settings_file, 'w') as f: json.dump(settings_to_save, f, indent=4)
        print("\nSettings saved for the next session.")
        
        # Build file management dictionary
        file_management_settings = {}
        if file_action_to_use != 'none' and output_folder_path:
            for category in CATEGORIES:
                if category != "Other":
                    file_management_settings[category] = {'prefix': "", 'destination': os.path.join(output_folder_path, category)}

        # Execute processing
        files_to_process = initial_file_list
        if preprocess_choice in ['y', 'yes'] and file_action_to_use != 'none':
            files_to_process = preprocess_by_filename(initial_file_list, CATEGORIES, file_management_settings, file_action_to_use)

        all_summaries = process_files_in_folder(files_to_process, CATEGORIES, 
            defaults['pdf_padding'], defaults['pdf_chunks'], defaults['generic_padding'], defaults['generic_chunks'],
            file_management_settings, defaults['confidence_threshold'], defaults['secondary_confidence_threshold'],
            defaults['pdf_chunk_length'], defaults['generic_chunk_length'], defaults['max_summary_length'], defaults['min_summary_length'],
            defaults['tagging_enabled_choice'] in ['y', 'yes'], file_action_to_use, output_folder_path
        )

        if all_summaries and output_folder_path:
            summary_file_path = os.path.join(output_folder_path, "all_summaries.txt")
            with open(summary_file_path, "w", encoding="utf-8") as f:
                for summary in all_summaries:
                    f.write("="*50 + "\n")
                    f.write(f"File: {summary['file_path']}\n")
                    f.write(f"Primary Category: {summary['primary_category']}\n")
                    if summary['all_categories'][1:]: f.write(f"All Categories: {', '.join(summary['all_categories'])}\n")
                    f.write(f"\nSummary:\n{summary['summary']}\n")
            print(f"All summaries consolidated into: {summary_file_path}")
            
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. Exiting gracefully.")
