# comboaisort33.py
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

# --- Helper Function ---
def sanitize_for_path(name):
    """Truncates and sanitizes a string to be a valid directory/file name."""
    sanitized = re.sub(r'[<>:"/\\|?*]', '', name)
    return sanitized[:50].strip()

# --- Load and Edit Categories Function ---
def load_and_edit_categories():
    category_file = "categories.json"
    default_categories = ["Example Category 1", "Example Category 2"]
    categories = default_categories
    try:
        if os.path.exists(category_file):
            with open(category_file, 'r', encoding='utf-8') as f:
                categories = json.load(f)
            print("Step 1.2: Loaded categories from previous session.")
    except Exception as e:
        print(f"Error loading categories file: {e}. Using default categories.")

    print("\n--- Current Categories ---")
    for i, cat in enumerate(categories): print(f"[{i+1}] {cat}")
    print("--------------------------")
    
    if input("Do you want to edit these categories? (y/n) [default: no]: ").strip().lower() in ['y', 'yes']:
        instructions = "\nEditing categories: 'add <name>', 'remove <num>', 'edit <num> <name>', 'list', 'done'"
        print(instructions)
        while True:
            command = input("> ").strip().lower()
            if command == 'done': break
            parts = command.split()
            if not parts: continue
            action = parts[0]
            if action == 'add' and len(parts) >= 2: categories.append(" ".join(parts[1:])); print(f"Added: {' '.join(parts[1:])}")
            elif action == 'remove' and len(parts) == 2 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories): print(f"Removed: {categories.pop(idx)}")
                else: print("Invalid number.")
            elif action == 'edit' and len(parts) >= 3 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    old_cat = categories[idx]
                    categories[idx] = " ".join(parts[2:])
                    print(f"Edited '{old_cat}' to '{categories[idx]}'")
                else: print("Invalid number.")
            elif action == 'list':
                for i, cat in enumerate(categories): print(f"[{i+1}] {cat}")
            else: print("Invalid command.")
    
    with open(category_file, 'w', encoding='utf-8') as f: json.dump(categories, f, indent=4)
    print("\nStep 1.3: Categories saved.")
    return categories

# --- Load Models ---
print("\nStep 1: Loading AI models...")
summarizer = pipeline("summarization", model="t5-small", max_length=1024, device=device, framework="pt")
try:
    classifier = pipeline("zero-shot-classification", model="MoritzLaurer/xtremedistil-l6-h256-zeroshot-v1.1-all-33", device=device, multi_label=True)
    print("Step 1.1: Classification model loaded.")
except Exception as e:
    print(f"Error loading classification model: {e}"); classifier = None

# --- File I/O & AI Functions ---
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
            for sheet in pd.ExcelFile(file_path).sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet) # CORRECTED LINE
                words.extend(df.to_string(index=False).split())
                if len(words) >= total_words_needed: break
        elif ext == ".pptx":
            prs = Presentation(file_path)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        words.extend(shape.text.split())
                        if len(words) >= total_words_needed: break
                if len(words) >= total_words_needed: break
    except Exception as e:
        print(f"Error reading {ext} file: {e}"); return ""
    start_index, end_index = (start_chunk - 1) * max_chunk_length, end_chunk * max_chunk_length
    return " ".join(words[start_index:end_index])

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
            summaries.append("Summary failed.")
            print(f"Error summarizing chunk {i + 1}: {e}")
    return ". ".join(summaries).split(". ")
    
def categorize_summary(summary_text, categories):
    if classifier is None: return {'labels': ['Other'], 'scores': [1.0]}
    try:
        return classifier(summary_text, candidate_labels=categories, multi_label=True)
    except Exception as e:
        print(f"Error during categorization: {e}"); return {'labels': ['Other'], 'scores': [1.0]}

def tag_file_properties(file_path, categories_to_tag):
    ext = os.path.splitext(file_path)[1].lower()
    if not os.path.exists(file_path):
        print(f"Tagging skipped: File not found at {file_path}"); return
    try:
        current_keywords, obj, keywords = set(), None, ""
        if ext == ".pdf":
            doc = fitz.open(file_path)
            keywords = doc.metadata.get("keywords", "")
            current_keywords = {kw.strip() for kw in keywords.split(';') if kw.strip()}
            for cat in categories_to_tag: current_keywords.add(cat)
            doc.metadata["keywords"] = "; ".join(sorted(list(current_keywords)))
            doc.save(file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP); doc.close()
        elif ext in [".docx", ".pptx", ".xlsx"]:
            if ext == ".docx": obj = Document(file_path); keywords = obj.core_properties.keywords
            elif ext == ".pptx": obj = Presentation(file_path); keywords = obj.core_properties.keywords
            elif ext == ".xlsx": obj = load_workbook(file_path); keywords = obj.properties.keywords
            current_keywords = {kw.strip() for kw in (keywords or "").split(';') if kw.strip()}
            for cat in categories_to_tag: current_keywords.add(cat)
            if ext != ".xlsx": obj.core_properties.keywords = "; ".join(sorted(list(current_keywords)))
            else: obj.properties.keywords = "; ".join(sorted(list(current_keywords)))
            obj.save(file_path)
        else:
            print(f"Tagging not supported for {ext} files."); return
        print(f"Tagged '{os.path.basename(file_path)}' with: {', '.join(categories_to_tag)}")
    except Exception as e:
        print(f"Error tagging file {os.path.basename(file_path)}: {e}")

def create_shortcut(source_path, shortcut_path):
    if sys.platform == "win32":
        shortcut_path += ".url"
        if os.path.exists(shortcut_path): return
        with open(shortcut_path, "w") as f:
            f.write(f"[InternetShortcut]\nURL=file:///{os.path.abspath(source_path)}\n")
    else:
        if os.path.lexists(shortcut_path): return
        os.symlink(source_path, shortcut_path)

# --- Pre-processing Function ---
def preprocess_by_filename(file_list, categories, file_action, output_folder_path, file_management_settings):
    print("\n--- Starting Pre-processing Step (Filename Scan) ---")
    remaining_files = list(file_list)
    
    for file_path in file_list:
        file_name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
        words_in_name = {word.lower() for word in file_name_no_ext.replace('_', ' ').replace('-', ' ').split()}
        
        found_category = None
        for cat in categories: 
            if set(cat.lower().split()).issubset(words_in_name):
                found_category = cat; break 
        
        if found_category:
            sanitized_cat = sanitize_for_path(found_category)
            destination_folder = os.path.join(output_folder_path, sanitized_cat)
            os.makedirs(destination_folder, exist_ok=True)
            
            prefix = file_management_settings.get(found_category, {}).get('prefix', '')
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
                remaining_files.remove(file_path)
            except Exception as e:
                print(f"Error pre-processing {os.path.basename(file_path)}: {e}")

    print(f"--- Pre-processing Complete: {len(file_list) - len(remaining_files)} files processed. ---")
    return remaining_files

# --- Main Processing Logic ---
def process_files_in_folder(file_list, categories, file_management_settings, settings):
    if not file_list: print("\nNo remaining files to process."); return []
    print(f"\n--- Starting Main AI Processing for {len(file_list)} files... ---")
    all_summaries = []
    for i, file_path in enumerate(file_list):
        print(f"\n--- AI Processing file {i + 1}/{len(file_list)}: {os.path.basename(file_path)} ---")
        
        is_pdf = os.path.splitext(file_path)[1].lower() == '.pdf'
        padding = settings['pdf_padding'] if is_pdf else settings['generic_padding']
        chunks_to_process = settings['pdf_chunks'] if is_pdf else settings['generic_chunks']
        max_chunk_length = settings['pdf_chunk_length'] if is_pdf else settings['generic_chunk_length']
        
        text = read_chunks_from_file(file_path, padding + 1, padding + chunks_to_process, max_chunk_length)
        if not text.strip(): print("No text found in range."); continue

        bullet_points = summarize_text(text, max_chunk_length, settings['max_summary_length'], settings['min_summary_length'])
        if not bullet_points or bullet_points == ["No text to summarize."]: print("Summarization failed."); continue

        cleaned_file_name = os.path.splitext(os.path.basename(file_path))[0].replace('_', ' ').replace('-', ' ')
        results = categorize_summary(". ".join([cleaned_file_name] + bullet_points), categories)
        
        primary_category, all_valid_categories = "Other", []
        if results['scores'][0] * 100 >= settings['confidence_threshold']:
            primary_category = results['labels'][0]
            all_valid_categories.append(primary_category)
            for label, score in zip(results['labels'][1:], results['scores'][1:]):
                if score * 100 >= settings['secondary_confidence_threshold']:
                    all_valid_categories.append(label)
        
        if settings['tagging_enabled_choice'] in ['y', 'yes'] and all_valid_categories:
            tag_file_properties(file_path, all_valid_categories)
        
        print(f"Primary Category: {Fore.GREEN}{Style.BRIGHT}{primary_category} ({results['scores'][0]*100:.2f}%)")
        if len(all_valid_categories) > 1: print(f"Secondary Categories: {Fore.YELLOW}{', '.join(all_valid_categories[1:])}")
        
        all_summaries.append({'file_path': os.path.abspath(file_path), 'primary_category': primary_category, 'all_categories': all_valid_categories, 

'summary': "\n".join(bullet_points)})

        if settings['file_action_choice'] != 'none' and all_valid_categories:
            for category in all_valid_categories:
                sanitized_cat = sanitize_for_path(category)
                destination_folder = os.path.join(settings['output_folder_path'], sanitized_cat)
                os.makedirs(destination_folder, exist_ok=True)

                prefix = file_management_settings.get(category, {}).get('prefix', '')
                new_file_name = f"{prefix}_{cleaned_file_name}" if prefix else cleaned_file_name
                destination_path = os.path.join(destination_folder, new_file_name)
                
                try:
                    if settings['file_action_choice'] == 'shortcut':
                        create_shortcut(file_path, destination_path)
                        print(f"Created shortcut in -> {destination_folder}")
                    elif category == primary_category:
                        destination_path_with_ext = destination_path + os.path.splitext(file_path)[1]
                        if settings['file_action_choice'] == 'move' and os.path.exists(file_path):
                            shutil.move(file_path, destination_path_with_ext)
                            if settings['tagging_enabled_choice'] in ['y', 'yes']: tag_file_properties(destination_path_with_ext, all_valid_categories)
                        elif settings['file_action_choice'] == 'copy':
                            shutil.copy2(file_path, destination_path_with_ext)
                            if settings['tagging_enabled_choice'] in ['y', 'yes']: tag_file_properties(destination_path_with_ext, all_valid_categories)
                except Exception as e: print(f"Error during file operation: {e}")
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
            'last_list_file': '', 'last_destination_folder': '', 'source_choice': 'f', 'preprocess_choice': 'n'
        }

        if os.path.exists(settings_file):
            if input("Found saved settings. Load them? (y/n) [default: y]: ").strip().lower() not in ['n', 'no']:
                try:
                    with open(settings_file, 'r') as f: defaults.update(json.load(f))
                    print("Settings loaded successfully.")
                except Exception as e: print(f"Error loading settings: {e}.")
        
        settings_to_save = defaults.copy()
        
        source_choice = input(f"Process files from a [f]older or a [l]ist file? [default: {settings_to_save['source_choice']}]: ").strip().lower() or 

settings_to_save['source_choice']
        
        initial_file_list, output_folder_path, list_file_path = [], "", settings_to_save.get('last_list_file', '')
        last_destination_folder = settings_to_save.get('last_destination_folder', '')

        if source_choice == 'f':
            folder_path = input("Enter the folder path to scan: ").strip()
            dest_choice = input("Send organized files to [s]ubdirectory or [d]efined location?: ").strip().lower()
            output_folder_path = folder_path if dest_choice == 's' else (input(f"Enter the destination folder path [default: {last_destination_folder}]: 

").strip() or last_destination_folder)
            scan_sub = input("Scan subdirectories of the source folder? (y/n) [default: n]: ").strip().lower() in ['y', 'yes']
            
            walk_target = os.walk(folder_path) if scan_sub else [(os.path.dirname(folder_path) if folder_path else '.', [], os.listdir(folder_path or 

'.'))]
            for root, _, files in walk_target:
                for file in files:
                    if os.path.splitext(file)[1].lower() in settings_to_save['selected_file_types']: initial_file_list.append(os.path.join(root, file))

        elif source_choice == 'l':
            list_file_path = input(f"Enter path to the list file [default: {list_file_path}]: ").strip() or list_file_path
            output_folder_path = input(f"Enter a destination folder path [default: {last_destination_folder}]: ").strip() or last_destination_folder
            with open(list_file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    file_path = line.strip()
                    if os.path.isfile(file_path) and os.path.splitext(file_path)[1].lower() in settings_to_save['selected_file_types']: 

initial_file_list.append(file_path)

        if not os.path.isdir(output_folder_path): os.makedirs(output_folder_path, exist_ok=True)
        settings_to_save['output_folder_path'] = output_folder_path

        action_map = {'m': 'move', 'c': 'copy', 's': 'shortcut', 'd': 'none'}
        action_input = input(f"Action: [M]ove, [C]opy, create [S]hortcut, or [D]o nothing? [default: {settings_to_save['file_action_choice']}]: ").strip

().lower()
        settings_to_save['file_action_choice'] = action_map.get(action_input, settings_to_save['file_action_choice'] if not action_input else 'none')

        settings_to_save['preprocess_choice'] = input(f"Pre-process by filename? (y/n) [default: {settings_to_save['preprocess_choice']}]: ").strip

().lower() or settings_to_save['preprocess_choice']
        
        print("\n--- Configuration Settings ---")
        settings_to_save['confidence_threshold'] = int(input(f"Enter primary confidence threshold (0-100) [default: {settings_to_save

['confidence_threshold']}]: ").strip() or settings_to_save['confidence_threshold'])
        settings_to_save['secondary_confidence_threshold'] = int(input(f"Enter secondary confidence threshold (0-100) [default: {settings_to_save

['secondary_confidence_threshold']}]: ").strip() or settings_to_save['secondary_confidence_threshold'])
        settings_to_save['generic_chunk_length'] = int(input(f"Enter non-PDF chunk length [default: {settings_to_save['generic_chunk_length']}]: ").strip

() or settings_to_save['generic_chunk_length'])
        settings_to_save['pdf_chunk_length'] = int(input(f"Enter PDF chunk length [default: {settings_to_save['pdf_chunk_length']}]: ").strip() or 

settings_to_save['pdf_chunk_length'])
        settings_to_save['max_summary_length'] = int(input(f"Enter max summary length [default: {settings_to_save['max_summary_length']}]: ").strip() or 

settings_to_save['max_summary_length'])
        settings_to_save['min_summary_length'] = int(input(f"Enter min summary length [default: {settings_to_save['min_summary_length']}]: ").strip() or 

settings_to_save['min_summary_length'])
        settings_to_save['pdf_padding'] = int(input(f"PDF chunks to skip (padding) [default: {settings_to_save['pdf_padding']}]: ").strip() or 

settings_to_save['pdf_padding'])
        settings_to_save['pdf_chunks'] = int(input(f"PDF chunks to process [default: {settings_to_save['pdf_chunks']}]: ").strip() or settings_to_save

['pdf_chunks'])
        settings_to_save['generic_padding'] = int(input(f"Non-PDF chunks to skip (padding) [default: {settings_to_save['generic_padding']}]: ").strip() or 

settings_to_save['generic_padding'])
        settings_to_save['generic_chunks'] = int(input(f"Non-PDF chunks to process [default: {settings_to_save['generic_chunks']}]: ").strip() or 

settings_to_save['generic_chunks'])
        settings_to_save['tagging_enabled_choice'] = input(f"Tag files with category keywords? (y/n) [default: {settings_to_save

['tagging_enabled_choice']}]: ").strip().lower() or settings_to_save['tagging_enabled_choice']

        if settings_to_save['file_action_choice'] != 'none':
            settings_to_save['add_prefix_choice'] = input(f"Add a prefix to organized files? (y/n) [default: {settings_to_save['add_prefix_choice']}]: 

").strip().lower() or settings_to_save['add_prefix_choice']
            if settings_to_save['add_prefix_choice'] in ['y', 'yes']:
                settings_to_save['global_prefix_choice'] = input(f"Apply a global prefix? (y/n) [default: {settings_to_save['global_prefix_choice']}]: 

").strip().lower() or settings_to_save['global_prefix_choice']
                if settings_to_save['global_prefix_choice'] in ['y', 'yes']:
                    settings_to_save['global_prefix'] = input(f"Enter a global prefix (max 12 chars) [default: {settings_to_save['global_prefix']}]: 

").strip() or settings_to_save['global_prefix']
                    settings_to_save['special_char'] = input(f"Enter a special character (1-3 chars) [default: {settings_to_save['special_char']}]: 

").strip() or settings_to_save['special_char']
        
        settings_to_save['source_choice'] = source_choice
        settings_to_save['last_destination_folder'] = output_folder_path
        with open(settings_file, 'w') as f: json.dump(settings_to_save, f, indent=4)
        print("\nSettings saved. Starting process...")
        
        file_management_settings = {}
        if settings_to_save['file_action_choice'] != 'none' and output_folder_path:
            for category in CATEGORIES:
                if category != "Other":
                    prefix_str = ""
                    if settings_to_save['add_prefix_choice'] in ['y', 'yes']:
                        if settings_to_save['global_prefix_choice'] in ['y', 'yes']: prefix_str = f"{settings_to_save['global_prefix']}{settings_to_save

['special_char']}"
                        else: prefix_str = sanitize_for_path(category).replace(' ', '')[:12].upper()
                    file_management_settings[category] = {'prefix': prefix_str, 'destination': os.path.join(output_folder_path, sanitize_for_path

(category))}
        
        files_to_process = initial_file_list
        if settings_to_save['preprocess_choice'] in ['y', 'yes'] and settings_to_save['file_action_choice'] != 'none':
            files_to_process = preprocess_by_filename(initial_file_list, CATEGORIES, settings_to_save['file_action_choice'], output_folder_path, 

file_management_settings)

        all_summaries = process_files_in_folder(files_to_process, CATEGORIES, file_management_settings, settings_to_save)

        if all_summaries and output_folder_path:
            summary_file_path = os.path.join(output_folder_path, "all_summaries.txt")
            with open(summary_file_path, "w", encoding="utf-8") as f:
                for summary in all_summaries:
                    f.write("="*50 + "\n")
                    f.write(f"File: {summary['file_path']}\n")
                    f.write(f"Primary Category: {summary['primary_category']}\n")
                    if len(summary['all_categories']) > 1: f.write(f"All Categories: {', '.join(summary['all_categories'])}\n")
                    f.write(f"\nSummary:\n{summary['summary']}\n")
            print(f"Summaries saved to: {summary_file_path}")
            
    except KeyboardInterrupt:
        print("\nProcess interrupted.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
