# comboaisort41.py
import os
import sys

# Suppress TensorFlow/oneDNN warnings
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'
os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'

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
import multiprocessing

# Initialize colorama to auto-reset styles after each print
colorama.init(autoreset=True)

# --- Global Variables for Models ---
summarizer = None
classifier = None

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

# --- Model and I/O Functions ---
def initialize_models():
    global summarizer, classifier
    print("\nStep 1: Loading AI models...")
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"Device set to use {'GPU' if device == 'cuda' else 'CPU'}.")
    summarizer = pipeline("summarization", model="t5-small", max_length=1024, device=device, framework="pt")
    try:
        classifier = pipeline("zero-shot-classification", model="MoritzLaurer/xtremedistil-l6-h256-zeroshot-v1.1-all-33", device=device, multi_label=True)
        print("Step 1.1: Classification model loaded.")
    except Exception as e:
        print(f"{Fore.LIGHTBLACK_EX}WARNING: Error loading classification model: {e}{Style.RESET_ALL}"); classifier = None

def run_summarization_in_process(summarizer_pipeline, chunk, queue, max_len, min_len):
    try:
        summary = summarizer_pipeline(chunk, max_length=max_len, min_length=min_len, truncation=True)
        queue.put(summary[0]['summary_text'])
    except Exception as e:
        queue.put(f"Summarization failed in subprocess: {e}")

# --- MODIFICATION START: Add chunk parsing function
def parse_chunks(chunks_string, file_path, max_chunk_length):
    """Parses a comma-separated string of chunks and ranges into a list of chunks."""
    if not chunks_string.strip():
        # Default to a single chunk if no specific chunks are provided
        print(f"No chunks specified for {os.path.basename(file_path)}. Processing the first chunk.")
        return [1]

    chunks = set()
    parts = chunks_string.split(',')
    for part in parts:
        part = part.strip()
        if not part: continue
        if '-' in part:
            start_str, end_str = part.split('-')
            start, end = int(start_str), int(end_str)
            if start <= end:
                chunks.update(range(start, end + 1))
        else:
            chunks.add(int(part))
    return sorted(list(chunks))
# --- MODIFICATION END ---

def read_chunks_from_file(file_path, chunk_numbers, max_chunk_length):
    ext = os.path.splitext(file_path)[1].lower()
    all_words = []
    
    try:
        if ext == ".pdf":
            with fitz.open(file_path) as doc:
                words = []
                for page in doc:
                    words.extend(page.get_text().split())
                for chunk_num in chunk_numbers:
                    start_index = (chunk_num - 1) * max_chunk_length
                    end_index = start_index + max_chunk_length
                    all_words.extend(words[start_index:end_index])
        elif ext == ".docx":
            doc = Document(file_path)
            words = []
            for para in doc.paragraphs: words.extend(para.text.split())
            for chunk_num in chunk_numbers:
                start_index = (chunk_num - 1) * max_chunk_length
                end_index = start_index + max_chunk_length
                all_words.extend(words[start_index:end_index])
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as f: words = f.read().split()
            for chunk_num in chunk_numbers:
                start_index = (chunk_num - 1) * max_chunk_length
                end_index = start_index + max_chunk_length
                all_words.extend(words[start_index:end_index])
        elif ext == ".xlsx":
            words = []
            for sheet in pd.ExcelFile(file_path).sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                words.extend(df.to_string(index=False).split())
            for chunk_num in chunk_numbers:
                start_index = (chunk_num - 1) * max_chunk_length
                end_index = start_index + max_chunk_length
                all_words.extend(words[start_index:end_index])
        elif ext == ".pptx":
            prs = Presentation(file_path)
            words = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        words.extend(shape.text.split())
            for chunk_num in chunk_numbers:
                start_index = (chunk_num - 1) * max_chunk_length
                end_index = start_index + max_chunk_length
                all_words.extend(words[start_index:end_index])
    except Exception as e:
        print(f"{Fore.LIGHTBLACK_EX}WARNING: Error reading {ext} file: {e}{Style.RESET_ALL}"); return ""
    return " ".join(all_words)


def summarize_text(text, max_chunk_length, max_summary_length, min_summary_length, timeout_seconds, show_verbiage):
    if not text.strip(): return ["No text to summarize."]
    words = text.split()
    chunks = [" ".join(words[i:i + max_chunk_length]) for i in range(0, len(words), max_chunk_length)]
    summaries = []
    for i, chunk in enumerate(chunks):
        if show_verbiage:
            print(f"\n--- Chunk {i+1} Verbiage ---")
            # --- MODIFICATION START: Colorize verbiage output
            print(f"{Fore.MAGENTA}{chunk}{Style.RESET_ALL}")
            # --- MODIFICATION END ---
            print("----------------------------\n")
        print(f"Step 7.1.1: Summarizing chunk {i + 1} of {len(chunks)}...")
        q = multiprocessing.Queue()
        p = multiprocessing.Process(target=run_summarization_in_process, args=(summarizer, chunk, q, max_summary_length, min_summary_length))
        p.start()
        p.join(timeout=timeout_seconds)
        if p.is_alive():
            print(f"{Fore.RED}Step 7.1.3: Summarization for chunk {i + 1} timed out after {timeout_seconds} seconds.")
            p.terminate(); p.join()
            summaries.append("Summary failed due to timeout.")
        else:
            try:
                result = q.get(timeout=5)
                summaries.append(result)
                if show_verbiage:
                    print(f"\n--- Summary for Chunk {i+1} ---")
                    # --- MODIFICATION START: Colorize summary text output
                    print(f"{Fore.MAGENTA}{result}{Style.RESET_ALL}")
                    # --- MODIFICATION END ---
                    print("------------------------------\n")
                print(f"Step 7.1.2: Summarization for chunk {i + 1} complete.")
            except Exception:
                summaries.append("Summary failed to return result from process.")
                print(f"{Fore.LIGHTBLACK_EX}WARNING: Error retrieving result for chunk {i + 1}.{Style.RESET_ALL}")
    return ". ".join(summaries).split(". ")

def categorize_summary(summary_text, categories):
    if classifier is None: return {'labels': ['Other'], 'scores': [1.0]}
    print("Step 8.1: Sending summary to classifier model...")
    try:
        results = classifier(summary_text, candidate_labels=categories, multi_label=True)
        # --- MODIFICATION START: Colorize classifier output
        print(f"{Fore.GREEN}Step 8.2: Categorization complete.{Style.RESET_ALL}")
        # --- MODIFICATION END ---
        return results
    except Exception as e:
        print(f"{Fore.LIGHTBLACK_EX}WARNING: Error during AI categorization: {e}{Style.RESET_ALL}"); return {'labels': ['Other'], 'scores': [1.0]}

def tag_file_properties(file_path, primary_category, secondary_categories):
    if not os.access(file_path, os.W_OK):
        print(f"{Fore.RED}Tagging skipped: No write permission for '{os.path.basename(file_path)}'.")
        return
    ext = os.path.splitext(file_path)[1].lower()
    if not os.path.exists(file_path):
        print(f"{Fore.LIGHTBLACK_EX}WARNING: Tagging skipped: File not found at {file_path}{Style.RESET_ALL}"); return
    keyword_parts = []
    if primary_category and primary_category != "Other": keyword_parts.append(f"AILFO Primary:; {primary_category}")
    for cat in secondary_categories: keyword_parts.append(f"2nd:; {cat}")
    keyword_string = ":: ".join(keyword_parts)
    if not keyword_string:
        print(f"No valid categories to tag for {os.path.basename(file_path)}"); return
    try:
        obj = None
        if ext == ".pdf":
            doc = fitz.open(file_path)
            doc.metadata["keywords"] = keyword_string
            doc.save(file_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP); doc.close()
        elif ext in [".docx", ".pptx", ".xlsx"]:
            if ext == ".docx": obj = Document(file_path)
            elif ext == ".pptx": obj = Presentation(file_path)
            elif ext == ".xlsx": obj = load_workbook(file_path)
            if ext != ".xlsx": obj.core_properties.keywords = keyword_string
            else: obj.properties.keywords = keyword_string
            obj.save(file_path)
        else:
            print(f"Tagging not supported for {ext} files."); return
        print(f"{Fore.CYAN}Tagged '{os.path.basename(file_path)}' with: {keyword_string}")
    except Exception as e:
        print(f"{Fore.LIGHTBLACK_EX}WARNING: Error tagging file {os.path.basename(file_path)}: {e}{Style.RESET_ALL}")

def create_shortcut(source_path, shortcut_path):
    if sys.platform == "win32":
        shortcut_path += ".url"
        if os.path.exists(shortcut_path): return
        with open(shortcut_path, "w") as f:
            f.write(f"[InternetShortcut]\nURL=file:///{os.path.abspath(source_path)}\n")
    else:
        if os.path.lexists(shortcut_path): return
        os.symlink(source_path, shortcut_path)

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
            
            if file_action != 'shortcut':
                destination_path_with_ext = destination_path + os.path.splitext(file_path)[1]
            else:
                destination_path_with_ext = destination_path
            
            try:
                if file_action == 'shortcut':
                    create_shortcut(file_path, destination_path)
                    print(f"{Fore.CYAN}Pre-processed shortcut for: {os.path.basename(file_path)} -> {destination_folder}")
                else:
                    if file_action == 'move':
                        shutil.move(file_path, destination_path_with_ext)
                        print(f"{Fore.CYAN}Pre-processed and moved: {os.path.basename(file_path)} -> {destination_folder}")
                    elif file_action == 'copy':
                        shutil.copy2(file_path, destination_path_with_ext)
                        print(f"{Fore.CYAN}Pre-processed and copied: {os.path.basename(file_path)} -> {destination_folder}")
                remaining_files.remove(file_path)
            except Exception as e:
                print(f"{Fore.LIGHTBLACK_EX}WARNING: Error pre-processing {os.path.basename(file_path)}: {e}{Style.RESET_ALL}")
    print(f"--- Pre-processing Complete: {len(file_list) - len(remaining_files)} files processed. ---")
    return remaining_files

def process_files_in_folder(file_list, categories, file_management_settings, settings):
    if not file_list: print("\nNo remaining files to process."); return []
    print(f"\n--- Starting Main AI Processing for {len(file_list)} files... ---")
    all_summaries = []
    for i, file_path in enumerate(file_list):
        print(f"\n--- AI Processing file {i + 1}/{len(file_list)}: {os.path.basename(file_path)} ---")
        if settings['file_action_choice'] != 'none':
            file_name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
            cleaned_file_name_for_skip = file_name_no_ext.replace('_', ' ').replace('-', ' ')
            already_processed = False
            for category in categories:
                if category == "Other": continue
                prefix = file_management_settings.get(category, {}).get('prefix', '')
                new_file_name_no_ext = f"{prefix}_{cleaned_file_name_for_skip}" if prefix else cleaned_file_name_for_skip
                if settings['file_action_choice'] == 'shortcut':
                    expected_output_name = new_file_name_no_ext + ".url" if sys.platform == "win32" else new_file_name_no_ext
                else:
                    expected_output_name = new_file_name_no_ext + os.path.splitext(file_path)[1]
                sanitized_cat = sanitize_for_path(category)
                potential_dest_path = os.path.join(settings['output_folder_path'], sanitized_cat, expected_output_name)
                if os.path.exists(potential_dest_path):
                    print(f"Skipping: Output file '{expected_output_name}' appears to exist in '{sanitized_cat}' subfolder.")
                    already_processed = True; break
            if already_processed: continue

        is_pdf = os.path.splitext(file_path)[1].lower() == '.pdf'
        is_text = os.path.splitext(file_path)[1].lower() == '.txt'
        
        # --- MODIFICATION START: Determine chunk settings based on file type and size
        max_chunk_length = settings['generic_chunk_length']
        chunks_list = []
        
        if is_pdf:
            try:
                with fitz.open(file_path) as doc:
                    page_count = doc.page_count
                    if page_count > settings['long_pdf_threshold']:
                        max_chunk_length = settings['long_pdf_chunk_length']
                        chunks_list = parse_chunks(settings['long_pdf_chunks_list'], file_path, max_chunk_length)
                        print(f"Detected long PDF ({page_count} pages). Using long PDF settings.")
                    else:
                        max_chunk_length = settings['short_pdf_chunk_length']
                        chunks_list = parse_chunks(settings['short_pdf_chunks_list'], file_path, max_chunk_length)
                        print(f"Detected short PDF ({page_count} pages). Using short PDF settings.")
            except Exception as e:
                print(f"{Fore.LIGHTBLACK_EX}WARNING: Error getting PDF page count: {e}. Using default PDF settings.{Style.RESET_ALL}")
                max_chunk_length = settings['pdf_chunk_length']
                chunks_list = parse_chunks(settings['pdf_chunks_list'], file_path, max_chunk_length)
        elif is_text:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                    word_count = len(content.split())
                    if word_count > settings['long_text_threshold']:
                        max_chunk_length = settings['long_text_chunk_length']
                        chunks_list = parse_chunks(settings['long_text_chunks_list'], file_path, max_chunk_length)
                        print(f"Detected long text file ({word_count} words). Using long text settings.")
                    else:
                        max_chunk_length = settings['short_text_chunk_length']
                        chunks_list = parse_chunks(settings['short_text_chunks_list'], file_path, max_chunk_length)
                        print(f"Detected short text file ({word_count} words). Using short text settings.")
            except Exception as e:
                print(f"{Fore.LIGHTBLACK_EX}WARNING: Error counting words in text file: {e}. Using default text settings.{Style.RESET_ALL}")
                max_chunk_length = settings['generic_chunk_length']
                chunks_list = parse_chunks(settings['generic_chunks_list'], file_path, max_chunk_length)
        else:
            max_chunk_length = settings['generic_chunk_length']
            chunks_list = parse_chunks(settings['generic_chunks_list'], file_path, max_chunk_length)
        # --- MODIFICATION END ---


        # --- MODIFICATION START: Read specific chunks
        if not chunks_list:
            print(f"{Fore.LIGHTBLACK_EX}WARNING: No chunks specified for this file type and size. Skipping.{Style.RESET_ALL}")
            continue
        
        chunk_string = ", ".join(map(str, chunks_list))
        print(f"Step 6: Extracting from chunk(s): {chunk_string}...")
        text = read_chunks_from_file(file_path, chunks_list, max_chunk_length)
        if not text.strip():
            print(f"{Fore.LIGHTBLACK_EX}WARNING: No text found in specified chunks, skipping.{Style.RESET_ALL}")
            continue
        print("Step 6.1: Text extracted successfully.")
        # --- MODIFICATION END ---
        
        print("Step 7: Text extracted successfully. Starting summarization...")

        max_summary_length_final = settings['max_summary_length']
        min_summary_length_final = settings['min_summary_length']
        if is_pdf:
            try:
                with fitz.open(file_path) as doc:
                    page_count = doc.page_count
                    if page_count > settings['long_pdf_threshold']:
                        max_summary_length_final = settings['long_pdf_max_summary_length']
                        min_summary_length_final = settings['long_pdf_min_summary_length']
                    else:
                        max_summary_length_final = settings['short_pdf_max_summary_length']
                        min_summary_length_final = settings['short_pdf_min_summary_length']
            except Exception as e:
                print(f"{Fore.LIGHTBLACK_EX}WARNING: Error getting PDF page count: {e}. Using default summary settings.{Style.RESET_ALL}")
        
        bullet_points = summarize_text(text, max_chunk_length, max_summary_length_final, min_summary_length_final, settings['summarizer_timeout'], 

settings['show_chunk_verbiage'])
        if not bullet_points or "Summary failed" in bullet_points[0]: print(f"{Fore.LIGHTBLACK_EX}WARNING: Summarization failed.{Style.RESET_ALL}"); 

continue

        print("\n--- Plain Text Summary ---")
        summary_text_for_file = ""
        # --- MODIFICATION START: Colorize the plain text summary output
        for point in bullet_points:
            p_strip = point.strip()
            if p_strip:
                print(f"{Fore.LIGHTGREEN_EX}- {p_strip}{Style.RESET_ALL}")
                summary_text_for_file += f"- {p_strip}\n"
        # --- MODIFICATION END ---

        cleaned_file_name = os.path.splitext(os.path.basename(file_path))[0].replace('_', ' ').replace('-', ' ')
        
        # --- MODIFICATION START: Capture original category before it's potentially changed to "Other"
        results = categorize_summary(". ".join([cleaned_file_name] + bullet_points), categories)
        original_primary_category = results['labels'][0]
        original_primary_score = results['scores'][0]
        # --- MODIFICATION END ---
        
        primary_category, secondary_categories, all_valid_categories = "Other", [], []
        if original_primary_score * 100 >= settings['confidence_threshold']:
            primary_category = original_primary_category
            all_valid_categories.append(primary_category)
            for label, score in zip(results['labels'][1:], results['scores'][1:]):
                if score * 100 >= settings['secondary_confidence_threshold']:
                    secondary_categories.append((label, score))
                    all_valid_categories.append(label)

        print("Step 9: Final summary complete.")

        if settings['tagging_enabled_choice'] in ['y', 'yes'] and all_valid_categories:
            sec_cat_names = [cat for cat, score in secondary_categories]
            tag_file_properties(file_path, primary_category, sec_cat_names)

        print(f"\nFile: {file_path}")
        
        # --- MODIFICATION START: Show original and final category if changed
        category_color = Fore.YELLOW if primary_category == "Other" else Fore.GREEN
        if primary_category == "Other" and original_primary_category != "Other":
            print(f"Primary Category: {category_color}{Style.BRIGHT}{primary_category} (was '{original_primary_category}' at 

{original_primary_score*100:.2f}%)")
        else:
            print(f"Primary Category: {category_color}{Style.BRIGHT}{primary_category} ({results['scores'][0]*100:.2f}%)")
        # --- MODIFICATION END ---

        if secondary_categories:
            colors = [Fore.CYAN, Fore.LIGHTGREEN_EX]
            colored_cats = [f"{colors[i % 2]}{cat} ({Fore.YELLOW}{score*100:.2f}%)" for i, (cat, score) in enumerate(secondary_categories)]
            print(f"Secondary Categories: {', '.join(colored_cats)}")

        print(f"\n{Fore.CYAN}--- Colorized Summary ---{Style.RESET_ALL}")
        if primary_category == "Other":
            print(f"{Fore.MAGENTA}{summary_text_for_file}")
            continue

        sec_cat_map = {cat.lower(): i for i, (cat, score) in enumerate(secondary_categories)}
        colors = [Fore.CYAN, Fore.LIGHTGREEN_EX]
        for point in bullet_points:
            p_strip = point.strip()
            if not p_strip: continue
            colorized_point = ""
            words_and_delimiters = re.split(r'([ ,.;:!?])', p_strip)
            for word in words_and_delimiters:
                if word.lower() == primary_category.lower(): colorized_point += f"{Fore.GREEN}{Style.BRIGHT}{word}{Style.RESET_ALL}"
                elif word.lower() in sec_cat_map:
                    idx = sec_cat_map[word.lower()]
                    color = colors[idx % 2]
                    colorized_point += f"{color}{word}{Style.RESET_ALL}"
                else: colorized_point += word
            print(f"{Fore.LIGHTCYAN_EX}- {colorized_point}{Style.RESET_ALL}")
            
        all_summaries.append({'file_path': os.path.abspath(file_path), 'primary_category': primary_category, 'all_categories': all_valid_categories, 

'summary': summary_text_for_file})

        if settings['file_action_choice'] != 'none' and all_valid_categories:
            for category in all_valid_categories:
                sanitized_cat = sanitize_for_path(category)
                destination_folder = os.path.join(settings['output_folder_path'], sanitized_cat)
                os.makedirs(destination_folder, exist_ok=True)
                prefix = file_management_settings.get(category, {}).get('prefix', '')
                new_file_name = f"{prefix}_{cleaned_file_name}" if prefix else cleaned_file_name
                
                destination_path = os.path.join(destination_folder, new_file_name)
                
                if settings['file_action_choice'] != 'shortcut':
                    destination_path_with_ext = destination_path + os.path.splitext(file_path)[1]
                else:
                    destination_path_with_ext = destination_path
                
                try:
                    if settings['file_action_choice'] == 'shortcut':
                        create_shortcut(file_path, destination_path)
                        print(f"{Fore.CYAN}Created shortcut in -> {destination_folder}")
                    elif category == primary_category:
                        if settings['file_action_choice'] == 'move' and os.path.exists(file_path):
                            shutil.move(file_path, destination_path_with_ext)
                            print(f"{Fore.CYAN}Moved and renamed: {file_path} -> {destination_path_with_ext}")
                            if settings['tagging_enabled_choice'] in ['y', 'yes']:
                                sec_cat_names = [cat for cat, score in secondary_categories]
                                tag_file_properties(destination_path_with_ext, primary_category, sec_cat_names)
                        elif settings['file_action_choice'] == 'copy':
                            shutil.copy2(file_path, destination_path_with_ext)
                            print(f"{Fore.CYAN}Copied and renamed: {file_path} -> {destination_path_with_ext}")
                            if settings['tagging_enabled_choice'] in ['y', 'yes']:
                                sec_cat_names = [cat for cat, score in secondary_categories]
                                tag_file_properties(destination_path_with_ext, primary_category, sec_cat_names)
                except Exception as e: print(f"Error during file operation: {e}")
    return all_summaries

# --- Main Execution ---
if __name__ == "__main__":
    multiprocessing.freeze_support()
    initialize_models()
    try:
        CATEGORIES = load_and_edit_categories()
        settings_file = "settings.json"
        defaults = {
            'pdf_chunk_length': 2048, 'generic_chunk_length': 250, 'max_summary_length': 150, 'min_summary_length': 40,
            'pdf_padding': 0, 'pdf_chunks': 1, 'generic_padding': 0, 'generic_chunks': 1,
            'confidence_threshold': 50, 'secondary_confidence_threshold': 20, 'tagging_enabled_choice': 'n',
            'add_prefix_choice': 'n', 'global_prefix_choice': 'n', 'global_prefix': '', 'special_char': '',
            'selected_file_types': ['.pdf', '.docx', '.txt', '.xlsx', '.pptx'], 'file_action_choice': 'none',
            'last_list_file': '', 'last_destination_folder': '', 'source_choice': 'f', 'preprocess_choice': 'n',
            'summarizer_timeout': 60,
            'long_pdf_threshold': 10,
            'long_text_threshold': 1000,
            'short_pdf_chunk_length': 1025,
            'long_pdf_chunk_length': 4096,
            'short_text_chunk_length': 1025,
            'long_text_chunk_length': 4096,
            'long_pdf_max_summary_length': 200,
            'long_pdf_min_summary_length': 50,
            'short_pdf_max_summary_length': 150,
            'short_pdf_min_summary_length': 40,
            'short_pdf_chunks_list': '',
            'long_pdf_chunks_list': '',
            'short_text_chunks_list': '',
            'long_text_chunks_list': '',
            'generic_chunks_list': '',
            'show_chunk_verbiage': 'n'
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
        settings_to_save['summarizer_timeout'] = int(input(f"Enter summarizer timeout in seconds [default: {settings_to_save['summarizer_timeout']}]: 

").strip() or settings_to_save['summarizer_timeout'])
        settings_to_save['confidence_threshold'] = int(input(f"Enter primary confidence threshold (0-100) [default: {settings_to_save

['confidence_threshold']}]: ").strip() or settings_to_save['confidence_threshold'])
        settings_to_save['secondary_confidence_threshold'] = int(input(f"Enter secondary confidence threshold (0-100) [default: {settings_to_save

['secondary_confidence_threshold']}]: ").strip() or settings_to_save['secondary_confidence_threshold'])

        settings_to_save['generic_chunk_length'] = int(input(f"Enter non-PDF chunk length [default: {settings_to_save['generic_chunk_length']}]: ").strip

() or settings_to_save['generic_chunk_length'])
        settings_to_save['short_pdf_chunk_length'] = int(input(f"Enter SHORT PDF chunk length [default: {settings_to_save.get('short_pdf_chunk_length', 

defaults['short_pdf_chunk_length'])}]: ").strip() or settings_to_save.get('short_pdf_chunk_length', defaults['short_pdf_chunk_length']))
        settings_to_save['long_pdf_chunk_length'] = int(input(f"Enter LONG PDF chunk length [default: {settings_to_save.get('long_pdf_chunk_length', 

defaults['long_pdf_chunk_length'])}]: ").strip() or settings_to_save.get('long_pdf_chunk_length', defaults['long_pdf_chunk_length']))
        
        settings_to_save['max_summary_length'] = int(input(f"Enter SHORT PDF max summary length [default: {settings_to_save['max_summary_length']}]: 

").strip() or settings_to_save['max_summary_length'])
        settings_to_save['min_summary_length'] = int(input(f"Enter SHORT PDF min summary length [default: {settings_to_save['min_summary_length']}]: 

").strip() or settings_to_save['min_summary_length'])
        settings_to_save['long_pdf_threshold'] = int(input(f"Enter the LONG PDF threshold (in pages) [default: {settings_to_save['long_pdf_threshold']}]: 

").strip() or settings_to_save['long_pdf_threshold'])
        settings_to_save['long_pdf_max_summary_length'] = int(input(f"Enter LONG PDF max summary length [default: {settings_to_save

['long_pdf_max_summary_length']}]: ").strip() or settings_to_save['long_pdf_max_summary_length'])
        settings_to_save['long_pdf_min_summary_length'] = int(input(f"Enter LONG PDF min summary length [default: {settings_to_save

['long_pdf_min_summary_length']}]: ").strip() or settings_to_save['long_pdf_min_summary_length'])
        
        settings_to_save['tagging_enabled_choice'] = input(f"Tag files with category keywords? (y/n) [default: {settings_to_save

['tagging_enabled_choice']}]: ").strip().lower() or settings_to_save['tagging_enabled_choice']

        print("\n--- Chunk Selection (comma-separated list/ranges) ---")
        settings_to_save['short_pdf_chunks_list'] = input(f"Enter chunks for SHORT PDFs (e.g., '1, 3-5') [default: {settings_to_save.get

('short_pdf_chunks_list', '')}]: ").strip() or settings_to_save.get('short_pdf_chunks_list', '')
        settings_to_save['long_pdf_chunks_list'] = input(f"Enter chunks for LONG PDFs (e.g., '1-2, 10') [default: {settings_to_save.get

('long_pdf_chunks_list', '')}]: ").strip() or settings_to_save.get('long_pdf_chunks_list', '')
        settings_to_save['short_text_chunks_list'] = input(f"Enter chunks for SHORT text files (e.g., '1') [default: {settings_to_save.get

('short_text_chunks_list', '')}]: ").strip() or settings_to_save.get('short_text_chunks_list', '')
        settings_to_save['long_text_chunks_list'] = input(f"Enter chunks for LONG text files (e.g., '1-5') [default: {settings_to_save.get

('long_text_chunks_list', '')}]: ").strip() or settings_to_save.get('long_text_chunks_list', '')
        settings_to_save['long_text_threshold'] = int(input(f"Enter the LONG TEXT threshold (in words) [default: {settings_to_save

['long_text_threshold']}]: ").strip() or settings_to_save['long_text_threshold'])
        settings_to_save['short_text_chunk_length'] = int(input(f"Enter SHORT text chunk length [default: {settings_to_save['short_text_chunk_length']}]: 

").strip() or settings_to_save['short_text_chunk_length'])
        settings_to_save['long_text_chunk_length'] = int(input(f"Enter LONG text chunk length [default: {settings_to_save['long_text_chunk_length']}]: 

").strip() or settings_to_save['long_text_chunk_length'])
        
        current_show_verbiage = settings_to_save.get('show_chunk_verbiage', 'n')
        settings_to_save['show_chunk_verbiage'] = input(f"Show extraction and summary text for each chunk? (y/n) [default: {current_show_verbiage}]: 

").strip().lower() or current_show_verbiage
        
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
        if source_choice == 'l': settings_to_save['last_list_file'] = list_file_path
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
