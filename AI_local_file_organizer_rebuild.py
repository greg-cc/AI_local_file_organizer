import os
import fitz  # PyMuPDF for PDF
from docx import Document
import pandas as pd
from pptx import Presentation
from transformers import pipeline
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import shutil
import json
import colorama
from colorama import Fore, Style
import re

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
# These are now set via user input in the main execution block.

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
                print("\n--- Current Categories ---")
                for i, cat in enumerate(categories):
                    print(f"[{i+1}] {cat}")
                print("--------------------------")
                print(instructions)
            elif action == 'remove' and len(parts) == 2 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    removed_cat = categories.pop(idx)
                    print(f"Removed: {removed_cat}")
                    print("\n--- Current Categories ---")
                    for i, cat in enumerate(categories):
                        print(f"[{i+1}] {cat}")
                    print("--------------------------")
                    print(instructions)
                else:
                    print("Invalid category number.")
            elif action == 'edit' and len(parts) >= 3 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    new_cat = " ".join(parts[2:])
                    old_cat = categories[idx]
                    categories[idx] = new_cat
                    print(f"Edited: '{old_cat}' to '{new_cat}'")
                    print("\n--- Current Categories ---")
                    for i, cat in enumerate(categories):
                        print(f"[{i+1}] {cat}")
                    print("--------------------------")
                    print(instructions)
                else:
                    print("Invalid category number.")
            elif action == 'list':
                print("\n--- Current Categories ---")
                for i, cat in enumerate(categories):
                    print(f"[{i+1}] {cat}")
                print("--------------------------")
                print(instructions)
            else:
                print("Invalid command. Please try again.")
    
    # Save the final list of categories
    with open(category_file, 'w', encoding='utf-8') as f:
        json.dump(categories, f, indent=4)
    print("\nStep 1.3: Categories saved for next session.")
    
    return categories

# --- Load Models ---
print("\nStep 1: Loading summarization and classification models...")
print("If this is the first time you are running the script, a large file download will begin now. Please wait for it to complete.")

# Load the SMALLER summarization pipeline for speed
summarizer = pipeline(
    "summarization",
    model="t5-small", 
    max_length=1024, 
    device=device,
    framework="pt"  # Explicitly tell the pipeline to use PyTorch
)

# Load a dedicated model for text classification
try:
    classifier = pipeline(
        "zero-shot-classification",
        model="MoritzLaurer/xtremedistil-l6-h256-zeroshot-v1.1-all-33",
        device=device
    )
    print("Step 1.1: Classification model loaded successfully.")
except Exception as e:
    print(f"Error loading classification model: {e}")
    classifier = None
    print("Step 1.1: Falling back to 'Other' for all classifications.")


# --- File Reading Functions ---
def read_chunks_from_file(file_path, start_chunk, end_chunk, max_chunk_length):
    """
    Efficiently reads a specific range of word chunks from any file type.
    It stops reading the file once the desired number of chunks has been collected.
    """
    ext = os.path.splitext(file_path)[1].lower()
    words = []
    total_words_needed = end_chunk * max_chunk_length

    try:
        if ext == ".pdf":
            with fitz.open(file_path) as doc:
                for page in doc:
                    words.extend(page.get_text().split())
                    if len(words) >= total_words_needed:
                        break
        elif ext == ".docx":
            doc = Document(file_path)
            for para in doc.paragraphs:
                words.extend(para.text.split())
                if len(words) >= total_words_needed:
                    break
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as f:
                # This is less efficient for huge txt files, but is the standard way
                words = f.read().split()
        elif ext == ".xlsx":
            sheets = pd.ExcelFile(file_path).sheet_names
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet)
                words.extend(df.to_string(index=False).split())
                if len(words) >= total_words_needed:
                    break
        elif ext == ".pptx":
            presentation = Presentation(file_path)
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        words.extend(shape.text.split())
                        if len(words) >= total_words_needed:
                            break
                if len(words) >= total_words_needed:
                    break
    except Exception as e:
        print(f"Error reading {ext} file: {e}")
        return ""

    start_index = (start_chunk - 1) * max_chunk_length
    end_index = end_chunk * max_chunk_length
    
    # Return the requested slice of words, joined back into a string
    return " ".join(words[start_index:end_index])

# --- Summarization Function ---
def summarize_text(text, max_chunk_length, max_summary_length, min_summary_length):
    """
    Generates a summary from the provided text by processing it in word chunks.
    """
    if not text.strip():
        return ["No text to summarize."]

    words = text.split()
    chunks = [" ".join(words[i:i + max_chunk_length]) for i in range(0, len(words), max_chunk_length)]
    
    summaries = []
    
    for i, chunk in enumerate(chunks):
        print(f"Step 7.1.1: Summarizing chunk {i + 1} of {len(chunks)}...")
        try:
            summary = summarizer(
                chunk, 
                max_length=max_summary_length, 
                min_length=min_summary_length, 
                truncation=True
            )
            summaries.append(summary[0]['summary_text'])
            print(f"Step 7.1.2: Summarization for chunk {i + 1} complete.")
        except Exception as e:
            print(f"Step 7.1.3: Error during summarization of chunk {i + 1}: {e}")
            summaries.append("Summary failed for this chunk.")
    
    combined_summary = ". ".join(summaries)
    return combined_summary.split(". ")
    
# --- Categorization Function ---
def categorize_summary(summary_text, categories):
    """
    Categorizes a summary using a zero-shot classification model.
    Returns the full results dictionary from the classifier.
    """
    if classifier is None:
        return {'labels': ['Other'], 'scores': [1.0]}

    print("Step 8.1: Sending summary to classifier model...")
    try:
        results = classifier(summary_text, candidate_labels=categories)
        print("Step 8.2: Categorization complete.")
        return results
    except Exception as e:
        print(f"Step 8.3: Error during AI categorization: {e}")
        return {'labels': ['Other'], 'scores': [1.0]}

# --- NEW: Pre-processing Function ---
def preprocess_by_filename(file_list, categories, file_management_settings):
    """
    Scans filenames for category words and moves them before full processing.
    Returns a list of files that were NOT moved.
    """
    print("\n--- Starting Pre-processing Step ---")
    moved_files = set()
    remaining_files = []
    
    for file_path in file_list:
        file_name_no_ext = os.path.splitext(os.path.basename(file_path))[0]
        cleaned_file_name = file_name_no_ext.replace('_', ' ').replace('-', ' ')
        words_in_name = {word.lower() for word in cleaned_file_name.split()}
        
        found_category = None
        # --- CORRECTED: Improved matching logic for multi-word categories ---
        for cat in categories: 
            category_words = set(cat.lower().split())
            if category_words.issubset(words_in_name):
                found_category = cat
                break # Found a match, stop searching
        
        if found_category and found_category in file_management_settings:
            settings = file_management_settings[found_category]
            prefix = settings['prefix']
            destination_folder = settings['destination']
            
            os.makedirs(destination_folder, exist_ok=True)

            original_file_extension = os.path.splitext(file_path)[1]
            original_file_name_no_ext = os.path.basename(os.path.splitext(file_path)[0])
            
            if prefix:
                new_file_name = f"{prefix}_{original_file_name_no_ext}{original_file_extension}"
            else:
                new_file_name = f"{original_file_name_no_ext}{original_file_extension}"

            destination_file_path = os.path.join(destination_folder, new_file_name)
            
            try:
                shutil.move(file_path, destination_file_path)
                print(f"Pre-processed and moved: {os.path.basename(file_path)} -> {destination_folder}")
                moved_files.add(file_path)
            except Exception as move_e:
                print(f"Error pre-processing and moving {os.path.basename(file_path)}: {move_e}")
                remaining_files.append(file_path) # If move fails, keep it for main processing
        else:
            remaining_files.append(file_path)

    print(f"--- Pre-processing Complete: {len(moved_files)} files moved. ---")
    return remaining_files


# --- Main Processing Logic ---
def process_files_in_folder(file_list, categories, pdf_padding, pdf_chunks, generic_padding, generic_chunks, file_management_settings, confidence_threshold, max_chunk_length, max_summary_length, min_summary_length):
    """
    Walks a folder, processes supported files, and generates summaries.
    Includes detailed progress checks.
    """
    if not file_list:
        print("\nNo remaining files to process after pre-processing step.")
        return []

    print(f"\nStep 4: Starting main processing for {len(file_list)} remaining files...")

    all_summaries = []

    for i, file_path in enumerate(file_list):
        print(f"\nStep 5: Processing file {i + 1}/{len(file_list)} - {os.path.basename(file_path)}")
        
        is_pdf = os.path.splitext(file_path)[1].lower() == '.pdf'
        
        if is_pdf:
            padding = pdf_padding
            chunks_to_process = pdf_chunks
        else:
            padding = generic_padding
            chunks_to_process = generic_chunks
            
        start_chunk = padding + 1
        end_chunk = padding + chunks_to_process
        
        print(f"\nStep 6: Extracting from chunk(s) {start_chunk} to {end_chunk}...")
        text = read_chunks_from_file(file_path, start_chunk, end_chunk, max_chunk_length)

        if text.strip():
            print("Step 7: Text extracted successfully. Starting summarization...")
            bullet_points = summarize_text(text, max_chunk_length, max_summary_length, min_summary_length)
            
            if bullet_points and bullet_points != ["No text to summarize."]:
                file_name_for_analysis = os.path.basename(file_path)
                file_name_no_ext = os.path.splitext(file_name_for_analysis)[0]
                cleaned_file_name = file_name_no_ext.replace('_', ' ').replace('-', ' ')
                
                all_chunks_for_classifier = [cleaned_file_name] + bullet_points
                text_for_classifier = ". ".join(all_chunks_for_classifier)
                
                print("Step 8: Categorizing summary offline...")
                results = categorize_summary(text_for_classifier, categories)
                
                top_score = results['scores'][0]
                
                if top_score * 100 >= confidence_threshold:
                    top_category = results['labels'][0]
                    other_categories = results['labels'][1:]
                else:
                    top_category = "Other"
                    other_categories = results['labels']
                    print(f"Top category '{results['labels'][0]}' with score {top_score*100:.2f}% is below threshold of {confidence_threshold}%. Assigning to 'Other'.")

                print("Step 9: Final summary complete.")
                
                print(f"\nFile: {file_path}")
                print("Summary:\n")
                
                print(f"Category: {Fore.GREEN}{Style.BRIGHT}{top_category} ({top_score*100:.2f}%)")
                
                summary_text_for_file = ""
                for point in bullet_points:
                    colorized_point = ""
                    words_and_delimiters = re.split(r'([ ,.;:!?])', point)
                    
                    for word in words_and_delimiters:
                        if word.lower() == top_category.lower():
                            colorized_point += f"{Fore.GREEN}{Style.BRIGHT}{word}{Style.RESET_ALL}"
                        elif word.lower() in [cat.lower() for cat in other_categories]:
                            colorized_point += f"{Fore.YELLOW}{word}{Style.RESET_ALL}"
                        else:
                            colorized_point += word
                    
                    print(f"- {colorized_point.strip()}")
                    summary_text_for_file += f"- {point.strip()}\n"
                
                colorized_filename = ""
                filename_words_and_delimiters = re.split(r'([ ,.;:!?])', cleaned_file_name)
                for word in filename_words_and_delimiters:
                    if word.lower() == top_category.lower():
                        colorized_filename += f"{Fore.GREEN}{Style.BRIGHT}{word}{Style.RESET_ALL}"
                    elif word.lower() in [cat.lower() for cat in other_categories]:
                        colorized_filename += f"{Fore.YELLOW}{word}{Style.RESET_ALL}"
                    else:
                        colorized_filename += word
                print(f"- {colorized_filename.strip()}")
                
                print("\n" + "-"*50)

                all_summaries.append({
                    'file_path': os.path.abspath(file_path),
                    'category': top_category,
                    'summary': summary_text_for_file
                })

                if top_category in file_management_settings and file_management_settings[top_category].get('destination'):
                    settings = file_management_settings[top_category]
                    prefix = settings['prefix']
                    destination_folder = settings['destination']
                    
                    os.makedirs(destination_folder, exist_ok=True)

                    original_file_extension = os.path.splitext(file_path)[1]
                    original_file_name_no_ext = os.path.basename(os.path.splitext(file_path)[0])
                    
                    if prefix:
                        new_file_name = f"{prefix}_{original_file_name_no_ext}{original_file_extension}"
                    else:
                        new_file_name = f"{original_file_name_no_ext}{original_file_extension}"

                    destination_file_path = os.path.join(destination_folder, new_file_name)
                    
                    if os.path.exists(file_path):
                        try:
                            shutil.move(file_path, destination_file_path)
                            print(f"Moved and renamed: {file_path} -> {destination_file_path}")
                        except Exception as move_e:
                            print(f"Error moving original file: {move_e}")
            else:
                print("Step 9.1: Skipping file - summarization failed.")
        else:
            print("Step 9.1: Skipping file - no text found in the specified range.")
    
    print("\n--- Step 10: All supported files processed ---")
    return all_summaries


# --- Main Execution ---
if __name__ == "__main__":
    try:
        CATEGORIES = load_and_edit_categories()

        while True:
            folder_path = input("Enter the folder path containing the files: ").strip()
            if os.path.isdir(folder_path):
                break
            print("\nError: The provided path is not a valid directory. Please try again.")

        scan_sub_input = input("Scan subdirectories? (y/n) [default: n]: ").strip().lower()
        scan_subdirectories = scan_sub_input in ['y', 'yes']

        # --- NEW: Configuration Settings Prompts ---
        print("\n--- Configuration Settings ---")
        while True:
            try:
                chunk_len_input = input("Enter chunk length (words per chunk, default: 30): ").strip()
                max_chunk_length_to_use = int(chunk_len_input) if chunk_len_input else 30
                if max_chunk_length_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        while True:
            try:
                max_sum_input = input("Enter max summary length (words, default: 10): ").strip()
                max_summary_length_to_use = int(max_sum_input) if max_sum_input else 10
                if max_summary_length_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        while True:
            try:
                min_sum_input = input("Enter min summary length (words, default: 4): ").strip()
                min_summary_length_to_use = int(min_sum_input) if min_sum_input else 4
                if min_summary_length_to_use > 0 and min_summary_length_to_use <= max_summary_length_to_use:
                    break
                else:
                    print(f"Please enter a number greater than 0 and no more than {max_summary_length_to_use}.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        # --- CORRECTED: Independent and clear prompts ---
        print(f"\n--- PDF Settings (in {max_chunk_length_to_use}-word chunks) ---")
        while True:
            try:
                padding_input = input("Enter number of initial chunks to skip (padding, default: 0): ").strip()
                pdf_padding_to_use = int(padding_input) if padding_input else 0
                if pdf_padding_to_use >= 0:
                    break
                else:
                    print("Please enter a non-negative number.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        while True:
            try:
                chunks_input = input("Enter number of chunks to process after padding: ").strip()
                pdf_chunks_to_use = int(chunks_input)
                if pdf_chunks_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        print(f"\n--- Settings for Other File Types (in {max_chunk_length_to_use}-word chunks) ---")
        while True:
            try:
                padding_input = input("Enter number of initial chunks to skip (padding, default: 0): ").strip()
                generic_padding_to_use = int(padding_input) if padding_input else 0
                if generic_padding_to_use >= 0:
                    break
                else:
                    print("Please enter a non-negative number.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        while True:
            try:
                chunks_input = input("Enter number of chunks to process after padding: ").strip()
                generic_chunks_to_use = int(chunks_input)
                if generic_chunks_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        # --- NEW: Confidence Threshold Prompt ---
        while True:
            try:
                threshold_input = input("\nEnter confidence threshold (0-100, default: 0): ").strip()
                confidence_threshold_to_use = int(threshold_input) if threshold_input else 0
                if 0 <= confidence_threshold_to_use <= 100:
                    break
                else:
                    print("Please enter a number between 0 and 100.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")


        file_management_settings = {}
        move_files_choice = input("\nAutomatically move all categorized files into subdirectories? (y/n) [default: n]: ").strip().lower()

        if move_files_choice in ['y', 'yes']:
            add_prefix_choice = input("Add a prefix to moved files? (y/n) [default: n]: ").strip().lower()
            
            if add_prefix_choice in ['y', 'yes']:
                global_prefix_choice = input("Apply a global prefix to all moved files? (y/n) [default: n]: ").strip().lower()
                
                if global_prefix_choice in ['y', 'yes']:
                    global_prefix = ""
                    while True:
                        global_prefix = input("Enter a prefix (max 12 characters) for all files: ").strip()
                        if len(global_prefix) <= 12:
                            break
                        else:
                            print(f"The prefix must be 12 characters or less. Please try again.")
                    
                    special_char = ""
                    while True:
                        special_char = input("Enter a special character to follow the prefix (1 to 3 characters): ").strip()
                        if 1 <= len(special_char) <= 3:
                            break
                        else:
                            print("Please enter 1 to 3 special characters.")
                    
                    for category in CATEGORIES:
                        if category != "Other":
                            file_management_settings[category] = {
                                'prefix': f"{global_prefix}{special_char}",
                                'destination': os.path.join(folder_path, category)
                            }
                else:
                    print("Using default prefixes for moved files.")
                    for category in CATEGORIES:
                        if category != "Other":
                            file_management_settings[category] = {
                                'prefix': category.replace(' ', '')[:12].upper(),
                                'destination': os.path.join(folder_path, category)
                            }
            else:
                print("Files will be moved without a prefix.")
                for category in CATEGORIES:
                    if category != "Other":
                        file_management_settings[category] = {
                            'prefix': '',
                            'destination': os.path.join(folder_path, category)
                        }
            
            # --- MODIFIED: Always create a rule for the "Other" category if moving is enabled ---
            print("Files categorized as 'Other' will be moved to an 'Other' folder.")
            file_management_settings["Other"] = {
                'prefix': '', # No prefix for the 'Other' category
                'destination': os.path.join(folder_path, "Other")
            }
        else:
            print("Categorized files will not be moved.")

        # --- NEW: Pre-processing Step Integration ---
        preprocess_choice = input("\nPre-process files by moving them based on filename? (y/n) [default: n]: ").strip().lower()
        
        # Initial scan for all files
        print("\nScanning for files... This may take a moment for large directories.")
        supported_extensions = [".pdf", ".docx", ".txt", ".xlsx", ".pptx"]
        initial_file_list = []
        if scan_subdirectories:
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if os.path.splitext(file)[1].lower() in supported_extensions:
                        initial_file_list.append(os.path.join(root, file))
        else:
            for file in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file)
                if os.path.isfile(file_path) and os.path.splitext(file)[1].lower() in supported_extensions:
                    initial_file_list.append(file_path)

        files_to_process = initial_file_list
        if preprocess_choice in ['y', 'yes']:
            if move_files_choice in ['y', 'yes']:
                 files_to_process = preprocess_by_filename(initial_file_list, CATEGORIES, file_management_settings)
            else:
                print("Cannot pre-process because file moving is disabled. Skipping.")


        print("\nStep 11: Folder path is valid. Starting the batch process...")
        
        all_summaries = process_files_in_folder(files_to_process, CATEGORIES, pdf_padding_to_use, pdf_chunks_to_use, generic_padding_to_use, generic_chunks_to_use, file_management_settings, confidence_threshold_to_use, max_chunk_length_to_use, max_summary_length_to_use, min_summary_length_to_use)

        if all_summaries:
            print("\n--- Saving all summaries to a single file ---")
            consolidated_summary_file = os.path.join(folder_path, "all_summaries.txt")
            with open(consolidated_summary_file, "w", encoding="utf-8") as f:
                for summary_data in all_summaries:
                    f.write("="*50 + "\n")
                    f.write(f"File: {summary_data['file_path']}\n")
                    f.write(f"Category: {summary_data['category']}\n\n")
                    f.write(f"Summary:\n{summary_data['summary']}\n")
            print(f"All summaries consolidated into: {consolidated_summary_file}")
            
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. Exiting gracefully.")
