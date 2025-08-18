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
MAX_CHUNK_LENGTH = 30
MAX_SUMMARY_LENGTH = 10
MIN_SUMMARY_LENGTH = 4

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
        print("\nEditing categories. Enter 'add <new_category>', 'remove <number>', 'edit <number> <new_category>', 'list', or 'done' to finish.")
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
            elif action == 'remove' and len(parts) == 2 and parts[1].isdigit():
                idx = int(parts[1]) - 1
                if 0 <= idx < len(categories):
                    removed_cat = categories.pop(idx)
                    print(f"Removed: {removed_cat}")
                    print("\n--- Current Categories ---")
                    for i, cat in enumerate(categories):
                        print(f"[{i+1}] {cat}")
                    print("--------------------------")
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
                else:
                    print("Invalid category number.")
            elif action == 'list':
                print("\n--- Current Categories ---")
                for i, cat in enumerate(categories):
                    print(f"[{i+1}] {cat}")
                print("--------------------------")
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
def read_pdf_pages(file_path, start_page, end_page):
    """Extracts text from a specific page range in a PDF."""
    text = ""
    try:
        with fitz.open(file_path) as doc:
            # Ensure start_page is at least 1 and end_page is not beyond the doc length
            start_page = max(1, start_page)
            end_page = min(doc.page_count, end_page)
            
            # Convert to 0-based index for PyMuPDF
            for i in range(start_page - 1, end_page):
                text += doc[i].get_text()
    except Exception as e:
        print(f"Error reading PDF pages: {e}")
    return text

def read_generic_chunks(file_path, start_chunk, end_chunk):
    """Reads a range of word-chunks from non-PDF files."""
    full_text = ""
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == ".docx":
            doc = Document(file_path)
            full_text = "\n".join([para.text for para in doc.paragraphs])
        elif ext == ".txt":
            with open(file_path, "r", encoding="utf-8") as file:
                full_text = file.read()
        elif ext == ".xlsx":
            all_words = []
            sheets = pd.ExcelFile(file_path).sheet_names
            for sheet in sheets:
                df = pd.read_excel(file_path, sheet_name=sheet)
                all_words.extend(df.to_string(index=False).split())
            full_text = " ".join(all_words)
        elif ext == ".pptx":
            presentation = Presentation(file_path)
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        full_text += shape.text + "\n"
    except Exception as e:
        print(f"Error reading {ext} file: {e}")
        return ""

    words = full_text.split()
    start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
    end_index = end_chunk * MAX_CHUNK_LENGTH
    return " ".join(words[start_index:end_index])

# --- Summarization Function ---
def summarize_text(text):
    """
    Generates a summary from the provided text by processing it in chunks.
    """
    if not text.strip():
        return ["No text to summarize."]

    words = text.split()
    chunks = [" ".join(words[i:i + MAX_CHUNK_LENGTH]) for i in range(0, len(words), MAX_CHUNK_LENGTH)]
    
    summaries = []
    
    for i, chunk in enumerate(chunks):
        print(f"Step 7.1.1: Summarizing chunk {i + 1} of {len(chunks)}...")
        try:
            summary = summarizer(
                chunk, 
                max_length=MAX_SUMMARY_LENGTH, 
                min_length=MIN_SUMMARY_LENGTH, 
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
    Returns a list of all labels, sorted by score.
    """
    if classifier is None:
        return ["Other"]

    print("Step 8.1: Sending summary to classifier model...")
    try:
        results = classifier(summary_text, candidate_labels=categories)
        print("Step 8.2: Categorization complete.")
        return results['labels']
    except Exception as e:
        print(f"Step 8.3: Error during AI categorization: {e}")
        return ["Other"]

# --- Main Processing Logic ---
def process_files_in_folder(folder_path, scan_subdirectories, categories, pdf_padding, pdf_chunks, generic_padding, generic_chunks, file_management_settings):
    """
    Walks a folder, processes supported files, and generates summaries.
    Includes detailed progress checks.
    """
    supported_extensions = [".pdf", ".docx", ".txt", ".xlsx", ".pptx"]
    
    print("\nStep 2: Starting folder scan...")
    file_list = []
    if scan_subdirectories:
        for root, _, files in os.walk(folder_path):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in supported_extensions:
                    file_list.append(os.path.join(root, file))
    else:
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file)[1].lower()
                if ext in supported_extensions:
                    file_list.append(file_path)

    if not file_list:
        print("Step 3: No supported files found. Exiting.")
        return []

    print(f"Step 4: Found {len(file_list)} supported files. Processing them now...")

    all_summaries = []

    for i, file_path in enumerate(file_list):
        print(f"\nStep 5: Processing file {i + 1}/{len(file_list)} - {os.path.basename(file_path)}")
        
        is_pdf = os.path.splitext(file_path)[1].lower() == '.pdf'
        text = ""
        
        if is_pdf:
            start_page = pdf_padding + 1
            end_page = pdf_padding + pdf_chunks
            print(f"\nStep 6: Extracting from page(s) {start_page} to {end_page}...")
            text = read_pdf_pages(file_path, start_page, end_page)
        else:
            start_chunk = generic_padding + 1
            end_chunk = generic_padding + generic_chunks
            print(f"\nStep 6: Extracting from chunk(s) {start_chunk} to {end_chunk}...")
            text = read_generic_chunks(file_path, start_chunk, end_chunk)

        if text.strip():
            print("Step 7: Text extracted successfully. Starting summarization...")
            bullet_points = summarize_text(text)
            
            if bullet_points and bullet_points != ["No text to summarize."]:
                file_name_for_analysis = os.path.basename(file_path)
                file_name_no_ext = os.path.splitext(file_name_for_analysis)[0]
                cleaned_file_name = file_name_no_ext.replace('_', ' ').replace('-', ' ')
                
                all_chunks_for_classifier = [cleaned_file_name] + bullet_points
                text_for_classifier = ". ".join(all_chunks_for_classifier)
                
                print("Step 8: Categorizing summary offline...")
                ranked_categories = categorize_summary(text_for_classifier, categories)
                top_category = ranked_categories[0]
                other_categories = ranked_categories[1:]

                print("Step 9: Final summary complete.")
                
                print(f"\nFile: {file_path}")
                print("Summary:\n")
                
                print(f"Category: {Fore.GREEN}{Style.BRIGHT}{top_category}")
                
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

        # --- MODIFIED: Clearer, separate prompts ---
        print("\n--- PDF Settings ---")
        while True:
            try:
                padding_input = input(f"Enter number of initial pages to skip (padding, default: 0): ").strip()
                pdf_padding_to_use = int(padding_input) if padding_input else 0
                if pdf_padding_to_use >= 0:
                    break
                else:
                    print("Please enter a non-negative number.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        while True:
            try:
                chunks_input = input("Enter number of pages to process after padding: ").strip()
                pdf_chunks_to_use = int(chunks_input)
                if pdf_chunks_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        print(f"\n--- Settings for Other File Types (1 chunk = {MAX_CHUNK_LENGTH} words) ---")
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
        else:
            print("Categorized files will not be moved.")

        print("\nStep 11: Folder path is valid. Starting the batch process...")
        
        all_summaries = process_files_in_folder(folder_path, scan_subdirectories, CATEGORIES, pdf_padding_to_use, pdf_chunks_to_use, generic_padding_to_use, generic_chunks_to_use, file_management_settings)

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
