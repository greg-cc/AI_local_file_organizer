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

# Check for GPU availability and set the device
if torch.cuda.is_available():
    device = "cuda"
    print("Device set to use GPU.")
else:
    device = "cpu"
    print("GPU not found. Device set to use CPU.")

# --- Configuration ---
MAX_CHUNK_LENGTH = 512
MAX_SUMMARY_LENGTH = 300
MIN_SUMMARY_LENGTH = 80

# Define the categories for categorization
CATEGORIES = [
    "Health",
    "History",
    "Zionism, Jews, and Israel",
    "Surf, Weather, Metrological",
    "Hyperspectral imaging including anything from JPL or mentioning spectral terms",
    "Other"
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
        "Zionism, Jews, and Israel",
        "Surf, Weather, Metrological",
        "Hyperspectral imaging including anything from JPL or mentioning spectral terms",
        "Other"
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
def read_pdf(file_path, num_chunks):
    """Extracts text from a PDF file using PyMuPDF, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        with fitz.open(file_path) as pdf:
            for page in pdf:
                page_text = page.get_text()
                words = page_text.split()
                if word_count + len(words) >= num_chunks * MAX_CHUNK_LENGTH:
                    remaining_words = (num_chunks * MAX_CHUNK_LENGTH) - word_count
                    text += " ".join(words[:remaining_words])
                    break
                else:
                    text += page_text
                    word_count += len(words)
    except Exception as e:
        print(f"Error reading PDF: {e}")
        text = ""
    return text

def read_docx(file_path, num_chunks):
    """Extracts text from a DOCX file, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            words = para.text.split()
            if word_count + len(words) >= num_chunks * MAX_CHUNK_LENGTH:
                remaining_words = (num_chunks * MAX_CHUNK_LENGTH) - word_count
                text += " ".join(words[:remaining_words])
                break
            else:
                text += para.text + "\n"
                word_count += len(words)
    except Exception as e:
        print(f"Error reading DOCX: {e}")
        text = ""
    return text

def read_txt(file_path, num_chunks):
    """Reads text from a plain TXT file, stopping after a specified number of chunks."""
    text = ""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            words = file.read().split()
            text = " ".join(words[:num_chunks * MAX_CHUNK_LENGTH])
    except Exception as e:
        print(f"Error reading TXT: {e}")
        text = ""
    return text

def read_excel(file_path, num_chunks):
    """Extracts text from an XLSX (Excel) file, stopping after a specified number of chunks."""
    text = ""
    try:
        sheets = pd.ExcelFile(file_path).sheet_names
        for sheet in sheets:
            df = pd.read_excel(file_path, sheet_name=sheet)
            words = df.to_string(index=False).split()
            text = " ".join(words[:num_chunks * MAX_CHUNK_LENGTH])
            if len(words) >= num_chunks * MAX_CHUNK_LENGTH:
                break
    except Exception as e:
        print(f"Error reading XLSX: {e}")
        text = ""
    return text

def read_pptx(file_path, num_chunks):
    """Extracts text from a PPTX (PowerPoint) file, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        presentation = Presentation(file_path)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    words = shape.text.split()
                    if word_count + len(words) >= num_chunks * MAX_CHUNK_LENGTH:
                        remaining_words = (num_chunks * MAX_CHUNK_LENGTH) - word_count
                        text += " ".join(words[:remaining_words])
                        break
                    else:
                        text += shape.text + "\n"
                        word_count += len(words)
            if word_count >= num_chunks * MAX_CHUNK_LENGTH:
                break
    except Exception as e:
        print(f"Error reading PPTX: {e}")
        text = ""
    return text

# --- Summarization Function ---
def summarize_text(text, num_chunks):
    """
    Generates a summary from a specified number of chunks of text.
    """
    if not text.strip():
        return ["No text to summarize."]

    words = text.split()
    chunks = [" ".join(words[i:i + MAX_CHUNK_LENGTH]) for i in range(0, len(words), MAX_CHUNK_LENGTH)]
    
    summaries = []
    
    # Process only the specified number of chunks
    for i in range(min(num_chunks, len(chunks))):
        print(f"Step 6.1.1: Summarizing chunk {i + 1} of {len(chunks)}...")
        try:
            summary = summarizer(
                chunks[i], 
                max_length=MAX_SUMMARY_LENGTH, 
                min_length=MIN_SUMMARY_LENGTH, 
                truncation=True
            )
            summaries.append(summary[0]['summary_text'])
            print(f"Step 6.1.2: Summarization for chunk {i + 1} complete.")
        except Exception as e:
            print(f"Step 6.1.3: Error during summarization of chunk {i + 1}: {e}")
            summaries.append("Summary failed for this chunk.")
    
    # Combine the chunk summaries and return as a single list of bullet points
    combined_summary = ". ".join(summaries)
    return combined_summary.split(". ")
    
# --- Categorization Function ---
def categorize_summary(summary_text, categories):
    """Categorizes a summary using a zero-shot classification model."""
    if classifier is None:
        return "Other"

    print("Step 7.1.1: Sending summary to classifier model...")
    try:
        results = classifier(summary_text, candidate_labels=categories)
        print("Step 7.1.2: Categorization complete.")
        return results['labels'][0]
    except Exception as e:
        print(f"Step 7.1.3: Error during AI categorization: {e}")
        return "Other"

# --- Main Processing Logic ---
def process_files_in_folder(folder_path, scan_subdirectories, categories, num_chunks, file_management_settings):
    """
    Walks a folder, processes supported files, and generates summaries.
    Includes detailed progress checks.
    """
    supported_extensions = {
        ".pdf": read_pdf, 
        ".docx": read_docx, 
        ".txt": read_txt, 
        ".xlsx": read_excel, 
        ".pptx": read_pptx
    }
    
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

    # Main loop to process files one-by-one and display summaries immediately
    for i, file_path in enumerate(file_list):
        print(f"\nStep 5: Processing file {i + 1}/{len(file_list)} - {os.path.basename(file_path)}")
        reader = supported_extensions[os.path.splitext(file_path)[1].lower()]

        print(f"Step 6: Extracting first {num_chunks} chunks of text...")
        text = reader(file_path, num_chunks)

        if text.strip():
            print("Step 7: First chunks extracted successfully. Starting summarization...")
            bullet_points = summarize_text(text, num_chunks)
            
            # Categorize the summary
            summary_text = " ".join(bullet_points)
            print("Step 7.1: Categorizing summary offline...")
            category = categorize_summary(summary_text, categories)

            print("Step 8: Final summary complete.")
            
            # --- Display and Save Summary Immediately ---
            summary_file_path = os.path.splitext(file_path)[0] + "_summary.txt"
            
            # Print to screen
            print(f"\nFile: {file_path}\nCategory: {category}\nSummary:\n")
            summary_text_for_file = ""
            for point in bullet_points:
                print(f"- {point.strip()}")
                summary_text_for_file += f"- {point.strip()}\n"
            print("\n" + "-"*50)

            # Store the summary information for the final consolidated file
            all_summaries.append({
                'file_path': os.path.abspath(file_path),
                'category': category,
                'summary': summary_text_for_file
            })

            # Perform file management immediately after saving
            if category in file_management_settings and category != "Other":
                settings = file_management_settings[category]
                prefix = settings['prefix']
                destination_folder = settings['destination']
                
                os.makedirs(destination_folder, exist_ok=True)

                # Get the original file path and name
                original_file_extension = os.path.splitext(file_path)[1]
                original_file_name_no_ext = os.path.basename(os.path.splitext(file_path)[0])
                
                # Check if the summary file exists before trying to move it
                if os.path.exists(summary_file_path):
                    
                    new_pdf_name = f"{prefix}_{original_file_name_no_ext}{original_file_extension}"
                    new_summary_name = f"{prefix}_{original_file_name_no_ext}_summary.txt"

                    destination_pdf_path = os.path.join(destination_folder, new_pdf_name)
                    summary_destination_path = os.path.join(destination_folder, new_summary_name)
                    
                    # Move the original PDF
                    if os.path.exists(file_path):
                        try:
                            shutil.move(file_path, destination_pdf_path)
                            print(f"Moved and renamed: {file_path} -> {destination_pdf_path}")
                        except Exception as move_e:
                            print(f"Error moving original file: {move_e}")
                    
                    # Also move the summary file
                    try:
                        shutil.move(summary_file_path, summary_destination_path)
                        print(f"Moved and renamed: {summary_file_path} -> {summary_destination_path}")
                    except Exception as move_e:
                        print(f"Error moving summary file: {move_e}")
        else:
            print("Step 8.1: Skipping file - unable to extract meaningful text.")
    
    print("\n--- Step 9: All supported files processed ---")
    return all_summaries


# --- Main Execution ---
if __name__ == "__main__":
    try:
        # Load or edit categories at the start
        CATEGORIES = load_and_edit_categories()

        while True:
            folder_path = input("Enter the folder path containing the files: ").strip()
            if os.path.isdir(folder_path):
                break
            print("\nError: The provided path is not a valid directory. Please try again.")

        scan_sub_input = input("Scan subdirectories? (y/n) [default: n]: ").strip().lower()
        scan_subdirectories = scan_sub_input in ['y', 'yes']

        while True:
            num_chunks_input = input("Enter the number of chunks to summarize per file: ").strip()
            try:
                num_chunks_to_use = int(num_chunks_input)
                if num_chunks_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        # New logic to set up file management rules before processing
        file_management_settings = {}
        print("\n--- Define File Management Rules ---")
        for category in CATEGORIES:
            if category != "Other":
                choice = input(f"Do you want to manage files for '{category}'? (y/n) [default: n]: ").strip().lower()
                if choice in ['y', 'yes']:
                    prefix = ""
                    while True:
                        prefix = input(f"Enter an 8-character prefix for '{category}': ").strip()
                        if len(prefix) <= 8:
                            break
                        else:
                            print("Prefix must be 8 characters or less. Please try again.")
                    
                    dest_folder_name = input(f"Enter the destination folder name for '{category}' (leave blank to leave inplace and not move it'): ").strip()
                    if dest_folder_name:
                        file_management_settings[category] = {
                            'prefix': prefix,
                            'destination': os.path.join(folder_path, dest_folder_name)
                        }
                    else:
                        print(f"Files for '{category}' will not be moved.")
        print("\nStep 10: Folder path is valid. Starting the batch process...")
        
        # Call the main processing function and collect all summaries
        all_summaries = process_files_in_folder(folder_path, scan_subdirectories, CATEGORIES, num_chunks_to_use, file_management_settings)

        # Write all summaries to a single file at the end
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
