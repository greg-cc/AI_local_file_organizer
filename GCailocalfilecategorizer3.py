import os
import fitz  # PyMuPDF for PDF
from docx import Document
import pandas as pd
from pptx import Presentation
from transformers import pipeline
import torch
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import shutil

# Check for GPU availability and set the device
if torch.cuda.is_available():
    device = "cuda"
    print("Device set to use GPU.")
else:
    device = "cpu"
    print("GPU not found. Device set to use CPU.")

# --- Configuration ---
NUM_CHUNKS_TO_SUMMARIZE = 3
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

# --- Load Models ---
print("\nStep 1: Loading summarization and classification models...")
print("If this is the first time you are running the script, a large file download will begin now. Please wait for it to complete.")

# Load the SMALLER summarization pipeline for speed
summarizer = pipeline(
    "summarization",
    model="sshleifer/distilbart-cnn-6-6", 
    max_length=1024, 
    device=device,
    framework="pt"  # Explicitly tell the pipeline to use PyTorch
)

# Load a dedicated model for text classification
try:
    classifier = pipeline(
        "zero-shot-classification",
        model="MoritzLaurer/mDeBERTa-v3-base-mnli-xnli",
        device=device
    )
    print("Step 1.1: Classification model loaded successfully.")
except Exception as e:
    print(f"Error loading classification model: {e}")
    classifier = None
    print("Step 1.1: Falling back to 'Other' for all classifications.")

# --- File Reading Functions ---
def read_pdf(file_path):
    """Extracts text from a PDF file using PyMuPDF, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        with fitz.open(file_path) as pdf:
            for page in pdf:
                page_text = page.get_text()
                words = page_text.split()
                if word_count + len(words) >= NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH:
                    remaining_words = (NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH) - word_count
                    text += " ".join(words[:remaining_words])
                    break
                else:
                    text += page_text
                    word_count += len(words)
    except Exception as e:
        print(f"Error reading PDF: {e}")
        text = ""
    return text

def read_docx(file_path):
    """Extracts text from a DOCX file, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        doc = Document(file_path)
        for para in doc.paragraphs:
            words = para.text.split()
            if word_count + len(words) >= NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH:
                remaining_words = (NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH) - word_count
                text += " ".join(words[:remaining_words])
                break
            else:
                text += para.text + "\n"
                word_count += len(words)
    except Exception as e:
        print(f"Error reading DOCX: {e}")
        text = ""
    return text

def read_txt(file_path):
    """Reads text from a plain TXT file, stopping after a specified number of chunks."""
    text = ""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            words = file.read().split()
            text = " ".join(words[:NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH])
    except Exception as e:
        print(f"Error reading TXT: {e}")
        text = ""
    return text

def read_excel(file_path):
    """Extracts text from an XLSX (Excel) file, stopping after a specified number of chunks."""
    text = ""
    try:
        sheets = pd.ExcelFile(file_path).sheet_names
        for sheet in sheets:
            df = pd.read_excel(file_path, sheet_name=sheet)
            words = df.to_string(index=False).split()
            text = " ".join(words[:NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH])
            if len(words) >= NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH:
                break
    except Exception as e:
        print(f"Error reading XLSX: {e}")
        text = ""
    return text

def read_pptx(file_path):
    """Extracts text from a PPTX (PowerPoint) file, stopping after a specified number of chunks."""
    text = ""
    word_count = 0
    try:
        presentation = Presentation(file_path)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    words = shape.text.split()
                    if word_count + len(words) >= NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH:
                        remaining_words = (NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH) - word_count
                        text += " ".join(words[:remaining_words])
                        break
                    else:
                        text += shape.text + "\n"
                        word_count += len(words)
            if word_count >= NUM_CHUNKS_TO_SUMMARIZE * MAX_CHUNK_LENGTH:
                break
    except Exception as e:
        print(f"Error reading PPTX: {e}")
        text = ""
    return text

# --- Summarization Function ---
def summarize_text(text):
    """
    Generates a summary from the first three chunks of text.
    """
    if not text.strip():
        return ["No text to summarize."]

    words = text.split()
    chunks = [" ".join(words[i:i + MAX_CHUNK_LENGTH]) for i in range(0, len(words), MAX_CHUNK_LENGTH)]
    
    summaries = []
    
    # Process only the first three chunks
    for i in range(min(NUM_CHUNKS_TO_SUMMARIZE, len(chunks))):
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
def process_files_in_folder(folder_path):
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
    for root, _, files in os.walk(folder_path):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in supported_extensions:
                file_list.append(os.path.join(root, file))

    if not file_list:
        print("Step 3: No supported files found. Exiting.")
        return

    print(f"Step 4: Found {len(file_list)} supported files. Processing them now...")

    # Main loop to process files one-by-one and display summaries immediately
    for i, file_path in enumerate(file_list):
        print(f"\nStep 5: Processing file {i + 1}/{len(file_list)} - {os.path.basename(file_path)}")
        reader = supported_extensions[os.path.splitext(file_path)[1].lower()]

        print(f"Step 6: Extracting first {NUM_CHUNKS_TO_SUMMARIZE} chunks of text...")
        text = reader(file_path)

        if text.strip():
            print("Step 7: First chunks extracted successfully. Starting summarization...")
            bullet_points = summarize_text(text)
            
            # Categorize the summary
            summary_text = " ".join(bullet_points)
            print("Step 7.1: Categorizing summary offline...")
            category = categorize_summary(summary_text, CATEGORIES)

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

            # Save to text file
            with open(summary_file_path, "w", encoding="utf-8") as summary_file:
                summary_file.write(f"Summary for: {file_path}\n\n")
                summary_file.write(f"Category: {category}\n\n")
                summary_file.write(summary_text_for_file)
                print(f"Summary saved to: {summary_file_path}")
        else:
            print("Step 8.1: Skipping file - unable to extract meaningful text.")
    
    print("\n--- Step 9: All supported files processed ---")


# --- Main Execution ---
if __name__ == "__main__":
    try:
        folder_path = input("Enter the folder path containing the files: ").strip()

        # Check if the folder path is valid
        if os.path.isdir(folder_path):
            print("\nStep 10: Folder path is valid. Starting the batch process...")
            process_files_in_folder(folder_path)

            # Prompt for file management after processing is complete
            print("\n--- File Management ---")
            category_to_move = input("Enter the category you want to move to its own folder (e.g., 'Health'): ").strip()
            
            # Check if the user wants to move a valid category
            if category_to_move in CATEGORIES and category_to_move != "Other":
                
                # Ask for a prefix for renaming
                prefix = input(f"Enter a prefix (max 8 characters) for renaming files in the '{category_to_move}' category: ").strip()
                while len(prefix) > 8:
                    print("Prefix is too long. Please enter a prefix with a maximum of 8 characters.")
                    prefix = input("Enter a new prefix: ").strip()

                # Create the destination folder
                destination_folder = os.path.join(folder_path, category_to_move)
                os.makedirs(destination_folder, exist_ok=True)
                print(f"\nMoving and renaming files for category '{category_to_move}' to '{destination_folder}'...")

                # Iterate through files and move/rename them
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        if file.lower().endswith('_summary.txt'):
                            file_path = os.path.join(root, file)
                            try:
                                with open(file_path, 'r', encoding='utf-8') as f:
                                    content = f.read()
                                    if f"Category: {category_to_move}" in content:
                                        # Get the original file path and name
                                        original_base_name = os.path.splitext(file)[0][:-8]
                                        original_pdf_path = os.path.join(root, original_base_name + '.pdf')
                                        
                                        # Create the new file names
                                        new_pdf_name = f"{prefix}_{os.path.basename(original_pdf_path)}"
                                        new_summary_name = f"{prefix}_{os.path.basename(file)}"
                                        
                                        destination_pdf_path = os.path.join(destination_folder, new_pdf_name)
                                        summary_destination_path = os.path.join(destination_folder, new_summary_name)
                                        
                                        # Move the original PDF
                                        if os.path.exists(original_pdf_path):
                                            shutil.move(original_pdf_path, destination_pdf_path)
                                            print(f"Moved and renamed: {original_pdf_path} -> {destination_pdf_path}")
                                            
                                        # Also move the summary file
                                        shutil.move(file_path, summary_destination_path)
                                        print(f"Moved and renamed: {file_path} -> {summary_destination_path}")
                                        
                            except Exception as e:
                                print(f"Error processing {file}: {e}")

                print("\nFile management complete.")
            else:
                print("No valid category selected for file management.")
        else:
            print("\nError: The provided path is not a valid directory.")
            
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. Exiting gracefully.")
