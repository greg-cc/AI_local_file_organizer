import os
import fitz  # PyMuPDF for PDF
from docx import Document
import pandas as pd
from pptx import Presentation
from transformers import pipeline
import torch
from transformers.pipelines.text_classification import TextClassificationPipeline
# NEW CODE: Added imports for explicit model and tokenizer loading
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM
# NEW CODE: Added import for Hugging Face Hub login
from huggingface_hub import login
import shutil
import json

# --- Custom Pipeline for Zero-Shot Classification with Threshold ---
class ThresholdZeroShotClassificationPipeline(TextClassificationPipeline):
    """
    Custom pipeline for zero-shot classification that returns 'Other' if the top
    score is below a specified threshold.
    """
    def __init__(self, *args, threshold=0.1, **kwargs):
        super().__init__(*args, **kwargs)
        self.threshold = threshold

    def __call__(self, sequences, *args, **kwargs):
        # Run the standard zero-shot classification
        results = super().__call__(sequences, *args, **kwargs)

        # Process results to apply the threshold
        processed_results = []
        # The pipeline can return a single dict or a list of dicts
        if isinstance(results, dict):
            results = [results]

        for result in results:
            if result['scores'][0] >= self.threshold:
                processed_results.append(result['labels'][0])
            else:
                processed_results.append("Other")

        # Return a single string if only one sequence was passed, else a list of strings
        return processed_results[0] if len(processed_results) == 1 else processed_results

# --- Model Selection Function ---
def select_summarization_model():
    """
    Prompts the user to select a summarization model from a list of options.
    """
    print("\nPlease select a summarization model to use:")
    print("\n--- General Purpose ---")
    print("1: (242 MB) t5-small - Balanced speed and quality (Default)")
    print("2: (496 MB) sshleifer/distilbart-cnn-12-6 - Faster, distilled model for speed")
    print("3: (1630 MB) facebook/bart-large-cnn - Larger, higher quality model (Slower)")
    print("\n--- Specialized Models ---")
    print("4: (2280 MB) google/pegasus-xsum - Headline-style summaries")
    print("5: (892 MB) mrm8488/t5-base-finetuned-summarize-news - Specialized for news articles")
    print("6: (1630 MB) philschmid/bart-large-cnn-samsum - Specialized for conversations/dialogue")
    print("7: (242 MB) Falconsai/medical_summarization - High-performing for medical text")
    print("8: (990 MB) google/flan-t5-base - Excellent for complex text & instructions")
    print("\n--- Lightweight & Distilled ---")
    print("9: (242 MB) Falconsai/text_summarization - Lightweight & efficient")
    print("10: (460 MB) sshleifer/distilbart-cnn-6-6 - Smaller DistilBART")
    print("11: (2440 MB) facebook/mbart-large-cc25 - Compact multilingual")
    print("12: (268 MB) distilbert-base-uncased-finetuned-sst-2-english - Tiny, for extractive tasks")
    print("\n--- Multilingual & Resource-Efficient Models ---")
    print("13: (2330 MB) csebuetnlp/mT5_multilingual_XLSum - Summarizes in 44 languages")
    print("14: (496 MB) ainize/kobart-news - Specialized for Korean news")
    print("15: (125 MB) google/t5-efficient-mini - Resource-efficient for custom datasets")
    
    models = {
        "1": "t5-small",
        "2": "sshleifer/distilbart-cnn-12-6",
        "3": "facebook/bart-large-cnn",
        "4": "google/pegasus-xsum",
        "5": "mrm8488/t5-base-finetuned-summarize-news",
        "6": "philschmid/bart-large-cnn-samsum",
        "7": "Falconsai/medical_summarization",
        "8": "google/flan-t5-base",
        "9": "Falconsai/text_summarization",
        "10": "sshleifer/distilbart-cnn-6-6",
        "11": "facebook/mbart-large-cc25",
        "12": "distilbert-base-uncased-finetuned-sst-2-english",
        "13": "csebuetnlp/mT5_multilingual_XLSum",
        "14": "ainize/kobart-news",
        "15": "google/t5-efficient-mini"
    }
    
    while True:
        choice = input("Enter your choice (1-15): ").strip()
        if choice in models:
            return models[choice]
        elif choice == "": # Default option
            return models["1"]
        print("Invalid choice. Please enter a number from 1 to 15.")

# --- NEW FUNCTION: Tiny Model Selection ---
def select_tiny_model():
    """
    Prompts the user to select a tiny summarization model.
    """
    print("\nPlease select a tiny summarization model:")
    print("1: (242 MB) moussaKam/t5-small-finetuned-xsum - General purpose, reliable")
    print("2: (242 MB) henryu-lin/t5-small-samsum-deepspeed - Tuned for dialogue")
    print("3: (268 MB) distilbert-base-uncased-finetuned-sst-2-english - Tiny, for extractive tasks")
    print("4: (125 MB) google/t5-efficient-mini - Most resource-efficient")
    print("5: (242 MB) Vamsi/T5_Paraphrase_Paws - Specialized in paraphrasing")
    
    models = {
        "1": "moussaKam/t5-small-finetuned-xsum",
        "2": "henryu-lin/t5-small-samsum-deepspeed",
        "3": "distilbert-base-uncased-finetuned-sst-2-english",
        "4": "google/t5-efficient-mini",
        "5": "Vamsi/T5_Paraphrase_Paws"
    }

    while True:
        choice = input("Enter your choice (1-5): ").strip()
        if choice in models:
            return models[choice]
        elif choice == "": # Default option
            return models["1"]
        print("Invalid choice. Please enter a number from 1 to 5.")

# --- NEW FUNCTION: Get Summary Lengths ---
def get_summary_lengths():
    """
    Prompts the user to set the min and max summary lengths.
    """
    min_len, max_len = 220, 400  # Default values
    
    while True:
        try:
            min_input = input(f"Enter MIN summary length [default: {min_len}]: ").strip()
            if min_input == "":
                min_len_to_use = min_len
            else:
                min_len_to_use = int(min_input)
            
            if min_len_to_use <= 0:
                print("Please enter a number greater than 0.")
                continue
            break
        except ValueError:
            print("Invalid input. Please enter a valid number.")

    while True:
        try:
            max_input = input(f"Enter MAX summary length [default: {max_len}]: ").strip()
            if max_input == "":
                max_len_to_use = max_len
            else:
                max_len_to_use = int(max_input)

            if max_len_to_use >= min_len_to_use:
                break
            else:
                print("Max length must be greater than or equal to the min length.")
        except ValueError:
            print("Invalid input. Please enter a valid number.")
            
    print(f"Summary lengths set to MIN: {min_len_to_use}, MAX: {max_len_to_use}")
    return min_len_to_use, max_len_to_use

# --- NEW FUNCTION: Get and remember classifier threshold for each model ---
def get_classifier_threshold(model_name):
    """
    Asks the user for a classifier threshold, remembering the value for each model.
    """
    settings_file = "model_thresholds.json"
    thresholds = {}
    
    # Load existing thresholds from file
    try:
        if os.path.exists(settings_file):
            with open(settings_file, 'r', encoding='utf-8') as f:
                thresholds = json.load(f)
    except Exception as e:
        print(f"Warning: Could not load threshold settings file: {e}")

    # Get the saved threshold for the current model, or use 0.1 as a default
    default_threshold = thresholds.get(model_name, 0.1)

    while True:
        prompt = f"\nEnter classifier threshold for '{model_name}' (0.0-1.0) [default: {default_threshold}]: "
        user_input = input(prompt).strip()

        if user_input == "":
            chosen_threshold = default_threshold
            break

        try:
            chosen_threshold = float(user_input)
            if 0.0 <= chosen_threshold <= 1.0:
                break
            else:
                print("Error: Please enter a number between 0.0 and 1.0.")
        except ValueError:
            print("Error: Invalid input. Please enter a valid number.")

    # Save the chosen threshold for this model
    thresholds[model_name] = chosen_threshold
    try:
        with open(settings_file, 'w', encoding='utf-8') as f:
            json.dump(thresholds, f, indent=4)
        print(f"Threshold for this model set to {chosen_threshold} and saved.")
    except Exception as e:
        print(f"Warning: Could not save threshold settings: {e}")
        
    return chosen_threshold

# Check for GPU availability and set the device
if torch.cuda.is_available():
    device = "cuda"
    print("Device set to use GPU.")
else:
    device = "cpu"
    print("GPU not found. Device set to use CPU.")

# --- Configuration ---
MAX_CHUNK_LENGTH = 512
# Summary lengths are now set dynamically

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
        "Technology",
        "Finance"
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
            command = input("> ").strip()
            if command.lower() == 'done':
                break
            
            parts = command.split()
            if not parts:
                continue

            action = parts[0].lower()
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
                pass # The list will be printed below
            else:
                print("Invalid command. Please try again.")

            print("\n--- Current Categories ---")
            for i, cat in enumerate(categories):
                print(f"[{i+1}] {cat}")
            print("--------------------------")
        
    # Save the final list of categories
    with open(category_file, 'w', encoding='utf-8') as f:
        json.dump(categories, f, indent=4)
    print("\nStep 1.3: Categories saved for next session.")
    
    return categories

# --- File Reading Functions ---
def read_pdf(file_path, start_chunk, end_chunk):
    """Extracts text from a PDF file within a specified word-chunk range."""
    text = ""
    try:
        full_text = ""
        with fitz.open(file_path) as pdf:
            for page in pdf:
                full_text += page.get_text()
        
        words = full_text.split()
        start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
        end_index = end_chunk * MAX_CHUNK_LENGTH
        # Ensure we don't go out of bounds
        end_index = min(end_index, len(words))
        
        text = " ".join(words[start_index:end_index])
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

def read_docx(file_path, start_chunk, end_chunk):
    """Extracts text from a DOCX file within a specified chunk range."""
    text = ""
    try:
        doc = Document(file_path)
        full_text = "\n".join([para.text for para in doc.paragraphs])
        words = full_text.split()
        start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
        end_index = end_chunk * MAX_CHUNK_LENGTH
        text = " ".join(words[start_index:end_index])
    except Exception as e:
        print(f"Error reading DOCX: {e}")
    return text

def read_txt(file_path, start_chunk, end_chunk):
    """Reads text from a plain TXT file within a specified chunk range."""
    text = ""
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            words = file.read().split()
            start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
            end_index = end_chunk * MAX_CHUNK_LENGTH
            text = " ".join(words[start_index:end_index])
    except Exception as e:
        print(f"Error reading TXT: {e}")
    return text

def read_excel(file_path, start_chunk, end_chunk):
    """Extracts text from an XLSX (Excel) file within a specified chunk range."""
    text = ""
    try:
        df = pd.read_excel(file_path, sheet_name=None)
        full_text = ""
        for sheet_name in df:
            full_text += df[sheet_name].to_string(index=False) + "\n"
        words = full_text.split()
        start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
        end_index = end_chunk * MAX_CHUNK_LENGTH
        text = " ".join(words[start_index:end_index])
    except Exception as e:
        print(f"Error reading XLSX: {e}")
    return text

def read_pptx(file_path, start_chunk, end_chunk):
    """Extracts text from a PPTX (PowerPoint) file within a specified chunk range."""
    text = ""
    try:
        presentation = Presentation(file_path)
        full_text = ""
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    full_text += shape.text + "\n"
        words = full_text.split()
        start_index = (start_chunk - 1) * MAX_CHUNK_LENGTH
        end_index = end_chunk * MAX_CHUNK_LENGTH
        text = " ".join(words[start_index:end_index])
    except Exception as e:
        print(f"Error reading PPTX: {e}")
    return text

# --- Summarization Function ---
def summarize_text(text, summarizer_pipeline, min_len, max_len):
    """
    Generates a summary from the text using the provided summarizer pipeline.
    """
    if not text.strip():
        return ["No text to summarize."]

    words = text.split()
    chunks = [" ".join(words[i:i + MAX_CHUNK_LENGTH]) for i in range(0, len(words), MAX_CHUNK_LENGTH)]
    
    summaries = []
    
    print(f"Step 6.1.1: Summarizing {len(chunks)} chunks in a batch...")
    try:
        results = summarizer_pipeline(
            chunks, 
            max_length=max_len, 
            min_length=min_len, 
            truncation=True,
            batch_size=8
        )
        summaries = [result['summary_text'] for result in results]
        print(f"Step 6.1.2: Summarization for all chunks complete.")
    except Exception as e:
        print(f"Step 6.1.3: Error during batch summarization: {e}")
        for i, chunk in enumerate(chunks):
            try:
                summary = summarizer_pipeline(
                    chunk, max_length=max_len, min_length=min_len, truncation=True
                )
                summaries.append(summary[0]['summary_text'])
            except Exception as chunk_e:
                print(f"Error on chunk {i+1}: {chunk_e}")
                summaries.append(f"Summary failed for chunk {i+1}.")

    combined_summary = ". ".join(summaries)
    return [s.strip() for s in combined_summary.split(".") if s.strip()]
    
# --- Categorization Function ---
def categorize_summary(summary_text, categories, classifier_pipeline):
    """Categorizes a summary using a zero-shot classification model with a built-in threshold."""
    if classifier_pipeline is None:
        return "Other"

    print("Step 7.1.1: Sending summary to classifier model...")
    try:
        category = classifier_pipeline(summary_text, candidate_labels=categories)
        return category
    except Exception as e:
        print(f"Step 7.1.3: Error during AI categorization: {e}")
        return "Other"

# --- Main Processing Logic ---
def process_files_in_folder(folder_path, scan_subdirectories, categories, start_chunk, end_chunk, file_management_settings, summarizer, classifier, 

min_len, max_len):
    """
    Walks a folder, processes supported files, and generates summaries.
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

    for i, file_path in enumerate(file_list):
        print(f"\nStep 5: Processing file {i + 1}/{len(file_list)} - {os.path.basename(file_path)}")
        reader = supported_extensions[os.path.splitext(file_path)[1].lower()]

        print(f"Step 6: Extracting text from chunks {start_chunk} to {end_chunk}...")
        text = reader(file_path, start_chunk, end_chunk)

        if text.strip():
            print("Step 7: Text extracted successfully. Starting summarization...")
            bullet_points = summarize_text(text, summarizer, min_len, max_len)
            
            summary_text = " ".join(bullet_points)
            print("Step 7.1: Categorizing summary...")
            category = categorize_summary(summary_text, categories, classifier)

            print("Step 8: Final summary complete.")
            
            print(f"\nFile: {file_path}\nCategory: {category}\nSummary:\n")
            summary_text_for_file = ""
            for point in bullet_points:
                print(f"- {point}")
                summary_text_for_file += f"- {point}\n"
            print("\n" + "-"*50)

            all_summaries.append({
                'file_path': os.path.abspath(file_path),
                'category': category,
                'summary': summary_text_for_file
            })

            if category in file_management_settings and file_management_settings[category].get('destination'):
                settings = file_management_settings[category]
                prefix = settings.get('prefix', '')
                destination_folder = settings['destination']
                
                os.makedirs(destination_folder, exist_ok=True)

                original_file_name, original_file_extension = os.path.splitext(os.path.basename(file_path))
                
                new_file_name = f"{prefix}{original_file_name}{original_file_extension}"
                destination_path = os.path.join(destination_folder, new_file_name)
                
                try:
                    shutil.move(file_path, destination_path)
                    print(f"Moved and renamed: {file_path} -> {destination_path}")
                except Exception as move_e:
                    print(f"Error moving original file: {move_e}")
            
        else:
            print("Step 8.1: Skipping file - unable to extract meaningful text.")
    
    print("\n--- Step 9: All supported files processed ---")
    return all_summaries


# --- Main Execution ---
if __name__ == "__main__":
    try:
        # Question 1: Select the summarization model
        summarizer_model_name = select_summarization_model()

        # Question 2: Set the classifier threshold for the selected model
        classifier_threshold_to_use = get_classifier_threshold(summarizer_model_name)

        # Question 3: Ask if user wants to switch to a tiny model
        tiny_choice = input("\nDo you want to use a tiny model instead for better performance? (y/n) [default: n]: ").strip().lower()
        if tiny_choice in ['y', 'yes']:
            summarizer_model_name = select_tiny_model()

        # Question 4: Set summary lengths
        min_summary_length, max_summary_length = get_summary_lengths()

        # --- Load Models ---
        print(f"\nStep 1: Loading summarization model ({summarizer_model_name}) and classification model...")
        print("If this is the first time you are running the script, a large file download will begin now. Please wait for it to complete.")

        # NEW: Added loop and token prompt to handle loading errors
        summarizer = None
        while summarizer is None:
            try:
                print("Step 1a: Explicitly loading tokenizer...")
                tokenizer = AutoTokenizer.from_pretrained(summarizer_model_name)
                
                print("Step 1b: Explicitly loading model...")
                model = AutoModelForSeq2SeqLM.from_pretrained(summarizer_model_name).to(device)
                
                print("Step 1c: Creating summarization pipeline...")
                summarizer = pipeline(
                    "summarization",
                    model=model,
                    tokenizer=tokenizer,
                    device=0 if device == 'cuda' else -1,
                    framework="pt"
                )
                print("Step 1d: Summarization model loaded successfully.")
            except Exception as e:
                print(f"\nAn error occurred while loading the summarization model: {e}")
                print("This can happen with private or gated models, or if the model format is incompatible.")
                
                token_choice = input("Would you like to try logging in with a Hugging Face token? (y/n): ").strip().lower()
                if token_choice in ['y', 'yes']:
                    token = input("Please enter your Hugging Face Hub token: ").strip()
                    login(token=token)
                    print("Token accepted. Retrying model download...")
                else:
                    print("Model loading failed. Please check the model name or your connection.")
                    exit()

        try:
            classifier = pipeline(
                task="zero-shot-classification",
                model="MoritzLaurer/xtremedistil-l6-h256-mnli-fever-anli-ling-binary",
                device=device,
                pipeline_class=ThresholdZeroShotClassificationPipeline,
                # Use the user-defined threshold
                threshold=classifier_threshold_to_use
            )
            print("Step 1.1: Classification model loaded successfully with custom threshold.")
        except Exception as e:
            print(f"Error loading classification model: {e}")
            classifier = None
            print("Step 1.1: Falling back to 'Other' for all classifications.")

        CATEGORIES = load_and_edit_categories()

        while True:
            folder_path = input("Enter the folder path containing the files: ").strip()
            if os.path.isdir(folder_path):
                break
            print("\nError: The provided path is not a valid directory. Please try again.")

        scan_sub_input = input("Scan subdirectories? (y/n) [default: n]: ").strip().lower()
        scan_subdirectories = scan_sub_input in ['y', 'yes']

        while True:
            try:
                start_chunk_input = input("Enter the number of the first chunk (e.g., page number): ").strip()
                start_chunk_to_use = int(start_chunk_input)
                if start_chunk_to_use > 0:
                    break
                else:
                    print("Please enter a number greater than 0.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")
        
        while True:
            try:
                end_chunk_input = input("Enter the number of the last chunk: ").strip()
                end_chunk_to_use = int(end_chunk_input)
                if end_chunk_to_use >= start_chunk_to_use:
                    break
                else:
                    print("Please enter a number greater than or equal to the starting chunk number.")
            except ValueError:
                print("Invalid input. Please enter a valid number.")

        file_management_settings = {}
        move_and_rename_choice = input("\nDo you want to move and rename files to subfolders based on their category? (y/n) [default: n]: ").strip

().lower()
        
        if move_and_rename_choice in ['y', 'yes']:
            print("\n--- Define Global File Management Rules ---")
            prefix = input("Enter a global prefix for all renamed files (e.g., 'PROJ_') [leave blank for none]: ").strip()
            
            for category in CATEGORIES:
                if category != "Other":
                    dest_folder_name = category.replace(' ', '_')
                    file_management_settings[category] = {
                        'prefix': prefix,
                        'destination': os.path.join(folder_path, dest_folder_name)
                    }
            print("File management rules have been set up automatically for all categories.")


        print("\nStep 10: Starting the batch process...")
        
        all_summaries = process_files_in_folder(
            folder_path, 
            scan_subdirectories, 
            CATEGORIES, 
            start_chunk_to_use, 
            end_chunk_to_use, 
            file_management_settings,
            summarizer,
            classifier,
            min_summary_length,
            max_summary_length
        )

        if all_summaries:
            consolidated_summary_file = os.path.join(folder_path, "all_summaries.txt")
            print(f"\n--- Saving all summaries to {consolidated_summary_file} ---")
            with open(consolidated_summary_file, "w", encoding="utf-8") as f:
                for summary_data in all_summaries:
                    f.write("="*50 + "\n")
                    f.write(f"File: {summary_data['file_path']}\n")
                    f.write(f"Category: {summary_data['category']}\n\n")
                    f.write(f"Summary:\n{summary_data['summary']}\n")
            print("All summaries consolidated.")
            
    except KeyboardInterrupt:
        print("\nProcess interrupted by user. Exiting gracefully.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
