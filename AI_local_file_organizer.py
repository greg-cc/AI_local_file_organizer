import os
import shutil
import json
from pathlib import Path
import streamlit as st
import fitz  # PyMuPDF
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import torch

try:
    torch.classes.__path__ = []
except Exception:
    pass

from sentence_transformers import SentenceTransformer, util

# --- PAGE CONFIG ---
st.set_page_config(page_title="Zero-Shot File Sorter & Renamer", page_icon="📁", layout="wide")

CONFIG_FILE = Path("saved_categories.json")

def load_saved_categories():
    """Loads saved category definitions and individual thresholds from disk to persist user work."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                migrated = {}
                for k, v in data.items():
                    if isinstance(v, str):
                        migrated[k] = {"description": v, "threshold": 0.3}
                    elif isinstance(v, dict):
                        migrated[k] = {
                            "description": v.get("description", ""),
                            "threshold": v.get("threshold", 0.3)
                        }
                return migrated
        except Exception:
            pass
    
    # Default fallback with expanded descriptions and individual thresholds
    defaults = {
        "Health": "Health, wellness, medical research, clinical diagnosis, surgical procedures, operative reports, surgical excision, neurovascular bundles, tendons, ligaments, epiphysis, chemotherapy, postoperative care, clinical trials, dietary supplements, whole-food multivitamins, herbal extracts, traditional Chinese medicine formulations, raw pet food nutrition, green tripe, biological body systems, patient health records, and medical diagnostics.",
        "History": "Historical chronicles, archival source texts, primary source evidence, political and historical grievances, chronological timelines, past geopolitical events, and structural historical analysis.",
        "Weather": "Weather forecasts, meteorological conditions, real-time NOAA satellite wind maps, GOES-West imagery, CDIP buoy observations, coastal wave patterns, and atmospheric pressure systems.",
        "Finance": "Financial records, stock prices, market valuations, asset pricing, revenue statements, monetary transactions, commercial ledgers, and economic cost estimations.",
        "Legal": "Legal contracts, binding paperwork, terms of service, regulatory compliance, statutory frameworks, Docassemble interview files, AcroForm field mappings, and formal legal notices.",
        "Paralegal": "Legal document preparation, serving subpoenas for medical records, attorney-client communications, legal discovery paperwork, medical-legal record requests, and deposition filings.",
        "python_coding": "Python programming, software development, script automation, Streamlit local web applications, sentence-transformers, PyTorch, PyMuPDF, git repositories, code repositories, and technical debugging logs.",
        "aviris": "AVIRIS and ASTER hyperspectral datasets, remote sensing, ENVI 5.3 spectral analysis, SAGA GIS geological mapping, short-wave infrared sensors, and mineral prospecting data.",
        "Handyman": "Property maintenance, building repair, hardware specifications, power tools, electric demolition breakers, surface scrapers, fence construction, and residential renovation projects.",
        "Real Estate": "Real estate property listings, housing markets, property management, local real estate professional services, and structural property assessments.",
        "transcript": "Audio transcripts, spoken dialogue records, faster-whisper speech-to-text outputs, meeting dictations, and literal audio-to-text text blocks.",
        "religion": "Religious texts, theological doctrines, spiritual belief systems, historical religious philosophy, and comparative faith studies.",
        "politics": "Political analysis, geopolitical strategies, institutional press releases, government policies, public affairs, and international relations commentary.",
        "israel": "Geopolitical affairs, regional Middle Eastern history, foreign policy developments, and associated security or diplomatic documentation concerning Israel."
    }
    return {k: {"description": desc, "threshold": 0.3} for k, desc in defaults.items()}

def save_categories_to_disk(categories_dict):
    """Saves category definitions and thresholds to disk so work is never lost."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(categories_dict, f, indent=4)
    except Exception as e:
        st.error(f"Error saving categories: {e}")

# --- LOAD LOCAL EMBEDDING MODEL ---
@st.cache_resource
def load_embedding_model():
    return SentenceTransformer("all-MiniLM-L6-v2")

def extract_and_chunk_file_text(file_path, max_pages=50, max_chunk_tokens=140, 
                                pdf_chunks=5, other_chunks=5, txt_chunks=5,
                                pdf_skip=0, json_skip=0, txt_skip=0):
    """Extracts text from PDF, MD, TXT, JSON, or HTML files with expanded page limits and robust parsing."""
    text = ""
    ext = file_path.suffix.lower()
    try:
        if ext == ".pdf":
            with fitz.open(file_path) as doc:
                pages_to_read = min(len(doc), max_pages)
                for page_num in range(pages_to_read):
                    try:
                        page_text = doc[page_num].get_text("text")
                        if page_text:
                            text += page_text + "\n"
                    except Exception:
                        continue 
        elif ext == ".json":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                try:
                    data = json.load(f)
                    text = json.dumps(data, indent=2)
                except Exception:
                    f.seek(0)
                    text = f.read()
        elif ext in [".html", ".htm"]:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                soup = BeautifulSoup(f.read(), "html.parser")
                for element in soup(["script", "style", "noscript", "header", "footer", "nav"]):
                    element.decompose()
                text = soup.get_text(separator=" ")
        elif ext in [".md", ".txt"]:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
    except Exception as e:
        st.error(f"Error processing {file_path.name}: {e}")
    
    cleaned_text = text.strip()
    words = cleaned_text.split()
    if not words:
        return ""
    
    chunks = []
    for i in range(0, len(words), max_chunk_tokens):
        chunk_words = words[i:i + max_chunk_tokens]
        chunks.append(" ".join(chunk_words))
    
    total_available_chunks = len(chunks)
    
    if ext == ".pdf":
        target_num_chunks = pdf_chunks
        skip_count = pdf_skip
    elif ext == ".json":
        target_num_chunks = other_chunks
        skip_count = json_skip
    elif ext == ".txt":
        target_num_chunks = txt_chunks
        skip_count = txt_skip
    else: 
        target_num_chunks = other_chunks
        skip_count = 0

    if total_available_chunks < skip_count:
        selected_chunks = chunks[:target_num_chunks] if target_num_chunks > 0 else chunks
    else:
        sliced_chunks = chunks[skip_count:]
        selected_chunks = sliced_chunks[:target_num_chunks] if target_num_chunks > 0 else sliced_chunks
        
    return " ".join(selected_chunks)

def sanitize_filename(name):
    """Removes invalid filesystem characters while maintaining character sequence."""
    return "".join(c for c in name if c.isalnum() or c in (' ', '_', '-')).strip()

def get_files_at_depth(base_path, max_depth):
    """Recursively finds PDF, MD, TXT, JSON, and HTML files up to a specified subdirectory depth."""
    target_files = []
    base_path = Path(base_path)
    if not base_path.exists():
        return target_files
        
    for ext in ("*.pdf", "*.md", "*.txt", "*.json", "*.html", "*.htm"):
        for path in base_path.rglob(ext):
            if path.is_file():
                try:
                    relative_path = path.relative_to(base_path)
                    depth = len(relative_path.parts) - 1
                    if max_depth == -1 or depth <= max_depth:
                        target_files.append(path)
                except ValueError:
                    continue
    return target_files

def classify_text_detailed(text, categories_dict, model, min_len=30, max_len=250):
    """Computes cosine similarity against all categories using a user-configurable max word cap."""
    if not text:
        return "Uncategorized", 0.0, []
    
    words = text.split()
    if max_len and max_len > 0 and len(words) > max_len:
        text = " ".join(words[:max_len])
    
    doc_embedding = model.encode(text, convert_to_tensor=True)
    
    scored_categories = []
    for cat_name, cat_data in categories_dict.items():
        cat_desc = cat_data["description"]
        cat_threshold = cat_data["threshold"]
        cat_embedding = model.encode(cat_desc, convert_to_tensor=True)
        score = util.cos_sim(doc_embedding, cat_embedding).item()
        scored_categories.append((cat_name, score, cat_threshold))
    
    scored_categories.sort(key=lambda x: x[1], reverse=True)
    
    passing = [(cat, sc, th) for cat, sc, th in scored_categories if sc >= th]
    
    if passing:
        top_cat = passing[0][0]
        top_score = passing[0][1]
    else:
        top_cat = "Uncategorized"
        top_score = scored_categories[0][1] if scored_categories else 0.0
        
    return top_cat, top_score, scored_categories

# --- UI INTERFACE ---
st.title("📂 Zero-Shot File Sorter & Renamer")
st.markdown("Batch process local PDF, Markdown, Text, JSON, and HTML files, extract content, categorize using zero-shot semantic embeddings, and organize into target directories.")

col1, col2 = st.columns(2)

with col1:
    source_folder = st.text_input("Source Directory Path", value=str(Path("C:/gregs_Stuff/zeroshot-engine/AI_file_rename_and_move/Input_Files")))
    output_folder = st.text_input("Output Directory Path", value=str(Path("C:/gregs_Stuff/zeroshot-engine/AI_file_rename_and_move/Organized_Files")))
    sub_depth = st.number_input("Subdirectory Depth (-1 for unlimited, 0 for root only)", value=0, min_value=-1, max_value=10, step=1)
    dry_run = st.checkbox("Dry Run (Preview changes without moving files)", value=True)
    
    st.markdown("#### ⚙️ Chunk & Skip Configuration")
    pdf_num_chunks = st.number_input("Number of token chunks for PDF files", value=5, min_value=1, step=1)
    other_num_chunks = st.number_input("Number of token chunks for all other types including JSON/HTML", value=5, min_value=1, step=1)
    txt_num_chunks = st.number_input("Number of token chunks for TXT files", value=5, min_value=1, step=1)

with col2:
    min_summary_len = st.number_input("MIN summary length", value=30, min_value=1, step=5)
    max_summary_len = st.number_input("MAX classification word limit (0 for unlimited)", value=250, min_value=0, step=25)
    max_token_chunk = st.number_input("Maximum token chunk size for models", value=140, min_value=10, step=10)
    
    st.markdown("#### 📌 Filename Weight Boosting")
    filename_multiplier = st.number_input("Filename Duplication Multiplier (Weight Boost)", value=3, min_value=1, max_value=15, step=1)
    
    st.markdown("#### ⏭️ Skip Chunks Configuration")
    pdf_skip_chunks = st.number_input("Number of chunks to skip at beginning for PDF files", value=0, min_value=0, step=1)
    json_skip_chunks = st.number_input("Number of chunks to skip at beginning for JSON/HTML files", value=0, min_value=0, step=1)
    txt_skip_chunks = st.number_input("Number of chunks to skip at beginning for TXT files", value=0, min_value=0, step=1)

st.markdown("---")
st.markdown("### Define Categories, Descriptions & Individual Thresholds")

if "categories" not in st.session_state:
    st.session_state.categories = load_saved_categories()

# Global Threshold Group Adjustment Row
g_col1, g_col2 = st.columns([3, 1])
global_reset_val = g_col1.number_input("Set all category thresholds to value:", value=0.3, min_value=0.0, max_value=1.0, step=0.05)
if g_col2.button("Apply to All Categories"):
    for cat in st.session_state.categories:
        st.session_state.categories[cat]["threshold"] = float(global_reset_val)
    save_categories_to_disk(st.session_state.categories)
    st.success(f"All category thresholds updated to {global_reset_val}!")
    st.rerun()

st.markdown("---")

cat_container = st.container()
with cat_container:
    col_a, col_b, col_c, col_d = st.columns([2, 4, 1, 1])
    new_cat_name = col_a.text_input("Category Name", placeholder="e.g., Medical_Records")
    new_cat_desc = col_b.text_input("Semantic Description", placeholder="e.g., Medical test results...")
    new_cat_thresh = col_c.number_input("Threshold", value=0.3, min_value=0.0, max_value=1.0, step=0.05)
    if col_d.button("Add Category") and new_cat_name and new_cat_desc:
        st.session_state.categories[new_cat_name] = {"description": new_cat_desc, "threshold": new_cat_thresh}
        save_categories_to_disk(st.session_state.categories)
        st.success(f"Category '{new_cat_name}' saved successfully!")
        st.rerun()

for cat, data in list(st.session_state.categories.items()):
    c1, c2, c3, c4 = st.columns([2, 4, 1, 1])
    c1.text(cat)
    updated_desc = c2.text_input(f"Desc_{cat}", value=data["description"], key=f"desc_{cat}", label_visibility="collapsed")
    updated_thresh = c3.number_input(f"Thresh_{cat}", value=float(data["threshold"]), min_value=0.0, max_value=1.0, step=0.05, key=f"thresh_{cat}", label_visibility="collapsed")
    
    if updated_desc != data["description"] or updated_thresh != data["threshold"]:
        st.session_state.categories[cat] = {"description": updated_desc, "threshold": updated_thresh}
        save_categories_to_disk(st.session_state.categories)
        
    if c4.button("Remove", key=f"del_{cat}"):
        del st.session_state.categories[cat]
        save_categories_to_disk(st.session_state.categories)
        st.rerun()

st.markdown("---")

# Main GUI Action & Exit Buttons
btn_col1, btn_col2 = st.columns([4, 1])
with btn_col1:
    process_clicked = st.button("🚀 Process and Organize Files", type="primary")
with btn_col2:
    if st.button("🔴 Exit Application", type="secondary"):
        os._exit(0)

if process_clicked:
    src_path = Path(source_folder)
    out_path = Path(output_folder)
    
    if not src_path.exists():
        st.error(f"Source directory does not exist: {src_path}")
    else:
        sem_model = load_embedding_model()
        target_files = get_files_at_depth(src_path, sub_depth)
        
        if not target_files:
            st.warning(f"No supported files found in {src_path} at depth {sub_depth}")
        else:
            out_path.mkdir(parents=True, exist_ok=True)
            results_log = []
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            live_container = st.container()
            total_files = len(target_files)
            
            for idx, file_path in enumerate(target_files):
                rel_path = file_path.relative_to(src_path)
                status_text.text(f"Processing [{idx + 1}/{total_files}]: {rel_path}")
                
                # Extract text content from file
                extracted_text = extract_and_chunk_file_text(
                    file_path, 
                    max_pages=50, 
                    max_chunk_tokens=max_token_chunk,
                    pdf_chunks=pdf_num_chunks,
                    other_chunks=other_num_chunks,
                    txt_chunks=txt_num_chunks,
                    pdf_skip=pdf_skip_chunks,
                    json_skip=json_skip_chunks,
                    txt_skip=txt_skip_chunks
                )
                
                # Check if text is empty (image-only PDF, blank file, or empty HTML)
                if not extracted_text.strip():
                    target_category = "noTEXT"
                    score = 0.0
                    top_3_str = "N/A (Image-only or blank document)"
                else:
                    # Clean filename stem and apply duplication weight boost
                    clean_filename = file_path.stem.replace("_", " ").replace("-", " ")
                    weighted_filename_prefix = ((clean_filename + ". ") * int(filename_multiplier))
                    
                    # Combine weighted filename prefix with extracted text
                    combined_classification_text = weighted_filename_prefix + extracted_text
                    
                    # Classify using combined text
                    best_category, score, scored_categories = classify_text_detailed(
                        combined_classification_text, 
                        st.session_state.categories, 
                        sem_model, 
                        min_len=min_summary_len, 
                        max_len=max_summary_len
                    )
                    target_category = best_category
                    top_3_str = ", ".join([f"{cat}: {sc:.3f} (th:{th:.2f})" for cat, sc, th in scored_categories[:3]]) if scored_categories else "N/A"
                    
                category_dir = out_path / target_category
                safe_stem = sanitize_filename(file_path.stem)
                new_filename = f"{target_category}_{safe_stem}{file_path.suffix}"
                destination_path = category_dir / new_filename
                
                results_log.append({
                    "Original File": str(rel_path),
                    "Assigned Category": target_category,
                    "Top Score": round(score, 3),
                    "Top 3 Categories": top_3_str,
                    "Destination Path": str(destination_path),
                    "New Name": new_filename,
                    "Full extracted text": extracted_text if extracted_text.strip() else "[NO EXTRACTED TEXT - IMAGE ONLY]"
                })
                
                if not dry_run:
                    category_dir.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(file_path), str(destination_path))
                    
                progress_bar.progress((idx + 1) / total_files)
                
                with live_container:
                    st.markdown(f"### 📄 Current File: `{rel_path.name}`")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        st.markdown(f"**Assigned Category:** `{target_category}`")
                        st.markdown(f"**New Assigned Name:** `{new_filename}`")
                        st.markdown(f"**Destination Path:** `{destination_path}`")
                    with col_p2:
                        st.markdown(f"**Top 3 Category Scores:** `{top_3_str}`")
                    st.markdown(f"**Full extracted text:**")
                    st.text(extracted_text if extracted_text.strip() else "[NO EXTRACTED TEXT - IMAGE ONLY]")
                    st.markdown("---")
            
            status_text.text(f"Processing complete! Processed {total_files} files.")
            st.success(f"Processing complete! Processed {total_files} files.")
            st.dataframe(pd.DataFrame(results_log), use_container_width=True)