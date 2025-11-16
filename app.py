import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURATION ---
MAX_SPEAKER_NAME_LENGTH = 35 
MAX_SPEAKER_NAME_WORDS = 4 

# List of common non-speaker phrases to explicitly exclude (must be lowercase)
NON_SPEAKER_PHRASES = [
    "the only problem",
    "note",
    "warning",
    "things",
    "and on the way we came across this",
    "this is the highest swing in europe",
    "and i swear",
    "which meant",
    "the only thing is",
    "and remember",         
    "official distance",    
    "first and foremost",   
    "i said"                
]

# List of common sentence/clause starters and articles (must be lowercase)
SENTENCE_STARTER_WORDS = [
    "the", "this", "that", "and", "but", "it", "i", "we", "you", "they", "he", "she", 
    "there", "here", "what", "which", "who", "when", "why", "how", "a", "an", "my", "his", 
    "her", "your", "its", "our", "their"
]

# Color palette for distinct speaker styling (18 unique styles)
COLOR_PALETTE = [
    'background-color: #ADD8E6; color: #000000',
    'background-color: #90EE90; color: #000000',
    'background-color: #FFB6C1; color: #000000',
    'background-color: #FFFFE0; color: #000000',
    'background-color: #DDA0DD; color: #000000',
    'background-color: #AFEEEE; color: #000000',
    'background-color: #F0E68C; color: #000000',
    'background-color: #FFA07A; color: #000000',
    'background-color: #E0FFFF; color: #000000',
    'background-color: #F5F5DC; color: #000000',
    'background-color: #2F4F4F; color: #FFFFFF',
    'background-color: #191970; color: #FFFFFF',
    'background-color: #006400; color: #FFFFFF',
    'background-color: #800000; color: #FFFFFF',
    'background-color: #4B0082; color: #FFFFFF',
    'background-color: #556B2F; color: #FFFFFF',
    'background-color: #8B4513; color: #FFFFFF',
    'background-color: #36454F; color: #FFFFFF',
]

# --- TEXT CLEANUP FUNCTION ---

def clean_dialogue_text(text):
    """
    Converts HTML/XML style formatting tags (i, b, u) to text enclosed in parentheses ().
    """
    # 1. Italic/Emphasis: <i>text</i> -> (text)
    text = re.sub(r'<i[^>]*>(.*?)</i[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # 2. Bold/Strong: <b>text</b> -> (text)
    text = re.sub(r'<b[^>]*>(.*?)</b[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # 3. Underline: <u>text</u> -> (text)
    text = re.sub(r'<u[^>]*>(.*?)</u[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # Remove any other remaining unknown tags
    text = re.sub(r'<[^>]*>', '', text, flags=re.DOTALL)
    
    # Final cleanup of extra spaces
    return re.sub(r'\s+', ' ', text).strip()

# --- SPEAKER VALIDATION ---

def is_valid_speaker_tag(tag):
    """
    Checks if a tag is likely a speaker name using multiple linguistic heuristics.
    """
    tag = tag.strip()
    
    if not tag:
        return False

    # 1. Exclusion Check
    if tag.lower() in NON_SPEAKER_PHRASES:
        return False
        
    # 2. Length check
    if len(tag) > MAX_SPEAKER_NAME_LENGTH:
        return False

    # 3. Word Count Heuristic Check
    normalized_tag = tag.replace(' and ', ' ').replace(' and', '').replace('&', ' ').strip()
    
    if not normalized_tag:
        return False
        
    word_count = len(normalized_tag.split())
    if word_count > MAX_SPEAKER_NAME_WORDS:
        return False 
    
    # Get the first word of the potential tag (in lowercase)
    first_word = normalized_tag.split()[0].lower()


    # 4. Sentence Starter Rejection
    if first_word in SENTENCE_STARTER_WORDS:
        if not tag.isupper(): # Allow all-caps roles (e.g., HOST)
            return False


    # 5. Capitalization check
    if first_word[0].isalpha() and first_word[0].islower():
        return False
        
    if tag.isupper():
        return True
        
    return True


# --- SRT PROCESSING FUNCTIONS (FIXED) ---

def parse_srt(srt_content):
    """
    Parses SRT content to extract Start, End timecodes, Speaker, and Dialogue.
    Handles same-line interjections and prevents interjection speaker bleeding.
    """
    data = []
    blocks = re.split(r'\n\s*\n', srt_content.strip())
    
    last_known_speaker = "Unknown" 

    for block in blocks:
        lines = block.strip().split('\n')
        if len(lines) < 3:
            continue

        # Extract Timecodes (Start and End)
        time_line = lines[1].strip()
        time_match = re.match(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})', time_line)
        if not time_match:
            continue

        time_start = time_match.group(1) 
        time_end = time_match.group(2)   

        dialogue_lines = lines[2:]

        block_entries = []
        current_speaker = last_known_speaker
        current_dialogue = ""
        
        # Speaker found in THIS block that should be carried over to the next block.
        newly_established_speaker = None 

        # Helper function to flush accumulated dialogue to block_entries
        def flush_dialogue(speaker, dialogue):
            if dialogue.strip():
                block_entries.append([time_start, time_end, speaker, clean_dialogue_text(dialogue)])
            return ""

        for line in dialogue_lines:
            line = line.strip()
            if not line:
                continue

            # CRITICAL REGEX SPLIT: Split by the tag, capturing the tag.
            segments = re.split(r'((?:[\w\s&]+?): )', line)
            
            i = 0
            while i < len(segments):
                segment = segments[i].strip()
                i += 1
                
                if not segment:
                    continue

                # 1. Check if the segment is an actual captured speaker tag (ends with ':')
                if segment.endswith(':') and len(segment) > 1:
                    potential_speaker = segment[:-1].strip()
                    
                    if is_valid_speaker_tag(potential_speaker):
                        
                        # --- 1a. Flush accumulated dialogue before the new tag ---
                        if current_dialogue:
                            current_dialogue = flush_dialogue(current_speaker, current_dialogue)
                        
                        # --- 1b. Process the new Speaker and their dialogue ---
                        current_speaker = potential_speaker
                        
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1 
                        
                        if dialogue_segment:
                            # Append the self-contained speaker entry
                            current_dialogue = flush_dialogue(current_speaker, dialogue_segment)
                            # CRITICAL: Keep this new speaker as the current_speaker for subsequent non-tagged lines.
                            
                        # Set local state for global carry-over:
                        newly_established_speaker = current_speaker
                        
                    else:
                        # 2. Invalid tag -> Reconstruct and accumulate
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        recombined_text = segment + " " + dialogue_segment
                        current_dialogue += (" " + recombined_text if current_dialogue else recombined_text)

                # 3. Dialogue text segment -> Accumulate
                else:
                    current_dialogue += (" " + segment if current_dialogue else segment)
                        
        # Final accumulation of the block's last dialogue segment
        if current_dialogue:
            current_dialogue = flush_dialogue(current_speaker, current_dialogue)
            
        # --- CRITICAL GLOBAL STATE UPDATE (FIX FOR BLEEDING) ---
        # Append all block entries to the main data list
        data.extend(block_entries)
        
        if newly_established_speaker:
             # Only update the global state if a new speaker was successfully identified in this block.
             last_known_speaker = newly_established_speaker
             
    return pd.DataFrame(data, columns=['Start', 'End', 'Speaker', 'Dialogue'])

def apply_styles(df):
    """Applies distinct background color styling and text color per speaker."""
    unique_speakers = df['Speaker'].unique()
    
    color_map = {
        speaker: COLOR_PALETTE[i % len(COLOR_PALETTE)]
        for i, speaker in enumerate(unique_speakers)
    }

    def highlight_speaker(row):
        color_style = color_map.get(row['Speaker'], 'background-color: #FFFFFF; color: #000000')
        return [color_style] * len(row)
    
    try:
        styled_df = df.style.apply(highlight_speaker, axis=1)
        return styled_df
    except Exception:
        return df

# --- STREAMLIT APP ---

def main_app():
    st.set_page_config(page_title="SRT to Excel Converter", layout="wide")
    st.title("🎬 SRT to Excel Converter") 
    st.markdown("---")

    uploaded_file = st.file_uploader("Upload SRT File (.srt)", type="srt")

    if uploaded_file is not None:
        try:
            try:
                srt_content = uploaded_file.read().decode("utf-8")
            except UnicodeDecodeError:
                srt_content = uploaded_file.read().decode("latin-1")
                
        except Exception:
            st.error("File encoding error. Please ensure your SRT file is correctly encoded (UTF-8 is recommended).")
            return

        with st.spinner('Analyzing SRT data...'):
            df_converted = parse_srt(srt_content)
        
        if df_converted.empty:
            st.error("Could not parse any subtitles.")
            return

        st.subheader("Converted Data Preview")
        
        styled_df_display = apply_styles(df_converted)
        st.dataframe(styled_df_display, use_container_width=True)

        st.markdown("---")
        
        output = io.BytesIO()
        styled_df_display.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        original_name_base = uploaded_file.name.rsplit('.', 1)[0]
        file_name = f"{original_name_base}.xlsx"
        
        st.download_button(
            label="💾 Download Excel File (.xlsx)",
            data=output.read(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success(f"File ready for download as **{file_name}**!")
        
    else:
        st.info("Start by uploading your SRT file.")

if __name__ == "__main__":
    main_app()
