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

# Color palette for distinct speaker styling (light background colors)
COLOR_PALETTE = [
    'background-color: #ADD8E6',
    'background-color: #90EE90',
    'background-color: #FFB6C1',
    'background-color: #FFFFE0',
    'background-color: #DDA0DD',
    'background-color: #E6E6FA',
    'background-color: #AFEEEE',
    'background-color: #F0E68C'
]

# --- TEXT CLEANUP FUNCTION (UPDATED) ---

def clean_dialogue_text(text):
    """
    Converts HTML/XML style formatting tags (i, b, u) to text enclosed in parentheses ().
    """
    # Use re.IGNORECASE for case-insensitive matching (e.g., <i> or <I>)
    
    # 1. Italic/Emphasis: <i>text</i> -> (text)
    # The non-greedy matching (.*?) ensures it handles multiple tags correctly.
    text = re.sub(r'<i[^>]*>(.*?)</i[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # 2. Bold/Strong: <b>text</b> -> (text)
    text = re.sub(r'<b[^>]*>(.*?)</b[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # 3. Underline: <u>text</u> -> (text)
    text = re.sub(r'<u[^>]*>(.*?)</u[^>]*>', r'(\1)', text, flags=re.IGNORECASE | re.DOTALL)
    
    # Remove any other remaining unknown tags (to clean up the final output)
    text = re.sub(r'<[^>]*>', '', text, flags=re.DOTALL)
    
    # Final cleanup of extra spaces
    return re.sub(r'\s+', ' ', text).strip()

# --- SPEAKER VALIDATION ---

def is_valid_speaker_tag(tag):
    """
    Checks if a tag is likely a speaker name based on exclusion list,
    word count, capitalization, and allowed character rules.
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


    # 4. Capitalization check (Heuristic to filter common nouns)
    
    first_word = normalized_tag.split()[0] if normalized_tag.split() else normalized_tag
    
    if first_word[0].isalpha() and first_word[0].islower():
        return False
        
    if tag.isupper():
        return True
        
    return True


# --- SRT PROCESSING FUNCTIONS ---

def parse_srt(srt_content):
    """
    Parses SRT content to extract Start, End timecodes, Speaker, and Dialogue.
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

        current_speaker = None
        current_dialogue = ""
        
        # Regex to allow letters, numbers, spaces, and '&' in potential speaker names
        for line in dialogue_lines:
            line = line.strip()
            if not line:
                continue

            # Check for "Speaker:" or "Speaker: Dialogue"
            speaker_match = re.match(r'^([\w\s&]+?): ?(.*)$', line, re.DOTALL)
            
            if speaker_match:
                potential_speaker = speaker_match.group(1).strip()
                
                # VALIDATION STEP: Check if the tag is likely a speaker name
                if is_valid_speaker_tag(potential_speaker):
                    
                    # 1. Finalize previous accumulated dialogue, if any
                    if current_dialogue:
                        speaker_to_use = current_speaker if current_speaker is not None else last_known_speaker
                        
                        # Apply cleanup before appending
                        cleaned_dialogue = clean_dialogue_text(current_dialogue)
                        data.append([time_start, time_end, speaker_to_use, cleaned_dialogue])
                    
                    speaker = potential_speaker
                    last_known_speaker = speaker

                    # Check if it's Speaker only (Case 1)
                    dialogue_part = speaker_match.group(2).strip()
                    if not dialogue_part:
                        # Case 1: Speaker only on this line (e.g., "Tyler:")
                        current_speaker = speaker
                        current_dialogue = ""
                    else:
                        # Case 2: Speaker and dialogue on the same line (e.g., "Tyler: Good game.")
                        
                        # Apply cleanup before appending
                        cleaned_dialogue_part = clean_dialogue_text(dialogue_part)
                        data.append([time_start, time_end, speaker, cleaned_dialogue_part])
                        
                        # Reset context for next line
                        current_speaker = None
                        current_dialogue = ""
                        
                else:
                    # Tag failed validation -> Treat as Continuation/Unknown Dialogue
                    if current_dialogue:
                        current_dialogue += " " + line
                    else:
                        current_dialogue = line

            else:
                # Case 3: No explicit Speaker -> Continuation/Unknown
                if current_dialogue:
                    current_dialogue += " " + line
                else:
                    current_dialogue = line

        # Finalize the last accumulated dialogue
        if current_dialogue:
            speaker_to_use = current_speaker if current_speaker is not None else last_known_speaker
            
            # Apply cleanup before appending
            cleaned_dialogue = clean_dialogue_text(current_dialogue)
            data.append([time_start, time_end, speaker_to_use, cleaned_dialogue])

    return pd.DataFrame(data, columns=['Start', 'End', 'Speaker', 'Dialogue'])

def apply_styles(df):
    """Applies distinct background color styling per speaker."""
    unique_speakers = df['Speaker'].unique()
    color_map = {
        speaker: COLOR_PALETTE[i % len(COLOR_PALETTE)]
        for i, speaker in enumerate(unique_speakers)
    }

    def highlight_speaker(row):
        color_style = color_map.get(row['Speaker'], 'background-color: #FFFFFF')
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
            # Read and decode file content with fallback encoding
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

        # Get original file name base
        original_name_base = uploaded_file.name.rsplit('.', 1)[0]
        # Set new file name
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
