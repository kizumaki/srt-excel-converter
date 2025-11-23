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
    "i said",
    "here we go", "next up", "step 1", "step 2", "step 3", "and step 3", "first up", 
    "so the question is", "i was growing up", "you might be wondering", "update", 
    "nashville to miami", "all i know is", "unlike judy", "the good news is", 
    "aer lingus seat", "the true test is", "just as i suspected", "like i said", 
    "star review and said", "i told them all", "and best of all", "the point is", 
    "americans", "i was thinking", "and they go", "first of all", "second", 
    "are you like", "as a reminder", "round 2", "round 1", "round 3", "round 4", 
    "round 5", "welcome to round 3", "the question is", "quick reminder", 
    "in 2nd place", "coming up", "first stop", "next step", "and that means", 
    "hashtag", "so to be clear", "your second word", "welcome to round 6", 
    "battle finale time", "number 1", "number 2", "but the truth is", 
    "score to beat", "and your winner", "\"crafty\" and \"betcha\". coming up", 
    "next one", "keep in mind", "and it says", "you could say", "welcome to round 2", 
    "and the best part", "onto round 2", "the ride we chose", "good news is", 
    "bad news", "good news", "he thought", "3 teams remain"
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
    # Use re.IGNORECASE for case-insensitive matching
    
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
    first_word = normalized_tag.split()[0].lower() if normalized_tag.split() else normalized_tag.lower()


    # 4. Sentence Starter Rejection
    if first_word in SENTENCE_STARTER_WORDS:
        # Reject if it's not ALL CAPS (e.g., HOST, GUYS is OK, but "The problem" is not)
        if not tag.isupper():
            return False

    # 5. Final Capitalization check
    if first_word[0].isalpha() and first_word[0].islower():
        return False
        
    if tag.isupper():
        return True
        
    return True


# --- SRT PROCESSING FUNCTIONS ---

def parse_srt(srt_content):
    """
    Parses SRT content to extract Start, End timecodes, Speaker, and Dialogue.
    Handles same-line interjections and prevents interjection speaker bleeding.
    """
    data = []
    blocks = re.split(r'\n\s*\n', srt_content.strip())
    
    last_known_speaker = "Unknown" 

    # Helper function to append a row and update last_known_speaker
    def append_row_and_update_state(speaker, dialogue):
        nonlocal last_known_speaker
        data.append([time_start, time_end, speaker, clean_dialogue_text(dialogue)])
        # CRITICAL FIX: Always update global state based on the last entry created.
        last_known_speaker = speaker 

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

        current_dialogue = ""
        # Store the speaker who initiated the block's dialogue, 
        # used as the default if no new speaker is found within the block.
        block_initial_speaker = last_known_speaker
        
        # --- Multi-line and Multi-speaker processing within the block ---
        
        for line in dialogue_lines:
            line = line.strip()
            if not line:
                continue

            # Pattern to split line by (Potential_Speaker: ) and capture the delimiter
            segments = re.split(r'((?:[\w\s&]+?): )', line)
            
            i = 0
            while i < len(segments):
                segment = segments[i].strip()
                i += 1
                
                if not segment:
                    continue

                # 1. Check if the segment is a captured speaker tag (ends with ':')
                if segment.endswith(':') and len(segment) > 1:
                    speaker_tag = segment[:-1].strip()
                    
                    if is_valid_speaker_tag(speaker_tag):
                        
                        # --- Flush Accumulated Dialogue Before New Speaker ---
                        if current_dialogue:
                            # Use block_initial_speaker for the accumulated segment if this is the first flush 
                            # Otherwise, use last_known_speaker.
                            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
                            append_row_and_update_state(speaker_to_use, current_dialogue)
                            current_dialogue = "" # Flush
                            
                        # --- Process New/Interjection Speaker ---
                        speaker = speaker_tag
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1 # Advance to the dialogue segment
                        
                        # Create a new entry immediately for the interjection/new speaker
                        if dialogue_segment:
                            append_row_and_update_state(speaker, dialogue_segment)
                            
                        # If this is the FIRST speaker identified in the block, set the block context
                        if block_initial_speaker == last_known_speaker:
                             block_initial_speaker = speaker
                            
                    else:
                        # 2. Invalid speaker tag (e.g., "The only problem:") -> Reconstruct and accumulate
                        dialogue_segment = segments[i].strip() if i < len(segments) else ""
                        i += 1
                        recombined_text = segment + " " + dialogue_segment
                        
                        if current_dialogue:
                            current_dialogue += " " + recombined_text
                        else:
                            current_dialogue = recombined_text
                        
                else:
                    # 3. This is dialogue text (no tag) -> Accumulate
                    if current_dialogue:
                        current_dialogue += " " + segment
                    else:
                        current_dialogue = segment

            # End of line processing for segments. current_dialogue may hold leftovers.
            
        # Finalize the last accumulated dialogue for the entire block
        if current_dialogue:
            # If the block has previous entries, use the last known speaker.
            # If the block is entirely one accumulated dialogue, use the global last_known_speaker (block_initial_speaker).
            speaker_to_use = block_initial_speaker if not data or data[-1][0] != time_start else last_known_speaker
            
            append_row_and_update_state(speaker_to_use, current_dialogue)

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

        # --- NEW FEATURE: SPEAKER STATISTICS ---
        st.subheader("📊 Speaker Statistics")
        
        # Calculate unique speakers
        unique_speakers = df_converted['Speaker'].unique()
        
        # Filter out the "Unknown" speaker and empty strings
        actual_speakers = [s for s in unique_speakers if s not in ["Unknown", ""]]
        speaker_count = len(actual_speakers)

        st.success(f"**Tổng số Người nói được nhận dạng:** {speaker_count} người.")
        
        if speaker_count > 0:
            with st.expander("Danh sách Người nói (Speaker List):"):
                speaker_list_str = "\n".join([f"* {s}" for s in actual_speakers])
                st.markdown(speaker_list_str)
        else:
            st.info("Không tìm thấy người nói rõ ràng (ngoại trừ các đoạn hội thoại không gắn tên).")
        # --- END NEW FEATURE ---

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
