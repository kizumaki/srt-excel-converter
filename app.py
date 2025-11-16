import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURATION ---
MAX_SPEAKER_NAME_LENGTH = 35 # Increased length limit for combined names

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

# --- SPEAKER VALIDATION (IMPROVED) ---

def is_valid_speaker_tag(tag):
    """
    Checks if a tag is likely a speaker name based on capitalization and allowed character rules.
    Allows for names connected by 'and' or '&'.
    """
    tag = tag.strip()
    
    # 1. Length check
    if len(tag) > MAX_SPEAKER_NAME_LENGTH:
        return False
        
    if not tag:
        return False

    # 2. Capitalization and Character check (Improved Heuristic)
    # The pattern should allow:
    # - Uppercase letters, lowercase letters
    # - Spaces
    # - The word 'and' (case-insensitive)
    # - The ampersand symbol '&'
    
    # Normalize the tag for easier checking: replace " and " with a single space, and '&' with a space.
    normalized_tag = tag.replace(' and ', ' ').replace(' and', '').replace('&', ' ').strip()
    
    if not normalized_tag:
        return False

    # Check if the first word/name in the group starts with an uppercase letter
    first_word = normalized_tag.split()[0] if normalized_tag.split() else normalized_tag
    
    if first_word[0].isalpha() and first_word[0].islower():
        # Fails if the very first word starts with a lowercase letter (e.g., "things:")
        return False
        
    # Check if the entire tag is uppercase (typical for roles/groups like "NARRATOR" or "GUYS")
    if tag.isupper():
        return True
        
    # Final check: Must primarily consist of letters, numbers, spaces, and connectors.
    # We rely heavily on the capitalization check to filter out common nouns.
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
        # Updated Regex to be more flexible with speaker names containing special characters and spaces
        time_match = re.match(r'(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})', time_line)
        if not time_match:
            continue

        time_start = time_match.group(1) 
        time_end = time_match.group(2)   

        dialogue_lines = lines[2:]

        current_speaker = None
        current_dialogue = ""
        
        # New Regex to allow spaces, & and 'and' in potential speaker names
        # Pattern: ^([\w\s&]+?): ?(.*)$
        # Group 1: The speaker name (allows letters, numbers, spaces, and '&')
        # The '?' makes the quantifier non-greedy, which is important for names.

        for line in dialogue_lines:
            line = line.strip()
            if not line:
                continue

            # Check for "Speaker:" or "Speaker: Dialogue"
            speaker_match = re.match(r'^([\w\s&]+?): ?(.*)$', line, re.DOTALL)
            
            if speaker_match:
                potential_speaker = speaker_match.group(1).strip()
                
                # VALIDATION STEP: Check if the tag is likely a speaker name (using the improved function)
                if is_valid_speaker_tag(potential_speaker):
                    
                    # 1. Finalize previous accumulated dialogue, if any
                    if current_dialogue:
                        speaker_to_use = current_speaker if current_speaker is not None else last_known_speaker
                        data.append([time_start, time_end, speaker_to_use, current_dialogue])
                    
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
                        data.append([time_start, time_end, speaker, dialogue_part])
                        
                        # Reset context for next line
                        current_speaker = None
                        current_dialogue = ""
                        
                else:
                    # Tag failed validation (e.g., "things:") -> Treat as Continuation/Unknown Dialogue
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
            data.append([time_start, time_end, speaker_to_use, current_dialogue])

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
    st.title("🎬 SRT to Excel Converter (Intelligent Speaker Recognition)")
    st.markdown("---")

    st.markdown("""
    **Hướng dẫn:** Tải lên file **SRT (.srt)** của bạn. Ứng dụng sẽ tự động phân tích và sử dụng **quy tắc viết hoa** cùng **nhận diện tên nhóm (ví dụ: Ethan & Leo)** để xác định Người nói một cách chính xác nhất.
    """)

    uploaded_file = st.file_uploader("Tải lên file SRT (.srt)", type="srt")

    if uploaded_file is not None:
        try:
            # Read and decode file content with fallback encoding
            try:
                srt_content = uploaded_file.read().decode("utf-8")
            except UnicodeDecodeError:
                srt_content = uploaded_file.read().decode("latin-1")
                
        except Exception:
            st.error("Lỗi mã hóa file. Vui lòng đảm bảo file SRT của bạn được lưu dưới dạng UTF-8.")
            return

        with st.spinner('Đang phân tích dữ liệu SRT...'):
            df_converted = parse_srt(srt_content)
        
        if df_converted.empty:
            st.error("Không thể phân tích bất kỳ phụ đề nào.")
            return

        st.subheader("Bản Xem Trước Dữ Liệu Đã Chuyển Đổi")
        
        styled_df_display = apply_styles(df_converted)
        st.dataframe(styled_df_display, use_container_width=True)

        st.markdown("---")
        
        output = io.BytesIO()
        styled_df_display.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"SRT_Converted_{timestamp}_V3.xlsx" # Updated version number V3
        
        st.download_button(
            label="💾 Tải xuống File Excel (.xlsx)",
            data=output.read(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success(f"File sẵn sàng tải xuống dưới dạng **{file_name}**!")
        
    else:
        st.info("Bắt đầu bằng cách tải lên file SRT của bạn.")

if __name__ == "__main__":
    main_app()
