import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURATION ---
MAX_SPEAKER_NAME_LENGTH = 25 

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

# --- SRT PROCESSING FUNCTIONS ---

def parse_srt(srt_content):
    """
    Parses SRT content to extract Start, End timecodes, Speaker, and Dialogue.
    Implements speaker propagation and intelligent speaker detection.
    """
    data = []
    # Split content into subtitle blocks
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

        for line in dialogue_lines:
            line = line.strip()
            if not line:
                continue

            # Check for "Speaker:" or "Speaker: Dialogue"
            speaker_match = re.match(r'^([\w\s]+): ?(.*)$', line, re.DOTALL)
            
            if speaker_match:
                potential_speaker = speaker_match.group(1).strip()
                
                # ONLY PROCESS IF SPEAKER NAME IS REASONABLE LENGTH
                if len(potential_speaker) <= MAX_SPEAKER_NAME_LENGTH:
                    
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
                    # Name too long -> Treat as Continuation/Unknown Dialogue
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
    
    return df.style.apply(highlight_speaker, axis=1)

# --- STREAMLIT APP ---

def main_app():
    st.set_page_config(page_title="SRT to Excel Converter", layout="wide")
    st.title("🎬 Công Cụ Chuyển Đổi SRT sang Excel (Có Phân Biệt Speaker)")
    st.markdown("---")

    st.markdown("""
    **Hướng dẫn:**
    1. Tải lên file **SRT (.srt)** của bạn.
    2. Ứng dụng sẽ tự động phân tích và hiển thị kết quả.
    3. Nhấn nút **Tải xuống File Excel (.xlsx)** để nhận file đã được tô màu theo từng Speaker.
    """)

    # File uploader
    uploaded_file = st.file_uploader("Tải lên file SRT (.srt)", type="srt")

    if uploaded_file is not None:
        try:
            # Read and decode file content
            srt_content = uploaded_file.read().decode("utf-8")
        except UnicodeDecodeError:
            st.error("Lỗi mã hóa file. Vui lòng đảm bảo file SRT của bạn được lưu dưới dạng **UTF-8**.")
            return

        # 1. Parse content
        with st.spinner('Đang phân tích dữ liệu SRT...'):
            df_converted = parse_srt(srt_content)
        
        if df_converted.empty:
            st.error("Không thể phân tích bất kỳ phụ đề nào. Vui lòng kiểm tra định dạng file SRT.")
            return

        st.subheader("Bản Xem Trước Dữ Liệu Đã Chuyển Đổi")
        
        # 2. Apply styling for display and Excel export
        styled_df_display = apply_styles(df_converted)
        
        # Display DataFrame in Streamlit
        st.dataframe(styled_df_display, use_container_width=True)

        # 3. Download Button setup
        st.markdown("---")
        
        # Use a BytesIO buffer to hold the Excel file in memory
        output = io.BytesIO()
        
        # Save the styled DataFrame to the buffer
        styled_df_display.to_excel(output, index=False, engine='openpyxl')
        
        output.seek(0) # Rewind the buffer

        # Create file name
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"SRT_Converted_{timestamp}.xlsx"
        
        # Download button
        st.download_button(
            label="💾 Tải xuống File Excel (.xlsx)",
            data=output.read(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success(f"File đã sẵn sàng tải xuống dưới dạng **{file_name}**!")
        
    else:
        st.info("Bắt đầu bằng cách tải lên file SRT của bạn.")

if __name__ == "__main__":
    import streamlit as st # Thư viện streamlit cần được import tại đây
    main_app()
