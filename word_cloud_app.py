import streamlit as st
from matplotlib.colors import ListedColormap
import pandas as pd
import numpy as np
from wordcloud import WordCloud, STOPWORDS
import PyPDF2
from docx import Document
from io import BytesIO
import os
from PIL import Image

# Set page config first
st.set_page_config(
    page_title="Word Cloud Generator",
    page_icon="☁️",
    layout="wide"
)

# Custom Modern FinTech / Dark UI Styling
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* Global Background */
.stApp {
    background-color: #121315;
    background-image: 
        radial-gradient(circle at 20% 40%, rgba(255, 107, 74, 0.06) 0%, transparent 45%),
        radial-gradient(circle at 80% 80%, rgba(212, 255, 0, 0.05) 0%, transparent 45%);
    color: #F0F0F0;
    font-family: 'Inter', sans-serif;
}

/* Sidebar Background */
[data-testid="stSidebar"] {
    background-color: #1A1B1E !important;
    border-right: 1px solid rgba(255, 255, 255, 0.07);
}

/* Headers and Text */
h1, h2, h3, h4, h5, h6 {
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    color: #FFFFFF !important;
    letter-spacing: -0.5px;
}

/* Accent texts, like labels */
label, p, .stMarkdown {
    font-family: 'Inter', sans-serif !important;
    color: #A0A0A5 !important;
    font-weight: 500;
}

/* Primary Buttons (Coral/Orange) */
.stButton>button, .stDownloadButton>button {
    background-color: #FF6B4A !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 24px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600;
    padding: 0.5rem 1.5rem !important;
    box-shadow: 0 4px 14px rgba(255, 107, 74, 0.3);
    transition: all 0.2s ease;
}

.stButton>button:hover, .stDownloadButton>button:hover {
    background-color: #FF7F62 !important;
    box-shadow: 0 6px 20px rgba(255, 107, 74, 0.4);
    transform: translateY(-1px);
    color: #FFFFFF !important;
}

/* File Uploader styling - Glassmorphism */
[data-testid="stFileUploadDropzone"] {
    background-color: rgba(26, 27, 30, 0.6) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    border-radius: 12px !important;
    backdrop-filter: blur(10px);
}

/* Hide Streamlit Chrome header/footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Inputs / Dropdowns */
.stSelectbox div[data-baseweb="select"] > div, .stNumberInput input, .stTextInput input, .stTextArea textarea {
    background-color: #222327 !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
}
.stSelectbox div[data-baseweb="select"] > div:hover, .stNumberInput input:focus, .stTextArea textarea:focus {
    border-color: #D4FF00 !important;
}

/* Dataframe Styling */
[data-testid="stDataFrame"] {
    background-color: #1A1B1E !important;
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 12px;
}

/* Warning/Success messages */
.stAlert {
    background-color: rgba(255, 107, 74, 0.1) !important;
    color: #FFFFFF !important;
    border-left-color: #FF6B4A !important;
    border-radius: 8px !important;
}

/* Expander styling */
.streamlit-expanderHeader {
    background-color: #1A1B1E !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
    font-family: 'Inter', sans-serif !important;
}

/* Slider styling */
.stSlider > div > div > div > div {
    background-color: #FF6B4A !important;
}

/* Radio button styling */
.stRadio > label {
    color: #A0A0A5 !important;
}

/* Color picker styling */
[data-testid="stColorPicker"] {
    background-color: transparent !important;
}

/* Checkbox styling */
.stCheckbox label span {
    color: #A0A0A5 !important;
}

[data-testid="stImage"] {
    display: flex !important;
    justify-content: center !important;
    align-items: center !important;
    margin: 1rem 0 !important;
}

[data-testid="stImage"] img {
    max-height: 520px !important;
    max-width: 100% !important;
    width: auto !important;
    height: auto !important;
    object-fit: contain !important;
    border-radius: 12px !important;
    margin: 0 auto !important;
}

/* ===== MOBILE RESPONSIVE ===== */
.block-container {
    padding: 1.5rem !important;
    max-width: 100% !important;
}

@media (max-width: 768px) {
    .block-container {
        padding: 0.8rem !important;
    }
    h1 {
        font-size: 1.5rem !important;
    }
    h2, h3 {
        font-size: 1.1rem !important;
    }
    .stButton>button, .stDownloadButton>button {
        width: 100% !important;
        padding: 0.6rem 1rem !important;
        font-size: 0.85rem !important;
    }
    [data-testid="stFileUploadDropzone"] {
        padding: 0.8rem !important;
    }
    [data-testid="stImage"] img {
        max-width: 100% !important;
        height: auto !important;
    }
}

@media (max-width: 480px) {
    .block-container {
        padding: 0.4rem !important;
    }
    h1 {
        font-size: 1.3rem !important;
    }
    .stButton>button, .stDownloadButton>button {
        padding: 0.7rem 0.8rem !important;
        font-size: 0.8rem !important;
        min-height: 44px !important;
    }
}
</style>
""", unsafe_allow_html=True)


# Function to extract text from any uploaded file
def extract_text(file):
    if file.type == "application/pdf":
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
        return text
    elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = Document(file)
        text = "\n".join([para.text for para in doc.paragraphs])
        return text
    elif file.type == "text/plain":
        return file.read().decode("utf-8", errors="ignore")
    else:
        try:
            return file.read().decode("utf-8", errors="ignore")
        except Exception:
            st.error(f"Unsupported file type: {file.type}")
            return None


# Function to filter out stop words
def filter_stop_words(text, additional_stopwords=None):
    stop_words = set(STOPWORDS)
    if additional_stopwords:
        stop_words.update([w.lower() for w in additional_stopwords])
    filtered_text = " ".join([word for word in text.split() if word.lower() not in stop_words])
    return filtered_text


# Header & Title
st.title("☁️ Word Cloud Generator")
st.markdown("Generate stunning, customized Word Clouds from your text documents or custom input with dynamic shapes and palettes.")

# Input Mode: File Upload or Direct Text
input_tab1, input_tab2 = st.tabs(["📁 Upload Files", "✍️ Enter / Paste Text"])

all_text = ""

with input_tab1:
    uploaded_files = st.file_uploader(
        "Upload one or more files (PDF, DOCX, TXT)",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True
    )
    if uploaded_files:
        for f in uploaded_files:
            file_text = extract_text(f)
            if file_text:
                all_text += file_text + "\n"
        st.caption(f"Loaded {len(uploaded_files)} file(s).")

with input_tab2:
    sample_text = (
        "Artificial intelligence machine learning deep neural networks natural language processing "
        "computer vision data science analytics cloud computing algorithms python programming "
        "innovation technology automation robotics software development intelligence knowledge "
        "insights model training transformer architecture generative AI vision creative future "
        "data engineering pipelines big data visualization scalability performance"
    )
    text_input = st.text_area(
        "Paste or type your text here:",
        value="" if uploaded_files else sample_text,
        height=150
    )
    if not uploaded_files and text_input.strip():
        all_text = text_input.strip()

# Check if we have text
if not all_text.strip():
    st.info("👈 Upload documents above or paste text to generate your word cloud.")
else:
    # Word count and frequency table
    words = all_text.split()
    word_count = pd.Series(words).value_counts().reset_index()
    word_count.columns = ["Word", "Frequency"]

    # Collapsible word frequency table ("hide and seek")
    with st.expander("📊 View Word Frequencies Table", expanded=False):
        st.dataframe(word_count, use_container_width=True)

    # Sidebar Options
    st.sidebar.title("⚙️ Word Cloud Settings")

    # Stop Words Options
    st.sidebar.header("Stop Words")
    remove_stopwords = st.sidebar.checkbox("Remove common stop words", value=True)
    additional_stopwords = []
    if remove_stopwords:
        additional_stopwords = st.sidebar.multiselect(
            "Select additional stop words to exclude:",
            options=word_count["Word"].tolist()
        )

    # Resolution & Output Quality
    st.sidebar.header("Resolution & Quality")
    resolution_presets = {
        "Low / Fast (800 × 600)": (800, 600),
        "SD 720p (1280 × 720)": (1280, 720),
        "Full HD 1080p (1920 × 1080)": (1920, 1080),
        "2K QHD 1440p (2560 × 1440)": (2560, 1440),
        "4K Ultra HD (3840 × 2160)": (3840, 2160),
        "Square 1:1 (1080 × 1080)": (1080, 1080),
        "Square 4K (2160 × 2160)": (2160, 2160),
        "Portrait Mobile (1080 × 1920)": (1080, 1920),
        "Custom Resolution": None
    }

    selected_res = st.sidebar.selectbox(
        "Choose Output Resolution:",
        list(resolution_presets.keys()),
        index=2  # Default to Full HD 1080p
    )

    if selected_res == "Custom Resolution":
        wc_width = st.sidebar.slider("Width (px)", min_value=400, max_value=4096, value=1920, step=20)
        wc_height = st.sidebar.slider("Height (px)", min_value=400, max_value=4096, value=1080, step=20)
    else:
        wc_width, wc_height = resolution_presets[selected_res]
        st.sidebar.caption(f"📐 **Resolution:** `{wc_width} × {wc_height} px`")

    # Color Settings
    st.sidebar.header("Colors & Palette")
    transparent_bg = st.sidebar.checkbox("Transparent Background", value=False)
    if transparent_bg:
        bg_color = None
    else:
        bg_color = st.sidebar.color_picker("Background Color", "#121315")

    color_mode = st.sidebar.radio("Word Colors Mode", ["Predefined Colormap", "Custom Colors"])

    if color_mode == "Predefined Colormap":
        colormaps = [
            "viridis", "plasma", "magma", "cividis", "cool", "spring", "summer", "autumn", "winter",
            "Pastel1", "Set2", "tab10", "turbo", "Spectral", "copper"
        ]
        selected_cmap = st.sidebar.selectbox("Choose a colormap:", colormaps, index=0)
    else:
        num_colors = st.sidebar.number_input("Number of custom colors", min_value=1, max_value=6, value=3)
        custom_colors = []
        default_colors = ["#FF6B4A", "#D4FF00", "#00E5FF", "#FF3366", "#9B59B6", "#33FF57"]
        for i in range(num_colors):
            c = st.sidebar.color_picker(f"Color {i+1}", value=default_colors[i % len(default_colors)])
            custom_colors.append(c)
        selected_cmap = ListedColormap(custom_colors)

    # Shape Selection (Two-Tiered Categories)
    st.sidebar.header("Shape & Mask")
    categories = {
        "Basic Shapes": ["Rectangle (Default)", "Heart", "Star", "Chart", "Bulb", "Circle"],
        "Mechanical": ["Gear", "Wrench", "Hammer", "Screw", "Piston", "Drill", "Engine", "Turbine", "Factory", "Robotic Arm"],
        "Civil": ["Excavator", "Bulldozer", "Crane", "Bridge", "Road", "Dam", "Concrete Mixer", "Hard Hat", "Traffic Cone", "Brick"],
        "Architecture": ["Blueprint", "Skyscraper", "Pillar", "House", "Castle", "Compass Ruler", "Triangle Ruler", "Arch", "Dome", "Floor Plan"],
        "Birds": ["Eagle", "Owl", "Dove", "Penguin", "Flamingo", "Duck", "Swan", "Hummingbird", "Parrot", "Woodpecker"],
        "Animals": ["Cat", "Dog", "Elephant", "Lion", "Tiger", "Bear", "Giraffe", "Kangaroo", "Rabbit", "Fox"],
        "Insects": ["Butterfly", "Spider", "Ladybird", "Ant", "Bee", "Mosquito", "Grasshopper", "Beetle", "Caterpillar", "Dragonfly"],
        "Modern Technology": ["Laptop", "Smartphone", "Smartwatch", "Drone", "Robot", "Server", "Cloud", "Satellite", "VR", "Microchip"],
        "Custom": ["Upload Custom Image"]
    }

    selected_category = st.sidebar.selectbox("Choose Category:", list(categories.keys()))

    if selected_category == "Custom":
        selected_shape = "Upload Custom Image"
    else:
        shape_options = categories[selected_category]
        selected_shape = st.sidebar.selectbox("Choose Shape:", shape_options)

    # Performance / Speed Engine Setting
    st.sidebar.header("Performance & Speed")
    speed_mode = st.sidebar.radio(
        "Rendering Engine:",
        ["⚡ Turbo Engine (Fast & Sharp Vector 4K)", "🎯 Full Native Raster"],
        index=0,
        help="Turbo Engine optimizes computation layout while rendering fonts at full 4K output sharpness, reducing render times by over 80%."
    )

    # Compute base canvas & scale for ultra-fast generation
    if speed_mode.startswith("⚡") and max(wc_width, wc_height) > 1000:
        base_max = 960
        calc_scale = max(wc_width, wc_height) / float(base_max)
        base_w = int(round(wc_width / calc_scale))
        base_h = int(round(wc_height / calc_scale))
    else:
        calc_scale = 1.0
        base_w = wc_width
        base_h = wc_height

    mask_array = None
    if selected_shape == "Upload Custom Image":
        mask_upload = st.sidebar.file_uploader(
            "Upload Image Mask (dark silhouette on light/transparent background)",
            type=["png", "jpg", "jpeg"]
        )
        if mask_upload is not None:
            img_rgba = Image.open(mask_upload).convert("RGBA")
            img_rgba = img_rgba.resize((base_w, base_h), Image.Resampling.LANCZOS)
            bg = Image.new("RGBA", (base_w, base_h), (255, 255, 255, 255))
            img = Image.alpha_composite(bg, img_rgba).convert("L")
            mask_array = np.array(img)
            mask_array = np.where(mask_array > 128, 255, 0).astype(np.uint8)
    elif selected_shape != "Rectangle (Default)":
        shape_key = selected_shape.lower().replace(" ", "-")
        mask_path = os.path.join("masks", f"{shape_key}.png")
        if os.path.exists(mask_path):
            img_rgba = Image.open(mask_path).convert("RGBA")
            img_rgba = img_rgba.resize((base_w, base_h), Image.Resampling.LANCZOS)
            bg = Image.new("RGBA", (base_w, base_h), (255, 255, 255, 255))
            img = Image.alpha_composite(bg, img_rgba).convert("L")
            mask_array = np.array(img)
            # Solid black on white stencil
            mask_array = np.where(mask_array > 128, 255, 0).astype(np.uint8)
        else:
            st.sidebar.warning(f"Shape stencil '{shape_key}.png' not found in masks directory.")

    # Filter stop words
    filtered_text = filter_stop_words(all_text, additional_stopwords) if remove_stopwords else all_text

    # Dynamic scaling for fonts and word limits
    dynamic_max_font = int(max(100, 160 * (max(base_w, base_h) / 800.0)))
    dynamic_min_font = 4
    dynamic_max_words = 400 if mask_array is not None else 600

    # Repeat text to ensure dense coverage if a shape stencil is selected
    if mask_array is not None:
        word_count_in_text = len(filtered_text.split())
        target_words = 450
        if word_count_in_text < target_words:
            repeat_times = max(1, target_words // max(word_count_in_text, 1))
            filtered_text = " ".join([filtered_text] * repeat_times)

    # WordCloud Generation
    wc_kwargs = {
        "width": base_w,
        "height": base_h,
        "scale": calc_scale,
        "background_color": bg_color,
        "colormap": selected_cmap,
        "max_font_size": dynamic_max_font,
        "min_font_size": dynamic_min_font,
        "max_words": dynamic_max_words,
        "prefer_horizontal": 0.85,
        "contour_width": 0,
        "repeat": True
    }
    if transparent_bg:
        wc_kwargs["mode"] = "RGBA"
    if mask_array is not None:
        wc_kwargs["mask"] = mask_array

    with st.spinner(f"✨ Generating Word Cloud ({wc_width} × {wc_height} px)..."):
        wordcloud = WordCloud(**wc_kwargs).generate(filtered_text)

    st.subheader(f"🎨 Generated Word Cloud ({wc_width} × {wc_height} px)")

    # Render image cleanly without matplotlib borders
    wordcloud_img = wordcloud.to_image()
    if transparent_bg:
        wordcloud_img = wordcloud_img.convert("RGBA")

    # Ensure output image matches exact requested target dimensions
    if wordcloud_img.size != (wc_width, wc_height):
        wordcloud_img = wordcloud_img.resize((wc_width, wc_height), Image.Resampling.LANCZOS)

    # Display preview in a suitable, centered layout
    prev_col1, prev_col2, prev_col3 = st.columns([1, 6, 1])
    with prev_col2:
        st.image(wordcloud_img, use_container_width=True)

    # Action / Download Buttons
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        buf = BytesIO()
        wordcloud_img.save(buf, format="PNG")
        buf.seek(0)
        st.download_button(
            label=f"📥 Download PNG Image ({wc_width}×{wc_height})",
            data=buf,
            file_name=f"word_cloud_{wc_width}x{wc_height}.png",
            mime="image/png",
            use_container_width=True
        )

    with btn_col2:
        csv_data = word_count.to_csv(index=False)
        st.download_button(
            label="📊 Download Word Frequencies (CSV)",
            data=csv_data,
            file_name="word_frequencies.csv",
            mime="text/csv",
            use_container_width=True
        )
