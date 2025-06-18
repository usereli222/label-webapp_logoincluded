import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches
import zipfile
import io
import os
import tempfile
import re
import base64

# Setup
st.set_page_config(page_title="Label Generator", layout="centered")
st.title("📦 Label Generator Web App")

# --- Load Default Logo ---
DEFAULT_LOGO_PATH = "logo.jpg"

def load_image_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

try:
    default_logo = Image.open(DEFAULT_LOGO_PATH)
    logo_base64 = load_image_base64(DEFAULT_LOGO_PATH)
except FileNotFoundError:
    st.error("Default logo file 'logo.jpg' not found.")
    st.stop()

# --- Upload Excel ---
uploaded_excel = st.file_uploader("📄 Upload Excel File", type=["xlsx", "xls"])

# --- Optional Custom Logo ---
uploaded_logo = st.file_uploader("🖼️ Optional: Upload Custom Logo", type=["jpg", "jpeg", "png"])

# --- Show logo preview ---
if uploaded_logo:
    logo = Image.open(uploaded_logo)
    st.markdown("✅ Using your **custom logo**:")
    st.image(logo, width=200)
else:
    logo = default_logo
    st.markdown(
        f"""
        ✅ Using the **default logo**:
        <div style='text-align: center; margin-bottom: 10px;'>
            <img src='data:image/png;base64,{logo_base64}' style='width:200px; border-radius:12px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);'/>
        </div>
        """, unsafe_allow_html=True
    )

# --- Generate Labels ---
if uploaded_excel and st.button("Generate Labels"):
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save Excel
        excel_path = os.path.join(temp_dir, "order.xlsx")
        with open(excel_path, "wb") as f:
            f.write(uploaded_excel.read())

        df = pd.read_excel(excel_path)
        content_list = df.values.tolist()

        alphabet = [chr(i) for i in range(97, 123)]
        l0 = [a + b for a in alphabet for b in alphabet]

        files_to_zip = []
        preview_img = None

        for i, row in enumerate(content_list):
            img = logo.copy()
            draw = ImageDraw.Draw(img)

            try:
                font = ImageFont.truetype("arial.ttf", 40)
            except:
                font = ImageFont.load_default()

            draw.text((0, 370), f"{row[0]}  {row[1]}", fill=(0, 0, 0), font=font)
            draw.text((0, 440), str(row[2]), fill=(0, 0, 0), font=font)

            img_path = os.path.join(temp_dir, f"result{l0[i]}.jpg")
            img.save(img_path)

            if i == 0:
                preview_img = img.copy()  # Save the first label as preview

            # Word doc
            doc = Document()
            doc.add_picture(img_path, width=Inches(3.6))
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", str(row[0]))
            doc_path = os.path.join(temp_dir, f"{safe_name}.docx")
            doc.save(doc_path)

            files_to_zip.append(doc_path)

        # Create ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for file in files_to_zip:
                zipf.write(file, arcname=os.path.basename(file))
        zip_buffer.seek(0)

        st.success("✅ Done! Download your documents below.")
        st.download_button("📥 Download ZIP", zip_buffer, file_name="labels.zip")

        # Show preview
        if preview_img:
            st.markdown("### 🔍 Preview of First Label:")
            st.image(preview_img, caption="First label preview", use_container_width=True)
