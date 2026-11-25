import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
import subprocess
import tempfile
import sys
import io
import os


# Using object notation
st.sidebar.title("📂 PDF Master")
st.sidebar.caption("Fast • Secure • Free")

with st.sidebar.expander("ℹ️ About"):
    st.write("Convert and compress PDF files quickly and safely.")

with st.sidebar.expander("📘 How to Use"):
    st.markdown("""
1. Select feature  
2. Upload file  
3. Choose option  
4. Click Convert  
5. Download  
""")

with st.sidebar.expander("📁 File Limits"):
    st.markdown("""
Max size: 20MB  
Formats: PDF, DOCX  
""")

with st.sidebar.expander("⚙️ Compression Guide"):
    st.markdown("""
Low → smallest size  
Medium → balanced  
High → best quality  
""")

with st.sidebar.expander("💡 Tips"):
    st.markdown("""
• Use Low for sharing  
• Use High for printing  
• Clean PDFs compress better  
""")

with st.sidebar.expander("❓ FAQ"):
    st.markdown("""
**Safe?** Yes  
**Formats?** PDF, DOCX  
**Slow?** Large files take time  
""")
    

with st.sidebar.expander("📩 Contact Us"):
    st.markdown("""
    <style>
    .contact a {
        text-decoration: none;
        color: #4da6ff;
        display: block;
        margin: 6px 0;
        font-size: 14px;
    }
    .contact a:hover {
        color: #80c1ff;
    }
    </style>

    <div class="contact">
        <a href="mailto:abhishekyadavar17@gmail.com">abhishekyadavar17@gmail.com</a>
        <a href="https://github.com/Abhishekyadav264" target="_blank">GitHub</a>
        <a href="https://www.linkedin.com/in/abhishek-yadav264/" target="_blank">LinkedIn</a> 
        <a href="https://x.com/ABHISHEKYada264" target="_blank">Twitter</a>
    </div>
    """, unsafe_allow_html=True)

st.sidebar.caption("v1.0 • Built with Streamlit")

# main code

# Windows fix
if sys.platform == "win32":
    import pythoncom
    pythoncom.CoInitialize()

if "active_tab" not in st.session_state:
    st.session_state.active_tab = None

st.set_page_config(page_title="PDF Master", page_icon="📂")

st.title("📂 PDF Master")
st.caption("Convert and compress PDFs quickly and securely")
st.divider()

col1, col2, col3 = st.columns(3)

# ---------------- PDF TO WORD ----------------
with col1:
    st.subheader("PDF to Word")
    st.image("https://play-lh.googleusercontent.com/isMQrVqUI-fhJfg08jX18j46PD7YeY_EYJuP5K7jhxG5TmzZ-b_Bgv4_vKY8EkdNXJo", width=200)
    st.caption("Convert PDF to editable Word")
    if st.button("Upload File", key="btn_pdf_word", use_container_width=True):
        st.session_state.active_tab = "pdf_to_word"

# ---------------- WORD TO PDF ----------------
with col2:
    st.subheader("Word to PDF")
    st.image("https://www.outsystems.com/Forge_CW/_image.aspx/Q8LvY--6WakOw9afDCuuGc2RYLXji3_7_XNADmUOOzg=/word-to-pdf-converter", width=200)
    st.caption("Convert Word to PDF format")
    if st.button("Upload File", key="btn_word_pdf", use_container_width=True):
        st.session_state.active_tab = "word_to_pdf"

# ---------------- COMPRESS ----------------
with col3:
    st.subheader("PDF Compress")
    st.image("https://play-lh.googleusercontent.com/VN1sjPs7gqC05rOcWcdLRv9cELQ2RhHz_XQBfCblB2RUDCSkNEUKrefC28Itj55qmNg", width=200)
    st.caption("Reduce PDF file size")
    if st.button("Upload File", key="btn_compress", use_container_width=True):
        st.session_state.active_tab = "compress"
st.divider()

# =====================================================
# PDF TO WORD
# =====================================================
if st.session_state.get("active_tab") == "pdf_to_word":
    st.header("PDF to Word")

    uploaded_file = st.file_uploader("Select file", type=["pdf"])

    if st.button("Convert") and uploaded_file:

        # ✅ Extract original filename
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{base_name}.docx"

        pdf = PdfReader(uploaded_file)

        with st.spinner("Converting..."):
            doc = Document()
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)

            word_io = io.BytesIO()
            doc.save(word_io)
            word_io.seek(0)
            
        st.success("Done ✅")
        st.download_button(
            "Download Word",
            data=word_io,
            file_name=output_filename
        )

# =====================================================
# WORD TO PDF
# =====================================================
elif st.session_state.get("active_tab") == "word_to_pdf":
    st.header("Word to PDF")

    uploaded_file = st.file_uploader("Select word file", type=["docx"])

    if st.button("Convert") and uploaded_file:
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{base_name}.pdf"

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.read())
            docx_path = tmp.name

        pdf_path = docx_path.replace(".docx", ".pdf")

        try:
            if sys.platform == "win32":
                from docx2pdf import convert
                convert(docx_path, pdf_path)
            else:
                subprocess.run([
                    "libreoffice", "--headless",
                    "--convert-to", "pdf",
                    "--outdir", tempfile.gettempdir(),
                    docx_path
                ], check=True)

            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
                
                
            st.success("Done ✅")
            st.download_button("Download PDF", pdf_bytes, output_filename)

        except Exception as e:
            st.error(f"Error: {e}")

        finally:
            if os.path.exists(docx_path):
                os.remove(docx_path)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)

# =====================================================
# PDF COMPRESSOR
# =====================================================
elif st.session_state.get("active_tab") == "compress":
    st.header("PDF Compressor")

    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

    quality = st.selectbox(
        "Select Compression Level",
        ["Low (Smallest Size)", "Medium", "High (Best Quality)"],
        index=0
    )

    quality_map = {
        "Low (Smallest Size)": "/screen",
        "Medium": "/ebook",
        "High (Best Quality)": "/printer"
    }
    gs_quality = quality_map[quality]

    def find_gs():
        import shutil
        for cmd in ["gswin64c", "gswin32c", "gs"]:
            if shutil.which(cmd):
                return cmd

        base_paths = [r"C:\Program Files\gs", r"C:\Program Files (x86)\gs"]
        for base in base_paths:
            if os.path.exists(base):
                for folder in os.listdir(base):
                    path = os.path.join(base, folder, "bin", "gswin64c.exe")
                    if os.path.exists(path):
                        return path
        return None

    if st.button("Compress") and uploaded_file:

        # ✅ Extract original filename
        base_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{base_name}_compressed.pdf"

        gs_cmd = find_gs()

        if not gs_cmd:
            st.error("Ghostscript not found")
            st.stop()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(uploaded_file.read())
            input_path = tmp.name

        output_path = input_path.replace(".pdf", "_compressed.pdf")

        try:
            subprocess.run([
                gs_cmd,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={gs_quality}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                input_path
            ], check=True)

            original = os.path.getsize(input_path) / 1024
            compressed = os.path.getsize(output_path) / 1024
            reduction = ((original - compressed) / original) * 100

            # st.success("Compression Completed!")
            st.write(f"Original: {original:.2f} KB")
            st.write(f"Compressed: {compressed:.2f} KB")
            st.write(f"Reduction: {reduction:.2f}%")

            with open(output_path, "rb") as f:
                st.success("Compression Completed!")
                st.download_button(
                    "Download Compressed PDF",
                    f,
                    output_filename
                )

        except Exception as e:
            st.error(f"Compression failed: {e}")

        finally:
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
# st.divider()
# st.caption("© 2026 PDF Master • Built with Streamlit")

st.divider()
st.markdown("""
© 2026 PDF Master • Built with Streamlit

[GitHub](https://github.com/Abhishekyadav264) | [LinkedIn ](https://www.linkedin.com/in/abhishek-yadav264/) | [Twitter](https://x.com/ABHISHEKYada264)
""")
