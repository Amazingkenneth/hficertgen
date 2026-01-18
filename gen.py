import streamlit as st
import pandas as pd
import tempfile
import zipfile
import platform
import threading
import time
from datetime import datetime, date
from pathlib import Path
from docxtpl import DocxTemplate
from pypdf import PdfWriter
from pypinyin import pinyin, Style

# Try importing docx2pdf, handle missing dependency gracefully for UI only
try:
    from docx2pdf import convert

    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# Session State Initialization
if "is_generating" not in st.session_state:
    st.session_state.is_generating = False


def get_english_name(chinese_name):
    """
    Converts Chinese name to 'Surname, Firstname' format.
    Assumes first character is surname for simplicity in standard names.
    If already English, returns as-is.
    """
    # Check if already English (contains only ASCII letters and spaces)
    if all(ord(c) < 128 for c in chinese_name if c.isalpha()):
        return chinese_name

    if not chinese_name or len(chinese_name) < 2:
        return ""

    # Get pinyin list
    py_list = pinyin(chinese_name, style=Style.NORMAL)
    # py_list is like [['zhang'], ['yi'], ['san']]

    if not py_list:
        return ""

    # Capitalize Surname (1st char)
    surname = py_list[0][0].capitalize()

    # Join Firstname parts (rest of chars) and capitalize
    firstname = "".join([x[0] for x in py_list[1:]]).capitalize()

    return f"{surname}, {firstname}"


def parse_id_card(id_card: str):
    """
    Parses Chinese Resident ID for Raw Gender (M/F) and Date Object.
    Returns: (gender_code, dob_date_obj, is_standard)
    """
    id_str = str(id_card).strip()

    if len(id_str) != 18 or not id_str[:-1].isdigit():
        return None, None, False

    try:
        # Extract DOB
        dob_str = id_str[6:14]
        dob_obj = datetime.strptime(dob_str, "%Y%m%d").date()

        # Extract Gender
        gender_num = int(id_str[16])
        gender_code = "M" if gender_num % 2 != 0 else "F"

        return gender_code, dob_obj, True
    except:
        return None, None, False


def num_to_ordinal(n: int):
    """Converts int (10) to string ('10th')."""
    try:
        n = int(n)
        if 11 <= (n % 100) <= 13:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
        return f"{n}{suffix}"
    except:
        return str(n)


def num_to_chinese_grade(n):
    """Simple mapping for grades 1-12 to Chinese characters."""
    mapping = {
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
        10: "十",
        11: "十一",
        12: "十二",
    }
    return mapping.get(int(n), str(n))


def num_to_month(n):
    """Converts int month (1-12) to string month name."""
    months = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ]
    try:
        return months[int(n) - 1]
    except:
        return str(n)


st.set_page_config(page_title="HFI Certificate Generator", layout="wide")

st.title("HFI Certificate Generator")
st.markdown(
    """
Upload a Word template and student data to generate personalized certificates.
1. Prepare a `.docx` template with Jinja2 tags (see instructions below).
2. Upload an Excel file with student data (names, IDs, etc.).
3. Review and edit parsed data in the table.
4. Generate and download certificates in `.docx` and optional `.pdf` formats.
"""
)

# Instructions for Template
with st.expander("ℹ️ How to prepare your Word Template"):
    st.markdown(
        """
    Ensure your `.docx` file uses these exact Jinja2 tags (double curly braces):
    
    * `{{ name_zh }}` - Student's Chinese Name
    * `{{ name_en }}` - Student's English Name
    * `{{ gender_zh }}` - Gender in Chinese (男/女)
    * `{{ gender_en }}` - Gender in English (Male/Female)
    * `{{ id_type_zh }}` - ID Type in Chinese (身份证/护照/证件)
    * `{{ id_type_en }}` - ID Type in English (ID Card/Passport/Document)
    * `{{ id_number }}` - ID Number
    * `{{ dob_zh }}` - Date of Birth in Chinese (YYYY年MM月DD日)
    * `{{ dob_en }}` - Date of Birth in English (MM/DD/YYYY)
    * `{{ student_id }}` - Student ID
    * `{{ admit_year }}` - Year of Admission
    * `{{ admit_month }}` - Month of Admission (1-12)
    * `{{ admit_month_en }}` - Month name in English
    * `{{ grade_zh }}` - Grade in Chinese (九/十/十一/十二...)
    * `{{ grade_en }}` - Grade in English (9th/10th/11th/12th...)
    * `{{ date_zh }}` - Current date in Chinese (YYYY年MM月DD日)
    * `{{ date_en }}` - Current date in English (Month DD, YYYY)
    """
    )

# Warning for PDF conversion
if platform.system() not in ["Windows", "Darwin"]:
    st.warning(
        "⚠️ Linux/Cloud environment detected. PDF conversion requires MS Word installed locally (Windows/Mac). Only .docx generation will work reliably here."
    )

col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("Upload Template (.docx)", type="docx")
with col2:
    data_file = st.file_uploader("Upload Student Data (.xlsx)", type="xlsx")

if template_file and data_file:
    st.divider()

    # Load Data
    try:
        df = pd.read_excel(data_file)
        df.columns = df.columns.str.strip()
        required_cols = ["名字", "证件号码", "学号", "年级", "入学年份", "入学月份"]
        if any(c not in df.columns for c in required_cols):
            st.error(f"Missing columns. Required: {required_cols}")
            st.stop()
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
        st.stop()

    editor_data = []

    for index, row in df.iterrows():
        raw_name = str(row["名字"])
        raw_id = str(row["证件号码"])
        grade_int = row["年级"]

        # 1. Infer Gender/DOB (Raw formats for the Editor)
        gender_code, dob_date, is_standard = parse_id_card(raw_id)

        name_en = get_english_name(raw_name)

        editor_data.append(
            {
                "Name_ZH": raw_name,
                "Name_EN": name_en,
                "Student_ID": row["学号"],
                "ID_Type": "身份证" if is_standard else "证件",
                "ID_Number": raw_id,
                "Gender": gender_code,  # 'M' or 'F' or None
                "DOB": dob_date,  # Date object or None
                "Grade": grade_int,
                "Admit_Year": row["入学年份"],
                "Admit_Month": row["入学月份"],
            }
        )

    edit_df = pd.DataFrame(editor_data)

    st.subheader("📝 Verify & Edit Data")
    st.caption(
        "For Passport holders, simply select Gender (M/F) and pick the Date of Birth."
    )

    # --- SMART EDITOR ---
    # Here we use column_config to give dropdowns and date pickers
    edited_df = st.data_editor(
        edit_df,
        column_config={
            "Name_EN": "Name",
            "Name_ZH": "名字",
            "Student_ID": "学号",
            "Admit_Year": "入学年份",
            "Admit_Month": "入学月份",
            "ID_Number": "证件号码",
            "Grade": "年级",
            "Gender": st.column_config.SelectboxColumn(
                "Gender (M/F)",
                help="Select M for Male, F for Female",
                width="small",
                options=["M", "F"],
                required=True,
            ),
            "DOB": st.column_config.DateColumn(
                "出生日期",
                format="MM-DD-YYYY",
                min_value=date(2006, 1, 1),
                max_value=date.today(),
                required=True,
            ),
            "ID_Type": st.column_config.SelectboxColumn(
                "证件类型",
                options=["身份证", "护照", "证件"],
                width="small",
                required=True,
            ),
        },
        num_rows="dynamic",
    )

    st.divider()
    # Settings
    c1, c2, c3, _ = st.columns(4)
    with c1:
        generate_pdf = st.checkbox("Generate PDFs", value=True)
    with c2:
        merge_pdf = st.checkbox("Merge into one PDF", value=False)
    with c3:

        def start_generation():
            st.session_state.is_generating = True

        st.button(
            "🚀 Generate",
            on_click=start_generation,
            disabled=st.session_state.is_generating,
            type="primary",
        )

    if st.session_state.is_generating:

        # Status Containers
        status_box = st.status("Processing...", expanded=True)
        progress_bar = status_box.progress(0)

        with tempfile.TemporaryDirectory() as tmpdirname:
            output_dir = Path(tmpdirname)
            word_dir = output_dir / "word"
            pdf_dir = output_dir / "pdf"
            word_dir.mkdir()
            pdf_dir.mkdir()

            temp_template_path = output_dir / "template.docx"
            with open(temp_template_path, "wb") as f:
                f.write(template_file.getbuffer())

            # Current Date
            today = datetime.now()
            today_zh = today.strftime("%Y年%m月%d日")
            today_en = today.strftime("%B %d, %Y")

            total_rows = len(edited_df)

            # --- Generation Loop ---
            for i, row in edited_df.iterrows():

                raw_gender = str(row["Gender"]).upper()
                if raw_gender.startswith("M"):
                    gender_zh, gender_en = "男", "Male"
                elif raw_gender.startswith("F"):
                    gender_zh, gender_en = "女", "Female"
                else:
                    gender_zh, gender_en = "N/A", "N/A"  # Fallback

                id_type_zh = row["ID_Type"]
                if id_type_zh == "护照":
                    id_type_en = "Passport"
                elif id_type_zh == "证件":
                    id_type_en = "ID Document"
                else:
                    id_type_en = "ID Card"

                raw_dob = row["DOB"]
                if isinstance(raw_dob, (date, datetime)):
                    dob_zh = raw_dob.strftime("%Y年%m月%d日")
                    dob_en = raw_dob.strftime("%B %d, %Y")
                else:
                    dob_zh, dob_en = str(raw_dob), str(raw_dob)

                grade_int = row["Grade"]
                grade_zh = num_to_chinese_grade(grade_int)
                grade_en = num_to_ordinal(grade_int)

                context = {
                    "name_zh": row["Name_ZH"],
                    "name_en": row["Name_EN"],
                    "gender_zh": gender_zh,
                    "gender_en": gender_en,
                    "id_type_zh": id_type_zh,
                    "id_type_en": id_type_en,
                    "id_number": row["ID_Number"],
                    "dob_zh": dob_zh,
                    "dob_en": dob_en,
                    "student_id": row["Student_ID"],
                    "admit_year": row["Admit_Year"],
                    "admit_month": row["Admit_Month"],
                    "admit_month_en": num_to_month(row["Admit_Month"]),
                    "grade_zh": grade_zh,
                    "grade_en": grade_en,
                    "date_zh": today_zh,
                    "date_en": today_en,
                }

                # 3. Render
                doc = DocxTemplate(temp_template_path)
                doc.render(context)

                filename_base = f"{row['Student_ID']}_{row['Name_ZH']}"
                docx_path = word_dir / f"{filename_base}.docx"
                doc.save(docx_path)
                progress_bar.progress((i + 1) / total_rows * 0.1)

            if generate_pdf and DOCX2PDF_AVAILABLE:
                status_box.write("Converting to PDF (Batch Processing)...")

                conversion_error = []

                def conversion_worker():
                    try:
                        convert(str(word_dir), str(pdf_dir))
                    except Exception as e:
                        conversion_error.append(e)

                # --- 2. Start Thread ---
                thread = threading.Thread(target=conversion_worker)
                thread.start()

                # --- 3. The "Log-Based" Asymptotic Loop ---
                current_progress = 0.1  # Starting at 20%
                limit = 0.90  # The limit we approach but never quite reach
                speed_factor = 0.001

                while thread.is_alive():
                    gap = limit - current_progress
                    step = gap * speed_factor
                    current_progress += step
                    progress_bar.progress(current_progress)
                    time.sleep(0.04)

                # --- 4. Thread Finished ---
                thread.join()

                progress_bar.progress(0.9)

            merged_path = None
            if generate_pdf and merge_pdf and DOCX2PDF_AVAILABLE:
                status_box.write("Merging PDFs...")
                merger = PdfWriter()
                for pdf in sorted(list(pdf_dir.glob("*.pdf"))):
                    merger.append(str(pdf))
                merged_path = output_dir / "All_Certificates.pdf"
                merger.write(str(merged_path))
                merger.close()

            # --- Finish ---
            progress_bar.progress(1.0)
            status_box.update(
                label="✅ Generation Complete!", state="complete", expanded=False
            )

            # Zip Creation
            zip_buffer = tempfile.NamedTemporaryFile(delete=False)
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in word_dir.glob("*.docx"):
                    zf.write(f, arcname=f"Word/{f.name}")
                if generate_pdf:
                    for f in pdf_dir.glob("*.pdf"):
                        zf.write(f, arcname=f"PDF/{f.name}")
                    if merged_path and merged_path.exists():
                        zf.write(merged_path, arcname="All_Certificates_Merged.pdf")

            # Download Button
            st.download_button(
                label="📥 Download Results (.zip)",
                data=open(zip_buffer.name, "rb").read(),
                file_name="Certificates.zip",
                mime="application/zip",
            )
else:
    st.info("Please upload Template and Data to start.")
