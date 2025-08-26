# In-House Guest List Transcriber 🏨

A Streamlit web application designed to automate the guest list management process for Best Western at Firestone. This tool transforms raw Excel exports into beautifully formatted guest lists, eliminating manual transcription work.

## 🌟 Features

- **Automated Processing**: Converts raw Excel exports into properly formatted guest lists
- **Smart Formatting**:
  - Two-column layout for efficient space use
  - Bold headers and room numbers
  - Clean borders and spacing
  - Organized guest information by room numbers
- **User-Friendly Interface**:
  - Simple upload/download process
  - Clear step-by-step instructions
  - Visual feedback with animations
- **Error Handling**: Robust error checking and user feedback

## 🚀 Getting Started

### Prerequisites

```bash
pip install streamlit pandas openpyxl xlrd
```

### Running the Application

```bash
streamlit run auto_bucket.py
```

## 📋 How to Use

1. 🚪 Open the '**Front Office**' user tab in your hotel management system
2. 📊 Select '**Reports**' from the top navigation bar
3. 📑 Click '**Front Office**' tab that appears
4. 📈 Hover over '**Reports**' (Bar Graph Icon)
5. 👥 Select '**In House Guest**' from dropdown menu
6. 🔄 Click '**Refresh**' button
7. ⬇️ Click '**Export**' button to save as Excel
8. 📘 Upload the exported file to this Streamlit app
9. ⚡ Download your formatted guest list!

## 📄 Output Format

The generated Excel file includes:
- Current date header
- Room numbers in bold
- Guest names and rates
- Clean, professional formatting
- Landscape orientation for better printing

## ⚙️ Technical Details

The application:
- Uses `pandas` for data processing
- Leverages `openpyxl` for Excel formatting
- Implements `streamlit` for the web interface
- Handles various edge cases and data validation

## 🛠️ Error Handling

The application handles:
- Invalid file formats
- Missing data
- Incorrect sheet names
- Data conversion issues
- General exceptions with user-friendly messages

## 👨‍💻 Developer

Created by Esteban C Loetz

---

*Where Python Wiz Meets Data Biz! 🌟*
