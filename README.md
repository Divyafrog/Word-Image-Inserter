# 📷 Word Image Inserter

A Python desktop app that inserts images directly into your Microsoft Word document at your cursor position — in a grid layout with captions.

No manual formatting. No dragging and dropping one by one. Just pick your images, click insert, and they appear exactly where you want them.

---

## Requirements

- Windows
- Microsoft Word (desktop version)
- Python 3.8 or higher → https://python.org

---

## Installation

**1. Clone or download this repo**

Click the green **Code** button above and choose **Download ZIP**, then unzip it. Or clone it:

```
git clone https://github.com/Divyafrog/word-image-inserter.git
cd word-image-inserter
```

**2. Install dependencies**

```
pip install python-docx Pillow pywin32
```

---

## Usage

**1.** Open your Word document and click where you want the images to go

**2.** Save the document locally (not OneDrive) — press **Ctrl+S**, choose **Browse**, and save to your Documents folder

**3.** Run the app:

```
python inserter.py
```

**4.** In the app window:
- Type a title (optional)
- Choose columns per row (Auto fits as many as the page allows)
- Click **+ Add Images** and select your photos
- Click **Insert into Word Document**

The app saves your document, inserts the images at your cursor position, and reopens it automatically.

---

## Notes

- Your document must be saved locally — OneDrive documents are not supported
- Supported image formats: JPG, PNG, BMP, GIF, WEBP
- If you have an odd number of images the last cell will be left blank
- Captions are pre-filled with the filename — just click and type over them in Word

---

## License

MIT — free to use, modify, and share.
