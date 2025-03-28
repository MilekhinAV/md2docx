# Pandoc Installation and Usage Guide

## Windows OS

1. **Download Pandoc**: [https://pandoc.org/installing.html](https://pandoc.org/installing.html)
2. **Check Installation**: Open Command Prompt and run:
   ```bash
   pandoc --version
   ```
3. **Install pywin32**: Execute the following in Command Prompt:
   ```bash
   pip install pywin32
   ```
4. **Convert Markdown to DOCX**: Navigate to the folder containing your files (`input.md` and `template.docx`) and run:
   ```bash
   pandoc -s input.md -o output.docx --reference-doc=template.docx
   ```

## macOS

1. **Install Pandoc** via Homebrew:
   ```bash
   brew install pandoc
   ```
2. **Install python-docx**:
   ```bash
   pip install python-docx
   ```
3. **Convert Markdown to DOCX**: Navigate to the folder containing your files (`input.md` and `template.docx`) and execute:
   ```bash
   pandoc -s input.md -o output.docx --reference-doc=template.docx
   ```

---

### File Descriptions:
- **`input.md`**: Your source Markdown file.
- **`output.docx`**: The name for your resulting DOCX file.
- **`template.docx`**: DOCX template containing predefined styles.

> You may use [this provided template](#) or any other `.docx` template.

**Note:** In your `output.docx`, format all code-related content using the style named **"Чанк кода"/"Code chunk"**.

