Markdown (`.md`) files are widely used due to their simplicity, readability, and compatibility across development and documentation tools. However, sometimes you'll need your content in Word (`.docx`) format instead. Below you'll find an easy guide to converting Markdown files to DOCX using Pandoc.
***

# Pandoc Installation and Usage Guide
## Windows OS

1. **Download Pandoc**: [https://pandoc.org/installing.html](https://pandoc.org/installing.html) or you can you Winget:
```powershell
winget install --id JohnMacFarlane.Pandoc --exact --source winget
```
3. **Check Installation**: Open Command Prompt and run:
   ```bash
   pandoc --version
   ```
4. **Install pywin32**: Execute the following in Command Prompt:
   ```bash
   pip install pywin32
   ```
5. **Convert Markdown to DOCX**: Navigate to the folder containing your files (`input.md` and `template.docx`) and run:
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

**Note:** In your `output.docx`, format all code-related content using the style named **"Чанк кода"/"Code chunk"**, if you use provided template. 

