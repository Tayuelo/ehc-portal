# ehc-portal

## Merge Functionality

The `merge.py` script processes PowerPoint presentations (`.pptx`) to extract specific slides and save them as new presentations. It identifies sections of a presentation based on bold text and hyperlinks, treating these as markers for the start and end of a "song" or section.

### Features
- Extracts slides containing specific text patterns.
- Saves extracted sections as new `.pptx` files.
- Automatically names output files based on the first bold text in the section.

### Requirements
The script requires the following Python packages, which are listed in `requirements.txt`:
- `python-pptx`
- `lxml`
- `pillow`

Install the dependencies using:
```bash
pip install -r requirements.txt
```

### Usage
1. Place the source PowerPoint file (`source.pptx`) in the same directory as `merge.py`.
2. Run the script:
   ```bash
   python merge.py
   ```
3. Extracted presentations will be saved in the `./output` directory.

### Notes
- Ensure the source presentation follows the expected format (bold text to start sections, hyperlinks to end sections).
- Modify the `output_dir` parameter in the script if you want to change the output directory.
