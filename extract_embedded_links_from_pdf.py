# Install PyMuPDF and pandas if not already installed.

import fitz  # PyMuPDF
import pandas as pd

def extract_links_with_text(pdf_path):
    doc = fitz.open(pdf_path)
    all_links = []

    for page_num, page in enumerate(doc, start=1):
        links = page.get_links()
        for link in links:
            if 'uri' in link:
                # Get the text near the link's rectangle
                rect = fitz.Rect(link['from'])
                words = page.get_text("words")
                linked_text = " ".join(
                    word[4] for word in words if fitz.Rect(word[:4]).intersects(rect)
                )
                all_links.append({
                    "Page Number": page_num,
                    "Linked Text": linked_text.strip(),
                    "URL": link['uri']
                })

    return all_links

# Extract and save to Excel
pdf_file = "C:\\Users\\xxx\\Documents\\extract_links.pdf"
results = extract_links_with_text(pdf_file)

# Convert to DataFrame and save
df = pd.DataFrame(results)
df.to_excel("C:\\Users\\xxx\\Documents\\output_links.xlsx", index=False)

print("Links extracted and saved to output_links.xlsx")