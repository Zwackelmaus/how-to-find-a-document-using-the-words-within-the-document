import os
from docx import Document

def search_docx_files(directory, keyword):
    found_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                file_path = os.path.join(root, file)
                try:
                    document = Document(file_path)
                    for paragraph in document.paragraphs:
                        # Convert both keyword and paragraph text to lowercase for case-insensitive comparison
                        if keyword.lower() in paragraph.text.lower():
                            found_files.append(file_path)
                            break
                except Exception as e:
                    print(f"Error reading file: {file_path}")
                    print(f"Error message: {str(e)}")
    
    return found_files

# Specify the directory and keyword
directory = 'here is your path；first place you need to alter'
keyword = 'here is the word mentioned in your document; second place alterred' 

# Search for docx files
result = search_docx_files(directory, keyword)

# Display the found files
if result:
    print(f"Found {len(result)} files containing the keyword '{keyword}':")
    for file_path in result:
        print(file_path)
else:
    print(f"No files found containing the keyword '{keyword}' in the specified directory.")
