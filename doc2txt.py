import os

def get_dir_content(dir_path):
    """Lists the contents of the specified directory."""
    try:
        return os.listdir(dir_path)
    except OSError as error:
        print(f"Error accessing directory: {error}")
        return []
    
def save_content(file_path, content):
    """Save content to a file."""
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)

word_folder_path = 'D:/doc'
IRE_INDEXING_INPUTS = 'D:/txt/'

word_doc_paths = []
docs = get_dir_content(word_folder_path)
for path in docs:
    if path.endswith(".doc") or path.endswith(".docx"):
        # print(path)
        file_path = f"/{path.replace(':', '')}"
        word_doc_paths.append(file_path)

# word_folder_path = 'dbfs:/mnt/intel_analysis/sanctions/'
# docs = get_dir_content(word_folder_path)
# for path in docs:
#     if path.endswith(".doc") or path.endswith(".docx"):
#         print(path)
#         file_path = f"/{path.replace(':', '')}"
#         word_doc_paths.append(file_path)

# print(len(word_doc_paths))
# print(word_doc_paths)

# import docx2txt 
import pypandoc
from pathlib import Path

error_c = 0

for word_doc in word_doc_paths:
    path = Path(word_doc)
    filename = path.stem.replace("[", "").replace("]", "").replace("{", "").replace("}", "") + "-doc.txt"
    parent = path.parts[0]
    # print(f'filename: {filename}')

    # full_path = Path(word_folder_path) / word_doc
    full_path = word_folder_path + word_doc
    save_path = IRE_INDEXING_INPUTS + filename
    # save_path = Path(IRE_INDEXING_INPUTS).joinpath(parent, filename)
    print(f'full_path: {full_path}')
    print(f'save_path: {save_path}')


    try:
        all_text = pypandoc.convert_file(full_path, 'plain')
        print(f"Saving {save_path} ...")
        save_content(save_path, all_text)

    except Exception as e:
        print(f"Skipping {word_doc} due to exception {e}")
        error_c = error_c + 1

# print(error_c)