import pandas as pd

def process_text_file(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()
    lines = [line for line in lines if not line.startswith('>')]
    combined_text = ''.join(lines).replace('\n', '')
    return combined_text

def find_sequence_and_surroundings(text, sequence):
    index = text.find(sequence)
    if index != -1:
        preceding_char = text[index - 1] if index > 0 else ''
        following_char = text[index + len(sequence)] if index + len(sequence) < len(text) else ''
        return (preceding_char, sequence, following_char)
    else:
        return None

def extract_sequences_from_excel(excel_path):
    df = pd.read_excel(excel_path)
    return df['Sequence to Find'].tolist()

def main_process(txt_file_path, input_excel_path, output_excel_path):
    sequences_to_find = extract_sequences_from_excel(input_excel_path)
    processed_text = process_text_file(txt_file_path)
    results = []

    for sequence in sequences_to_find:
        result = find_sequence_and_surroundings(processed_text, sequence)
        if result:
            results.append(result)

    if results:
        results_df = pd.DataFrame(results, columns=['Before', 'Sequence', 'After'])
        results_df.to_excel(output_excel_path, index=False)

# To use this script, you would call the main_process function with the appropriate paths:
# main_process(txt_file_path, input_excel_path, output_excel_path)
main_process('C:\\Users\\caoxu\\Desktop\\1.txt', "C:\\Users\\caoxu\\Desktop\\工作簿1.xlsx", "C:\\Users\\caoxu\\Desktop\\output.xlsx")