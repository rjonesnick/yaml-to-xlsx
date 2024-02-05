import argparse
import yaml
import xlsxwriter

def parse_yaml_to_xlsx(yaml_file, xlsx_file):
    with open(yaml_file, 'r', encoding='utf-8') as file:
        questions = yaml.safe_load(file)

    workbook = xlsxwriter.Workbook(xlsx_file)
    worksheet = workbook.add_worksheet()

    difficulty_mapping = {'EASY': 2, 'MEDIUM': 3, 'HARD': 4}

    row = 1

    for question in questions:

        q_text = question['question']
        tags = ', '.join(question['tags'])
        difficulty = difficulty_mapping.get(question['difficulty'], 2)
        duration = question['duration']  

        answers = [''] * 4  
        correct_answers = ''
        for i, choice in enumerate(question['choices']):
            key = 'correct' if 'correct' in choice else 'wrong'
            answers[i] = choice[key]
            if key == 'correct':
                correct_answers += str(i + 1)

        worksheet.write(row, 0, q_text)
        for i, answer in enumerate(answers):
            worksheet.write(row, i + 1, answer)
        worksheet.write(row, 5, correct_answers)
        worksheet.write(row, 6, tags)
        worksheet.write(row, 7, difficulty)
        worksheet.write(row, 8, duration)

        row += 1

    workbook.close()

def main():
    parser = argparse.ArgumentParser(description='Convert YAML questions to an XLSX file.')
    parser.add_argument('yaml_file', type=str, help='Path to the input YAML file')
    parser.add_argument('xlsx_file', type=str, help='Path to the output XLSX file')

    args = parser.parse_args()

    parse_yaml_to_xlsx(args.yaml_file, args.xlsx_file)
    print(f"Converted '{args.yaml_file}' to '{args.xlsx_file}' successfully.")

if __name__ == "__main__":
    main()
