import os

directory = './database_txt/'

for filename in os.listdir(directory):
    if filename.endswith('.txt'):
        filepath = os.path.join(directory, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        lines = [line for line in lines if line.strip()] 
        with open(filepath, 'w', encoding='utf-8') as f:
            f.writelines(lines)