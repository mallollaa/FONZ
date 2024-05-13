import os


test_path = '/Users/manalalajmi/PycharmProjects/FonzBOT/FonzOUTPUT'

try:

    test_file_path = os.path.join(test_path, 'test_file.txt')
    with open(test_file_path, 'w') as test_file:
        test_file.write('Permission test.')
    os.remove(test_file_path)
    print(f"Successfully wrote to {test_path}")
except PermissionError:
    print(f"Permission denied for writing to {test_path}")
except Exception as e:
    print(f"An error occurred: {e}")
