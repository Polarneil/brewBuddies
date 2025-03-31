import openpyxl
import random
import os


def read_simple_xlsx(filepath):
    """Reads an XLSX file and returns the rows as a list of tuples, skipping the first row (headers)."""
    try:
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        rows = [tuple(row) for row in sheet.iter_rows(values_only=True)][1:]  # Skip the first row
        return rows
    except FileNotFoundError:
        print(f"File not found: {filepath}")
        return []
    except Exception as e:
        print(f"An error occurred: {e}")
        return []


def create_random_pairings(rows, existing_pairs):
    """Creates random pairings from the list of rows using only the first column (name), avoiding existing pairs."""
    names = [row[0] for row in rows]  # Extract the first column (name)
    random.shuffle(names)
    pairings = []
    i = 0
    while i < len(names) - 1:
        pair = (names[i], names[i + 1])
        if pair not in existing_pairs and (pair[1], pair[0]) not in existing_pairs:
            pairings.append(pair)
            i += 2
        else:
            random.shuffle(names)
            i = 0
    if len(names) % 2 == 1:
        if len(pairings) > 0:
            pairings[-1] = (pairings[-1][0], pairings[-1][1], names[-1])  # Make a group of three
        else:
            pairings.append((names[-1],))  # Handle the case with only one name
    return pairings


def save_pairings_to_xlsx(pairings, filepath):
    """Saves the pairings to a new XLSX file."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for i, pairing in enumerate(pairings, start=1):
        for j, name in enumerate(pairing, start=1):
            sheet.cell(row=i, column=j, value=name)
    workbook.save(filepath)


def get_next_available_filename(directory, base_name, extension):
    """Finds the next available filename in the specified directory."""
    i = 1
    while True:
        filename = f"{base_name}_{i}.{extension}"
        if not os.path.exists(os.path.join(directory, filename)):
            return os.path.join(directory, filename)
        i += 1


def read_existing_pairs(directory, base_name, extension):
    """Reads all existing pairing files and returns a set of existing pairs."""
    existing_pairs = set()
    i = 1
    while True:
        filename = f"{base_name}_{i}.{extension}"
        filepath = os.path.join(directory, filename)
        if not os.path.exists(filepath):
            break
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        for row in sheet.iter_rows(values_only=True):
            if len(row) == 2:
                pair = (row[0], row[1])
                if pair not in existing_pairs and (pair[1], pair[0]) not in existing_pairs:
                    existing_pairs.add(pair)
            elif len(row) == 3:
                pairs = [(row[0], row[1]), (row[0], row[2]), (row[1], row[2])]
                for pair in pairs:
                    if pair not in existing_pairs and (pair[1], pair[0]) not in existing_pairs:
                        existing_pairs.add(pair)
        i += 1
    return existing_pairs


# Example usage:
xlsx_path = "../data/data_members.xlsx"
all_rows = read_simple_xlsx(xlsx_path)
if all_rows:
    output_dir = "../data/pairings"
    os.makedirs(output_dir, exist_ok=True)
    all_existing_pairs = read_existing_pairs(output_dir, "pairing", "xlsx")
    all_pairings = create_random_pairings(all_rows, all_existing_pairs)
    output_path = get_next_available_filename(output_dir, "pairing", "xlsx")
    save_pairings_to_xlsx(all_pairings, output_path)
