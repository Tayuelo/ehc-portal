import os

def has_hyperlink(run):
    try:
        return bool(run.hyperlink and run.hyperlink.address)
    except KeyError:
        return True

def print_file_grid(folder_path, columns=3, spacing=60):
    try:
        files = os.listdir(folder_path)
        files = sorted([f for f in files if f.endswith(".pptx")])

        print("\nAvailable files in ./output:")

        for i in range(0, len(files), columns):
            row = files[i:i+columns]
            print("".join(f.ljust(spacing) for f in row))

    except FileNotFoundError:
        print("'./output' folder not found.")