import os

# Set the directory to search (update if your config is elsewhere)
windsurf_dir = os.path.expanduser(r"C:\Users\kwtan\.codeium\windsurf")

def find_log_files(base_dir):
    log_files = []
    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.lower().endswith('.log'):
                log_files.append(os.path.join(root, file))
    return log_files

def print_last_lines(filepath, num_lines=20):
    print(f"\n--- {filepath} ---")
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
            for line in lines[-num_lines:]:
                print(line.rstrip())
    except Exception as e:
        print(f"Error reading {filepath}: {e}")

if __name__ == "__main__":
    logs = find_log_files(windsurf_dir)
    if not logs:
        print("No .log files found in", windsurf_dir)
    else:
        for log in logs:
            print_last_lines(log)
