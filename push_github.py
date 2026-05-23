import subprocess
import os
import shutil

os.chdir(r"C:\Users\deven\OneDrive\Documentos\Learning\IDN tool")

# Remove old git if exists
if os.path.exists(".git"):
    shutil.rmtree(".git")

# Initialize and configure
commands = [
    ["git", "init"],
    ["git", "config", "user.email", "deven@example.com"],
    ["git", "config", "user.name", "Deven"],
    ["git", "config", "core.pager", "cat"],
    ["git", "add", "."],
    ["git", "commit", "-m", "Initial commit: IDN/HCO Identification Pipeline Streamlit App"],
    ["git", "remote", "add", "origin", "https://github.com/devendralko/Devendra_VS_Code.git"],
    ["git", "branch", "-M", "main"],
    ["git", "push", "-u", "origin", "main", "--force"],
]

for cmd in commands:
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(f"Status: {result.returncode}")
    if result.stdout:
        print(f"Output: {result.stdout[:200]}")
    if result.stderr:
        print(f"Error: {result.stderr[:200]}")
    print("---")

print("\n✅ All git operations completed!")
