import os
import shutil
import pandas as pd

# === Paths ===
input_folder = r"C:\Users\Soumy\OneDrive\Desktop\IR\input"
results_file = r"C:\Users\Soumy\OneDrive\Desktop\IR\result.xlsx"
output_folder = r"C:\Users\Soumy\OneDrive\Desktop\IR\output"
output_file = os.path.join(output_folder, "comparison_output.xlsx")
unprocessed_folder = os.path.join(output_folder, "Unprocessed")

# Create output and Unprocessed folders if they don't exist
os.makedirs(output_folder, exist_ok=True)
os.makedirs(unprocessed_folder, exist_ok=True)

# === Step 1: Read input folder files ===
input_files = [f for f in os.listdir(input_folder) if f.endswith(".docx")]
input_files_no_ext = [os.path.splitext(f)[0] for f in input_files]

# === Step 2: Read processed files from results.xlsx ===
df = pd.read_excel(results_file)
processed_files = df.iloc[:, 0].astype(str).tolist()
processed_files_no_ext = [os.path.splitext(f)[0] for f in processed_files]

# === Step 3: Compare lists ===
missing_in_results = sorted(set(input_files_no_ext) - set(processed_files_no_ext))
extra_in_results = sorted(set(processed_files_no_ext) - set(input_files_no_ext))

# === Step 4: Save comparison results ===
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pd.DataFrame({"Missing in results.xlsx": missing_in_results}).to_excel(
        writer, sheet_name="Missing_Files", index=False
    )
    pd.DataFrame({"Extra in results.xlsx": extra_in_results}).to_excel(
        writer, sheet_name="Extra_Files", index=False
    )

# === Step 5: Copy unprocessed files to Unprocessed folder ===
for file in input_files:
    file_no_ext = os.path.splitext(file)[0]
    if file_no_ext in missing_in_results:
        src_path = os.path.join(input_folder, file)
        dest_path = os.path.join(unprocessed_folder, file)
        shutil.copy2(src_path, dest_path)  # copy keeps original safe

print(f"✅ Comparison complete. Results saved to {output_file}")
print(f"✅ Unprocessed files copied to {unprocessed_folder}")
