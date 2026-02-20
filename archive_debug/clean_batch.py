import os
import glob

OUTPUT_DIR = "Diplomas_Batch"

def clean():
    files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
    print(f"Deleting {len(files)} files...")
    for f in files:
        try:
            os.remove(f)
        except Exception as e:
            print(f"Error deleting {f}: {e}")

if __name__ == "__main__":
    clean()
