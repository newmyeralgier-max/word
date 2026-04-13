import os

def merge_md_files(input_dir, output_file, file_list):
    print(f"Merging {len(file_list)} files into {output_file}...")
    with open(output_file, 'w', encoding='utf-8') as f_out:
        for idx, filename in enumerate(file_list):
            file_path = os.path.join(input_dir, filename)
            if not os.path.exists(file_path):
                print(f"Warning: {filename} not found!")
                continue
            
            with open(file_path, 'r', encoding='utf-8') as f_in:
                content = f_in.read()
                f_out.write(content)
                # Ensure spacing between files
                if not content.endswith('\n\n'):
                    f_out.write('\n\n')
                # Add a separator if it's not the last file (builder treats this as page break)
                if idx < len(file_list) - 1:
                    f_out.write('---\n\n')
    print("Merge complete.")

if __name__ == "__main__":
    work_dir = r"d:\1. Project\Word\.tmp"
    sections = [
        "rewritten_guide_section1.md",
        "section2_part1.md",
        "section2_part2.md",
        "section2_part3.md",
        "section2_part4.md",
        "section3.md",
        "section4.md"
    ]
    output = os.path.join(work_dir, "final_full_guide.md")
    merge_md_files(work_dir, output, sections)
