import pandas as pd
import re
from collections import Counter

def extract_group_counts(file_path, search_keyword='Groups'):
    """
    Reads the input Excel/CSV file, extracts group names from a specified keyword,
    and returns a dictionary of group counts.
    """

    # Detect and read Excel/CSV file
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Please use .xlsx or .csv")

    # Find columns likely containing comments
    possible_cols = [col for col in df.columns if 'comment' in col.lower() or 'work' in col.lower()]
    if not possible_cols:
        raise ValueError("No 'comments' or 'worknotes' column found in the file.")

    # Merge all comment columns into one string per row
    comments_text = df[possible_cols].fillna('').astype(str).agg(' '.join, axis=1)

    # Regex pattern to find groups like: Groups : [code]<I>GroupName</I>[/code]
    pattern = rf'{search_keyword}\s*:\s*\[code\]<I>(.*?)</I>\[/code\]'

    all_groups = []
    for text in comments_text:
        matches = re.findall(pattern, text, flags=re.IGNORECASE)
        for match in matches:
            # Split multiple groups separated by commas
            groups = [g.strip() for g in match.split(',') if g.strip()]
            all_groups.extend(groups)

    # Normalize names (Title Case) and count occurrences
    counter = Counter(g.title() for g in all_groups)
    return counter


def save_group_counts(counter, output_file='group_counts.txt'):
    """
    Saves the results into a text file in the required format.
    """
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("Group name\tNumber of occurrences\n")
        for group, count in sorted(counter.items()):
            f.write(f"{group}\t{count}\n")
    print(f"✅ Text output saved successfully in '{output_file}'.")


def save_to_excel(counter, output_excel='group_counts.xlsx'):
    """
    Optional: Saves the output to an Excel file for easy viewing.
    """
    df_out = pd.DataFrame(sorted(counter.items()), columns=['Group name', 'Number of occurrences'])
    df_out.to_excel(output_excel, index=False)
    print(f" Excel output saved successfully in '{output_excel}'.")


if __name__ == "__main__":
    print("=== Group Extractor Automation ===")

    file_path = input("Enter input file path (.xlsx or .csv): ").strip()
    # If user forgets the extension, assume .xlsx
    if not (file_path.endswith('.xlsx') or file_path.endswith('.csv')):
        file_path += '.xlsx'

    keyword = input("Enter search keyword (default 'Groups'): ").strip() or 'Groups'

    # Run extraction
    try:
        results = extract_group_counts(file_path, keyword)
        if not results:
            print(" No matching groups found. Check your keyword or file.")
        else:
            save_group_counts(results)
            save_to_excel(results)
            print("\n Task completed successfully!")
            print("Results have been saved in both .txt and .xlsx formats.")
    except Exception as e:
        print(f" Error: {e}")
