#!/usr/bin/env python3

import pandas as pd
from collections import Counter

# Input file name is directly specified in the code
input_file = "MTHFR_histamine_data.xlsx"

try:
    # Read the Excel file
    print(f"Reading data from {input_file}...")
    df = pd.read_excel(input_file)
    
    # Process the dedicated hashtag columns
    all_hashtags = []
    
    # Find all hashtag columns
    hashtag_columns = [col for col in df.columns if col.startswith('hashtags/')]
    
    print(f"Found {len(hashtag_columns)} hashtag columns")
    
    # Extract all non-null hashtags from these columns
    for col in hashtag_columns:
        hashtags = df[col].dropna().tolist()
        # Add hashtag symbol if missing
        hashtags = ['#'+tag if not tag.startswith('#') else tag for tag in hashtags]
        all_hashtags.extend(hashtags)
    
    # Count occurrences of each hashtag
    hashtag_counts = Counter(all_hashtags)
    
    # Print results
    print(f"\nFound {len(hashtag_counts)} unique hashtags in {len(all_hashtags)} total hashtags")
    print("\nTop 20 hashtags by frequency:")
    for hashtag, count in hashtag_counts.most_common(20):
        print(f"{hashtag}: {count}")
    
    # Save results to CSV
    output_file = "hashtag_counts.csv"
    pd.DataFrame(hashtag_counts.most_common(), columns=['Hashtag', 'Count']).to_csv(output_file, index=False)
    print(f"\nSaved all hashtag counts to {output_file}")
    
except FileNotFoundError:
    print(f"Error: The file {input_file} was not found.")
except Exception as e:
    print(f"An error occurred: {e}")