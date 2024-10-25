import pandas
import re

input_file = "coding challenge test.xlsx"
output_file = "group_counts.xlsx"  # Path to save the output Excel file

input_values = pandas.read_excel(input_file)

# Defining the pattern that will match the group codes
pattern = r"(Groups|Group Names)\s*:\s*\[code\]\s*<I>([^<]+)<\/I>\s*\[\/code\]"

def extract_groups(comments):
    matches = re.findall(pattern, comments, re.IGNORECASE)
    group_names = []
    for match in matches:
        group_text = match[1]  # Extract group content from the match tuple
        groups = group_text.split(',')
        
        for group in groups:
            group = group.strip()  # Remove leading and trailing whitespace
            if 'Role-IN-Shop0363_GM' in group:
                # Only count the part after the comma
                if len(groups) > 1:
                    group_names.append(groups[1].strip())  # Append the second part
            else:
                group_names.append(group)  # Append the regular group
    return group_names

# Filtering the rows which "additional comments" column matches the pattern and extract the group codes
input_data_filtered = input_data['Additional comments'].dropna().apply(extract_groups)

print('filtered ',input_data_filtered)

group_codes = [group.strip() for sublist in input_data_filtered for group in sublist]

# Count the occurrences of each group
group_counts = pandas.Series(group_codes).value_counts()

# Print the result
print(group_counts)

group_counts_dataframe = pandas.DataFrame({
    'Group_name': group_counts.index,
    'Number of occurrences': group_counts.values
})

# Save the DataFrame to an Excel file
# group_counts_dataframe.to_excel(output_file, index=False)

# Confirmation
print(f"Output successfully saved to {output_file}")
