import os

# Get the current working directory
cwd = os.getcwd()

# Loop through all the files in the directory
for filename in os.listdir(cwd):
    # Check if the file is a txt file
    if filename.endswith('.txt'):
        # Read the contents of the file
        with open(filename, 'r') as f:
            lines = f.readlines()

        # Remove duplicate lines
        unique_lines = list(set(lines))

        # Write the unique lines back to the file
        with open(filename, 'w') as f:
            f.writelines(unique_lines)
