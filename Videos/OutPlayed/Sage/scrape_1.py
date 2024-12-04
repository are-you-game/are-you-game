import pandas as pd
import os
import subprocess

# Path to the Excel file
excel_file = './Valorant Sage Abilities.xlsx'
print('a')

# Path to yt-dlp executable (adjust this if yt-dlp.exe is not in the same folder as the script)
yt_dlp_path = './yt-dlp_x86'  # Replace with the full path if needed
print('b')

# Output directory for downloaded videos
output_dir = './'
os.makedirs(output_dir, exist_ok=True)  # Create the downloads directory if it doesn't exist

# Load the Excel file
df = pd.read_excel(excel_file)

# Iterate through all columns and rows
for column in df.columns:
    for index, link in df[column].dropna().items():
        # Ensure the link is a valid string and starts with 'https://'
        if isinstance(link, str) and link.startswith('https://'):
            print(index, link)
            # Uncomment this section to enable video downloading
            # try:
            #     # Call yt-dlp executable
            #     subprocess.run(
            #         [
            #             yt_dlp_path,
            #             link,
            #             '-o', os.path.join(output_dir, f'%(title)s.%(ext)s')  # Save videos with their title
            #         ],
            #         check=True  # Raises CalledProcessError if the command fails
            #     )
            # except subprocess.CalledProcessError as e:
            #     print(f"Error downloading {link}: {e}")

print("All videos processed!")
