import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import pandas as pd

# Preprocess the image to enhance text detection
def preprocess_image(image_path):
    image = Image.open(image_path)
    image = image.convert('L')  # Convert to grayscale
    image = image.filter(ImageFilter.SHARPEN)  # Sharpen image to enhance text
    return image

# Extract data using Tesseract
def extract_attendance(image_path):
    image = preprocess_image(image_path)
    text = pytesseract.image_to_string(image)
    
    # Split the text into lines and filter out any empty lines
    lines = text.split('\n')
    cleaned_lines = [line for line in lines if line.strip()]  # Remove empty lines
    
    data = []
    for line in cleaned_lines:
        parts = line.split()  # Split the line into parts
        if len(parts) >= 3:  # Ensure there are at least 3 components (S.N., Roll No, Name)
            sn = parts[0]
            roll_no = parts[1]
            name = " ".join(parts[2:-1])  # Combine the name parts
            sign = parts[-1]  # The last part is the signature or absence of it
            
            # Determine if the student is present or absent
            status = "Present" if sign != '' else "Absent"
            attendance_mark = 1 if status == "Present" else 0  # 1 for present, 0 for absent
            data.append([roll_no, name, status, attendance_mark])
        else:
            # Handle rows with missing or incomplete data
            if len(parts) >= 2:
                roll_no = parts[0]
                name = " ".join(parts[1:])  # Combine whatever parts are available for the name
                status = "Absent"  # If no signature column is found, mark absent
                attendance_mark = 0  # Mark absent as 0
                data.append([roll_no, name, status, attendance_mark])
    
    # Create a DataFrame from the extracted data
    new_df = pd.DataFrame(data, columns=["Roll No", "Name", "Status", "Attendance Mark"])
    
    # Path to the existing Excel sheet
    existing_excel_path = r"C:\Users\91989\Downloads\Mtech 2024 (1).xlsx"

    try:
        # Load the existing Excel file
        existing_df = pd.read_excel(existing_excel_path)
        
        # Add the 'Total Attendance' column if it doesn't exist
        if 'Total Attendance' not in existing_df.columns:
            existing_df['Total Attendance'] = 0
        
        # Merge new attendance data with the existing DataFrame
        for _, row in new_df.iterrows():
            # Find the row in the existing DataFrame where Roll No matches
            match_index = existing_df[existing_df['Roll No'] == row['Roll No']].index
            if not match_index.empty:
                # Update 'Total Attendance' by adding the new attendance mark
                existing_df.loc[match_index, 'Total Attendance'] += row['Attendance Mark']
                # Optionally update the attendance status
                existing_df.loc[match_index, 'Status'] = row['Status']
            else:
                # If the student does not exist, create a new entry as a DataFrame and use pd.concat
                new_row = pd.DataFrame({
                    'S.N.': [len(existing_df) + 1],
                    'Roll No': [row['Roll No']],
                    'Name': [row['Name']],
                    'Total Attendance': [row['Attendance Mark']],
                    'Status': [row['Status']]
                })
                # Concatenate the new row with the existing DataFrame
                existing_df = pd.concat([existing_df, new_row], ignore_index=True)
    except FileNotFoundError:
        # If the file doesn't exist, create a new DataFrame with the new data
        print(f"'{existing_excel_path}' not found, creating a new file.")
        new_df['Total Attendance'] = new_df['Attendance Mark']  # Set 'Total Attendance' initially to the attendance mark
        existing_df = new_df
    
    # Save the updated DataFrame back to Excel
    existing_df.to_excel(existing_excel_path, index=False)
    print(f"Attendance has been successfully extracted, merged, and saved to '{existing_excel_path}'.")

# Path to the image
image_path = r"C:\Users\91989\OneDrive\Desktop\SDL\img.jpeg"  # Adjust this path as needed

# Extract attendance and merge with the existing sheet
extract_attendance(image_path)
