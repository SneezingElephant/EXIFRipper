import os
import piexif
from PIL import Image
from fractions import Fraction
import pandas as pd
from tkinter import filedialog, Tk, simpledialog, messagebox
import json


# Helper function to convert EXIF rational values (fractions)
def convert_rational(value):
    try:
        return float(value[0]) / float(value[1])
    except (ZeroDivisionError, TypeError):
        return None


# Function to process EXIF data from an image
def process_image(image_path, label, processed_log):
    if image_path in processed_log:
        return None  # Skip already processed images

    try:
        image = Image.open(image_path)
        exif_data = piexif.load(image.info.get("exif", b""))  # Get EXIF data if available
        exif_info = {
            "Filename": os.path.basename(image.filename),
            "Label": label,
            "Image Size": f"{image.width}x{image.height}",
            "ISO": exif_data['Exif'].get(piexif.ExifIFD.ISOSpeedRatings),
            "Aperture": None,
            "Shutter Speed": None,
            "Exposure Time": None,
            "Light Value": None,
            "Date Taken": None
        }

        # Extract exposure, aperture, and date-time data
        if exif_data['Exif']:
            # ISO Speed
            iso = exif_data['Exif'].get(piexif.ExifIFD.ISOSpeedRatings)
            exif_info["ISO"] = iso

            # Exposure Time
            exposure_time = exif_data['Exif'].get(piexif.ExifIFD.ExposureTime)
            if exposure_time:
                exif_info["Exposure Time"] = convert_rational(exposure_time)
                exif_info["Shutter Speed"] = Fraction(exif_info["Exposure Time"]).limit_denominator()

            # Aperture
            aperture_value = exif_data['Exif'].get(piexif.ExifIFD.ApertureValue)
            if aperture_value:
                exif_info["Aperture"] = convert_rational(aperture_value)

            # Date and Time Taken
            datetime_taken = exif_data['Exif'].get(piexif.ExifIFD.DateTimeOriginal)
            if datetime_taken:
                exif_info["Date Taken"] = datetime_taken.decode('utf-8')

            # Calculate Light Value (LV)
            if iso and exif_info["Aperture"] and exif_info["Exposure Time"]:
                try:
                    light_value = math.log2((exif_info["Aperture"] ** 2) / exif_info["Exposure Time"] * (100 / iso))
                    exif_info["Light Value"] = round(light_value, 2)
                except Exception as e:
                    exif_info["Light Value"] = f"Error: {e}"

        # Add image to processed log
        processed_log[image_path] = True

        return exif_info
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return None


# Process all images in a folder
def process_folder(folder_path, label, processed_log):
    image_data = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.bmp')):
            full_path = os.path.join(folder_path, filename)
            exif_info = process_image(full_path, label, processed_log)
            if exif_info:
                image_data.append(exif_info)
    return image_data


# Load folder configuration (JSON)
def load_folders():
    if os.path.exists('folder_config.json'):
        with open('folder_config.json', 'r') as config_file:
            return json.load(config_file)
    else:
        return None


# Save folder configuration (JSON)
def save_folders(folders):
    with open('folder_config.json', 'w') as config_file:
        json.dump(folders, config_file)


# Load processed log to avoid duplicates (JSON)
def load_processed_log():
    if os.path.exists('processed_log.json'):
        with open('processed_log.json', 'r') as log_file:
            return json.load(log_file)
    return {}


# Save processed log (JSON)
def save_processed_log(log):
    with open('processed_log.json', 'w') as log_file:
        json.dump(log, log_file)


# Create Excel file with EXIF data
def generate_excel(data_by_label):
    with pd.ExcelWriter('ExifRippedData.xlsx') as writer:
        for label, data in data_by_label.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=label, index=False)


# Folder selection GUI (Tkinter)
def select_folder(label_name):
    Tk().withdraw()  # Hide Tkinter root window
    folder_path = filedialog.askdirectory(title=f"Select folder for label '{label_name}'")
    return folder_path


# GUI for initial folder setup
def setup_folders():
    Tk().withdraw()  # Hide Tkinter root window
    folders = {}
    for i in range(1, 4):
        label = simpledialog.askstring("Folder Label", f"Enter name for label {i}:")
        if label:
            folder_path = select_folder(label)
            if folder_path:
                folders[label] = folder_path
            else:
                messagebox.showwarning("Warning", f"No folder selected for label '{label}'.")
        else:
            messagebox.showwarning("Warning", "Label cannot be empty.")
    save_folders(folders)


# Main function to run the program
def main():
    # Check if folder configuration exists, otherwise run setup
    folders = load_folders()
    if not folders:
        setup_folders()
        folders = load_folders()

    processed_log = load_processed_log()
    data_by_label = {}

    # Process each folder and label
    for label, folder_path in folders.items():
        data_by_label[label] = process_folder(folder_path, label, processed_log)

    # Generate Excel file
    generate_excel(data_by_label)

    # Save processed log
    save_processed_log(processed_log)

    messagebox.showinfo("Success", "EXIF data ripped successfully and saved to 'ExifRippedData.xlsx'!")

print(f"Excel file saved at: {os.path.abspath('ExifRippedData.xlsx')}")

# Run the program
if __name__ == "__main__":
    main()
