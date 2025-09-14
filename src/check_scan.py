import re
import os
import pandas as pd
import cv2
import easyocr
from transformers import TrOCRProcessor, VisionEncoderDecoderModel, logging
from word2number import w2n
from PIL import Image
from datetime import datetime

# Set the logging level to error (or critical) to suppress warnings
logging.set_verbosity_error()


## Main Functions

def extract_amount(image_path, processor, model):
    """
    Extracts a numeric amount from a specified region in an image using OCR.

    Inputs:
    - image_path (str): Path to the image file containing the text to extract.
    - processor (object): Preprocessor for converting image data into model input format.
    - model (object): Text recognition model for extracting text from images.

    Outputs:
    - Extracted amount (str): Numeric string of the recognized amount, e.g., "125".
    - None: Returned if no valid amount is recognized.    
    """

    # Load the image
    image = cv2.imread(image_path)
    if image is None:
        raise ValueError(f"Image not found at {image_path}")
        
    (height, width, _) = image.shape

    # Define ROI parameters
    y_start = int(height * 0.45)
    y_end = int(height * 0.565)
    x_start = int(width * 0.05)
    initial_x_end = int(width * 0.45)
    final_x_end = int(width * 0.14)
    step = int(width * 0.01)

    # Iterate over decreasing ROI widths
    for x_end in range(initial_x_end, final_x_end - 1, -step):
        # Extract the ROI
        roi = image[y_start:y_end, x_start:x_end]

        # Convert the ROI to RGB
        roi_rgb = cv2.cvtColor(roi, cv2.COLOR_BGR2RGB)

        # Convert to PIL Image
        roi_image = Image.fromarray(roi_rgb)

        # Preprocess the image
        pixel_values = processor(images=roi_image, return_tensors="pt").pixel_values

        # Generate text
        generated_ids = model.generate(pixel_values, max_new_tokens=20)
        recognized_text = processor.batch_decode(generated_ids, skip_special_tokens=True)[0]

        # Attempt to convert recognized text to a number
        try:
            amount_value = w2n.word_to_num(recognized_text)
            return str(amount_value)
        except ValueError:
            # If conversion fails, continue to the next ROI width
            continue

    # If all attempts fail, raise an error
    return None


def extract_address_and_names(name_text):
    """
    Separates names and addresses from a given text, treating text starting with numbers as the address.

    Inputs:
    - name_text (str): Text containing names and address information separated by commas.

    Outputs:
    - names_text (str or None): Concatenated names as a single string, or None if no names are found.
    - address_text (str or None): Concatenated address as a single string, or None if no address is found.
    """
    # Split the text into parts by commas
    parts = name_text.split(",")
    
    # Initialize variables for names and address
    names = []
    address_parts = []

    # Flag to indicate if we are in the address section
    found_address = False

    # Loop through parts to classify as name or address
    for part in parts:
        part = part.strip()  # Clean up extra whitespace

        if not found_address and re.match(r'^\d+', part):
            # First part that starts with a numeric value marks the start of the address
            found_address = True
        
        if found_address:
            # Add to address if it's part of the address
            address_parts.append(part)
        else:
            # Otherwise, add to names
            names.append(part)

    # Join names and address into single strings
    names_text = ", ".join(names) if names else None
    address_text = ", ".join(address_parts) if address_parts else None
    
    return names_text, address_text


def extract_check_info(image_path, reader, processor, model):
    """
    Extracts names, address, and check donation amount information from a given image.

    Inputs:
    - image_path (str): Path to the image file containing the check information.
    - reader (object): OCR reader for extracting text from specific regions in the image.
    - processor (object): Preprocessor for converting image data into model input format (for amount extraction).
    - model (object): Text recognition model for extracting the donation amount.

    Outputs:
    - names_text (str or None): Concatenated names as a single string, or None if no names are found.
    - address_text (str or None): Concatenated address as a single string, or None if no address is found.
    - amount_text (str or None): Numeric string of the donation amount, e.g., "125", or None if no amount is recognized.
    """
    
    # Load the image using OpenCV
    image = cv2.imread(image_path)
    
    # Get the height and width of the grayscale image
    (height, width, _) = image.shape
    
    ## Name and Address
    # Define the region of interest (ROI) for names (top-left corner)
    roi_name = image[0:int(height * 0.3), 0:int(width * 0.5)]
    
    # Use selected OCR model to read text from the names ROI
    name_results = reader.readtext(roi_name)
    name_text = ", ".join([res[1] for res in name_results])  # Combine all detected texts

    names_text, address_text = extract_address_and_names(name_text)
    
    ## Check Donation Amount
    amount_text = extract_amount(image_path, processor, model)
    
    # Return the extracted text for names and donation amount
    return names_text, address_text, amount_text


def process_checks(check_directory, reader, processor, model, output_directory):
    """
    Process all scanned check images in a directory to extract Name, Address, and Donation Amount.
    
    Inputs:
    - check_directory (str): Path to the directory containing scanned check images (.tif files).
    - reader (easyocr.Reader): Initialized EasyOCR reader for text recognition.
    - processor (object): Preprocessor for converting image data into model input format (for amount extraction).
    - model (object): Text recognition model for extracting the donation amount.
    - output_directory (str): Output directory where the csv file will be saved. 
    """
    # List to store extracted data for each check
    data = []

    # Iterate through all files in the directory
    for file in os.listdir(check_directory):
        if file.endswith(".tif"):  # Process only .tif files
            file_path = os.path.join(check_directory, file)

            cleaned_names, cleaned_address, cleaned_amount = extract_check_info(file_path, reader, processor, model)

            # Append the extracted information to the data list
            data.append({
                "Names": cleaned_names,
                "Address": cleaned_address,
                "DonationAmount": cleaned_amount
            })

    # Convert the collected data into a Pandas DataFrame and export the csv file. 
    df = pd.DataFrame(data)
    filename = "check_data_" + datetime.today().strftime("%m_%d_%Y") + ".csv"
    df.to_csv(output_directory+filename, index=False)


## Execution
if __name__ == "__main__":

    # Set main models 
    print("Loading main models...")
    reader = easyocr.Reader(['en'])
    trocr_processor = TrOCRProcessor.from_pretrained('microsoft/trocr-base-handwritten')
    trocr_model = VisionEncoderDecoderModel.from_pretrained('microsoft/trocr-base-handwritten')
    print("Model loading complete.")

    # Set the directory
    date_str = datetime.today().strftime("%m-%d-%Y")
    check_image_dir = os.path.join(r"C:\Users\hkmcc\Documents\Finance_Automation\check_outputs", date_str, "image")
    check_output_dir = "C:\\Users\\hkmcc\\Documents\\Finance_Automation\\check_scanned_data\\"
 
    # Process scanned images and save the csv file 
    print("Processing scanned check images...")
    process_checks(check_directory=check_image_dir, 
                   reader=reader, 
                   processor=trocr_processor, 
                   model=trocr_model,
                   output_directory=check_output_dir)
    
    print("Processing and file export complete.")