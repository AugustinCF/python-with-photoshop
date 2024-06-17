import pandas as pd
import win32com.client
import os

# Step 1: Load the CSV file
csv_file_path = 'C:/Users/Panda/Documents/buldo generate/data.csv'
data = pd.read_csv(csv_file_path, delimiter=',')  # Comma-separated values

# Display the first few rows of the dataframe
print(data.head())

# Ensure Photoshop is running and create a COM object
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Visible = True

# Define the path to the PSD template and the directory to save the images
template_path = 'C:/Users/Panda/Documents/buldo generate/template.psd'
output_directory = 'C:/Users/Panda/Documents/buldo generate/result'
os.makedirs(output_directory, exist_ok=True)


# Function to update text layers
def update_text_layer(doc, layer_name, text):
    try:
        textLayer = doc.ArtLayers[layer_name]
        if textLayer.Kind == 2:  # 2 corresponds to the text layer
            textItem = textLayer.TextItem

            # Apply behavior only for the specified layer name
            if layer_name == 'oras':
                # Split text into words
                words = text.split()

                if len(words) > 2:
                    # Join the first two words on one line and the rest on the second line
                    first_line = " ".join(words[:2])
                    second_line = " ".join(words[2:])
                    textItem.Contents = f"{first_line}\r{second_line}"
                elif len(words) == 2:
                    # Join the words with a line break
                    textItem.Size = 50
                else:
                    textItem.Contents = text

                # Resize text if city name is longer than 9 characters
                if len(text) > 9:
                    textItem.Size = 50  # Set font size to 50 points
            else:
                textItem.Contents = text
    except Exception as e:
        print(f"Error updating layer {layer_name}: {e}")


# Step 2: Iterate through the dataframe and create images
for index, row in data.iterrows():
    city = row['oras']  # Column name for the city
    month = row['luna']  # Column name for the month

    # Open the PSD template
    doc = psApp.Open(template_path)

    # Update the text layers
    update_text_layer(doc, 'oras', city)  # Replace 'oras' with the actual name of the text layer for the city
    update_text_layer(doc, 'luna', month)  # Replace 'luna' with the actual name of the text layer for the month

    # Save the document with a new name
    output_path = os.path.join(output_directory, f"{city}_image_{index + 1}.psd")
    doc.SaveAs(output_path)

    # Optionally export as JPEG
    jpg_options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
    jpg_options.Format = 6  # JPEG format
    jpg_options.Quality = 100  # Maximum quality
    doc.Export(ExportIn=output_path.replace(".psd", ".jpg"), ExportAs=2, Options=jpg_options)

    # Close the document without saving
    doc.Close(2)  # 2 corresponds to not saving the document

print("Images generated successfully.")
