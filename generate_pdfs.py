import fitz  # PyMuPDF

# Open the existing PDF
pdf_document = fitz.open("medical_report.pdf")

# Loop through each row in the DataFrame (assuming you are reading from an Excel file)
import pandas as pd

# Load the Excel file
df = pd.read_excel("Karkhana.xlsx")

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Open the PDF for editing
    output_pdf = f"updated_report_{row['Employee_Name'].replace(' ', '_')}.pdf"
    pdf_document = fitz.open("medical_report.pdf")
    
    # Get the first page
    page = pdf_document[0]
    
    # Add text at specific coordinates
    page.insert_text((120, 170), f" {row['Employee_Name']}", fontsize=11)  # Modify coordinates as needed
    page.insert_text((110, 593), f" {row['Blood Group']}", fontsize=11)
    page.insert_text((285, 200), f" {row['Height']}", fontsize=11)
    page.insert_text((420, 200), f" {row['Weight']}", fontsize=11)
    page.insert_text((570, 200), f" {row['Sex']}", fontsize=11)
    page.insert_text((570, 185), f" {row['Age']}", fontsize=11)
    page.insert_text((65, 336), f" {row['Pulse']}", fontsize=11)
    page.insert_text((290, 336), f" {row['BP']}", fontsize=11)
    page.insert_text((120, 503), f" {row['Haemoglobin']}", fontsize=11)
    page.insert_text((120, 520), f" {row['Total W.B.C.']}", fontsize=11)
    page.insert_text((272, 591), f" {row['Sugar']}", fontsize=11)
    page.insert_text((573, 170), f" {row['Sr No']}", fontsize=11)
    
    
    # Save the modified PDF with a new name
    pdf_document.save(output_pdf)

    print(f"Generated PDF: {output_pdf}")

print("All PDFs generated successfully!")