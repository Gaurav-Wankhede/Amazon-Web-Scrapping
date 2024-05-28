import pyodbc
from bs4 import BeautifulSoup
from decimal import Decimal
import glob
import openpyxl

# Function to parse HTML file and extract data
def parse_html(html_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup.find_all('div', class_='puis-card-container s-card-container s-overflow-hidden aok-relative puis-include-content-margin puis puis-vbok7i09ua2q62ek5q2l21tt78 s-latency-cf-section puis-card-border')


# Connect to SQL Server Management Studio database
conn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=DESKTOP-F8QC9QH\SQLEXPRESS;'
                      'DATABASE=System_Information;'
                      'Trusted_Connection=yes;')
cursor = conn.cursor()

# List of HTML files
html_files = glob.glob("*.html")

# Create Excel workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.append(['Name', 'Price', 'Reviews', 'Image'])

# Loop through each HTML file
for html_file in html_files:
    print("Parsing HTML file:", html_file)

    # Parse the HTML file and extract divs
    divs = parse_html(html_file)

    # Loop through each div
    for div in divs:
        # Extract Name
        try:
            name = div.find('span', class_='a-size-medium a-color-base a-text-normal').text.strip()
        except AttributeError:
            name = ""

        # Extract Price
        try:
            price_text = div.find('span', class_='a-price-whole').text.strip()
            # Convert price to decimal
            price = Decimal(price_text.replace(',', ''))
        except AttributeError:
            price = None

        # Extract Reviews
        try:
            reviews = div.find('span', class_='a-icon-alt').text.strip()
        except AttributeError:
            reviews = ""

        # Extract Image
        try:
            image = div.find('div', class_='s-product-image-container aok-relative s-text-center s-image-overlay-grey puis-image-overlay-grey s-padding-left-small s-padding-right-small puis-flex-expand-height puis puis-vbok7i09ua2q62ek5q2l21tt78').find('img')['src']
        except AttributeError:
            image = ""

        # Check if a row with similar values already exists
        cursor.execute("SELECT COUNT(*) FROM Amazon_Scrape WHERE Name = ? AND Price = ? AND Reviews = ? AND Image = ?", (name, price, reviews, image))
        count = cursor.fetchone()[0]

        if count == 0:  # If no similar row exists, insert the data
            # Insert data into the database
            cursor.execute("INSERT INTO Amazon_Scrape (Name, Price, Reviews, Image) VALUES (?, ?, ?, ?)", (name, price, reviews, image))
            print("Data inserted successfully")
        else:
            print("Data already exists. Skipping insertion.")

        # Write data to Excel sheet
        sheet.append([name, price, reviews, image])

# Save Excel workbook
wb.save("Amazon_data.xlsx")

# Commit the transaction
conn.commit()

# Close the cursor and connection
cursor.close()
conn.close()
