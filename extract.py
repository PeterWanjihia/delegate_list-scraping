import mechanicalsoup
import pandas as pd

browser = mechanicalsoup.Browser()

url = "https://vote.jkuat.ac.ke"
login_page = browser.get(url)

login_html = login_page.soup

form = login_html.select("form")[0]
form.select("input")[0]["value"] = "ENB211-0049/2019"
form.select("input")[1]["value"] = "123-ABc"

profiles_page = browser.submit(form, login_page.url)

targeted_divs = []

# Find the div with the h4 element containing 'Applications'
targeted_divs = profiles_page.soup.find_all("div", class_="table-responsive")

# Assuming you are interested in the second div (index 1)
targeted_div = targeted_divs[1]

if targeted_div:
    # Extract information about the tables within the targeted div
    all_tables_info = []
    for i, table in enumerate(targeted_div.find_all("table")):
        headers = [header.text.strip() for header in table.find_all("th")]

        data = []
        for row in table.find_all("tr")[1:]:  # Skip the header row
            row_data = [cell.text.strip() for cell in row.find_all("td")]
            data.append(row_data)

        table_info = {"headers": headers, "data": data}
        all_tables_info.append(table_info)

        # Create a DataFrame and save it to Excel
        df = pd.DataFrame(data, columns=headers)
        df.to_excel(f"table_{i + 1}.xlsx", index=False)

    print("Data saved to Excel files.")
else:
    print("Targeted div not found.")
