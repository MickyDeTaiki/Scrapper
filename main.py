import streamlit as st
from scrape import scrape_website, append_to_excel, process_excel_file

# Title for the website
st.title("AI Web Scraper for Detailansicht")

# Add the Text Input for the website to scrape
url = st.text_input("Enter the URL of the website to scrape:")

# Button to scrape the website and process the entries
if st.button("Scrape and Save to Excel"):
    if url.strip():
        st.write("Starting the scraping process...")
        try:
            # Scrape the website and get detailed entries
            entries = scrape_website(url, 32)
            st.write(f"Extracted {len(entries)} detailed entries.")

            # Save to Excel
            file_path = "entries.xlsx"
            append_to_excel(entries, file_path)
            process_excel_file(file_path)
            st.success(f"Entries saved to {file_path}.")

            # Provide download link
            with open(file_path, "rb") as file:
                st.download_button(
                    label="Download Excel File",
                    data=file,
                    file_name="entries.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    else:
        st.warning("Please enter a valid URL.")
