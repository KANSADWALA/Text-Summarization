# Data Extraction and NLP Task 

## Overview of the Approach

The goal was to extract text from articles specified by URLs in an Excel file, perform text analysis on the extracted content, and then save the results—including hyperlinks to the original URLs—into a new Excel file.

## Steps in the Code

### 1. Import Required Libraries
The necessary libraries are imported, including:
- **pandas** for data manipulation
- **requests** for HTTP requests
- **BeautifulSoup** for web scraping
- **nltk** for natural language processing
- **openpyxl** for creating Excel files with hyperlinks

### 2. Download NLTK Resources
The NLTK library's tokenizer is downloaded to enable text processing.

### 3. Define the Article Extraction Function
The `extract_article_content` function takes a URL as input, sends an HTTP GET request to fetch the webpage, and extracts the article content:
- It uses BeautifulSoup to parse the HTML and retrieve the title (assuming it is in an `<h1>` tag) and the article text (assumed to be in `<p>` tags).
- If successful, it combines the text from the paragraphs into a single string; otherwise, it handles errors gracefully.

### 4. Load Input Data
The code reads an Excel file (`Input.xlsx`) containing URLs into a DataFrame.

### 5. Extract Article Text
The URLs are processed using the `extract_article_content` function, and the results are stored in a new column called `ExtractedText`. The updated DataFrame is saved as `Input_with_ExtractedText.xlsx`.

### 6. Unzip and Load Stop Words and Master Dictionary
The zip files containing stop words and the master dictionary (positive and negative words) are unzipped, and the contents are loaded into sets for later use in text analysis.

### 7. Text Analysis Functions
A series of functions are defined to perform various text analysis tasks, including:
- **Cleaning and Tokenizing Text**: Removing punctuation, digits, and stop words.
- **Sentiment Analysis**: Calculating positive and negative scores, polarity, and subjectivity scores.
- **Readability Metrics**: Calculating average sentence length, fog index, complex word count, and syllable count.
- **Word Statistics**: Counting total words, personal pronouns, and average word length.

### 8. Perform Analysis on Extracted Text
The code iterates over each row of the DataFrame, applying the text analysis functions to the `ExtractedText`. For each article, it calculates the relevant metrics and stores the results in a list of dictionaries.

### 9. Create Output DataFrame
A new DataFrame is created from the output data containing `URL_ID`, `URL`, and the analysis results.

### 10. Save Results to Excel with Hyperlinks
The results are saved to a new Excel file (`Output_Data_Structure.xlsx`) using openpyxl to format the URL column as clickable hyperlinks. A loop iterates through the DataFrame rows, setting the hyperlink for each URL.

### 11. Final Output
A completion message is printed to indicate that the analysis is complete and the results are saved.

## Conclusion
The approach taken involves careful structuring of the code to ensure each step logically follows from data extraction to analysis and finally to output generation. The use of functions allows for modular and reusable code, while libraries like pandas and openpyxl facilitate efficient data handling and Excel file manipulation.

This solution effectively combines web scraping, natural language processing, and data analysis, ensuring that the final output meets the specified requirements.
