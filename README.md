# Email Classifier & Product Enquiry Assistant

This project automates the classification of customer emails for a fashion store, processes order requests, and generates responses to product inquiries using OpenAI's GPT models and vector embeddings.

## Features

- **Email Classification:** Automatically classifies emails as "product inquiry" or "order request" using GPT-4o.
- **Order Extraction:** Extracts product IDs and quantities from order request emails, checks stock, and generates order status.
- **Product Inquiry Responses:** Uses OpenAI embeddings to match inquiries with relevant products and generates helpful replies.
- **Excel Integration:** Reads input data from `data.xlsx` and writes results to `output.xlsx`.

## Project Structure

- [`email_classifier.py`](email_classifier.py): Main script for classifying emails and processing order requests.
- [`product_enquiry.py`](product_enquiry.py): Handles product inquiry emails and generates responses using product embeddings.
- `data.xlsx`: Input Excel file containing emails and product data.
- `output.xlsx`: Output Excel file with classification results, order statuses, and responses.
- `.env`: Stores your OpenAI API key.

## Setup

1. **Install dependencies:**
   ```sh
   pip install pandas openai python-dotenv openpyxl scikit-learn numpy
   ```

2. **Add your OpenAI API key** to the `.env` file:
   ```
   OPENAI_API_KEY = "YOUR_OPENAI_API_KEY"
   ```

3. **Prepare your input data:**
   - Place your emails and products data in `data.xlsx` with sheets named `emails` and `products`.

## Usage

1. **Classify emails and process orders:**
   ```sh
   python email_classifier.py
   ```

2. **Generate product inquiry responses:**
   ```sh
   python product_enquiry.py
   ```

3. **Check `output.xlsx`** for results:
   - `email-classification`: Email categories.
   - `order-status`: Order extraction and stock check.
   - `order-response`: Automated order replies.
   - `inquiry-response`: Product inquiry replies.

## Notes

- Requires an OpenAI API key with access to GPT-4o and embeddings.
- Make sure `data.xlsx` is formatted correctly.
- The scripts will overwrite or update `output.xlsx` as needed.

## License

MIT License

---