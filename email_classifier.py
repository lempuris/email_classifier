import pandas as pd
from openai import OpenAI
from dotenv import load_dotenv
from openpyxl import load_workbook
import re
import os
from datetime import datetime

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
status = "Processed"

# Load excel spreadsheet
xls = pd.ExcelFile("data.xlsx")
emails_df = xls.parse("emails")[["email_id", "subject", "message"]]
products_df = xls.parse("products")

few_shot_examples = [
    {"subject": "Need details on Vintage Sunglasses", "message": "Hello, I saw your vintage sunglasses and wanted to know if they're suitable for summer wear?", "category": "product inquiry"},
    {"subject": "Order Cozy Shawl", "message": "Hi, I’d like to order two CSH1098 Cozy Shawls. Please confirm availability.", "category": "order request"},
    {"subject": "Help with choosing a bag", "message": "I need a fashionable and spacious bag for winter. Can you recommend something?", "category": "product inquiry"},
    {"subject": "Buy Infinity Scarves Order", "message": "Hi, I'd like to order three to four SFT1098 Infinity Scarves.", "category": "order request"}
]

def build_prompt(email):
    shots = "\n".join(
        f"Subject: {ex['subject']}\nMessage: {ex['message']}\nCategory: {ex['category']}\n"
        for ex in few_shot_examples
    )
    query = f"Subject: {email['subject']}\nMessage: {email['message']}\nCategory:"
    return f"{shots}\n{query}"
order_status_records = []
response_messages = []
results = []


for _, email in emails_df.iterrows():
    prompt = build_prompt(email)
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are an assistant classifying customer emails for a fashion store."},
            {"role": "user", "content": prompt}
        ],
        temperature=0
    )
    category = response.choices[0].message.content.strip().lower()
    results.append({"email_id": email["email_id"], "category": category})

classified_df = pd.DataFrame(results)
with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="w") as writer:
    classified_df.to_excel(writer, sheet_name="email-classification", index=False)

classified_df = pd.read_excel("output.xlsx", sheet_name="email-classification")
print("Columns in classified_df:", classified_df.columns.tolist())
order_requests = classified_df[classified_df["category"] == "order request"]
if not order_requests.empty:
    print("✅ Order requests found. Ready for processing...")
    print("📦 Total order requests detected:", len(order_requests))
    # TODO: Add logic to extract product IDs and quantities, check stock, and update order_status_sheet
else:
    print("⚠️ No order request data to update.")
responses = []


def extract_product_orders(text):
    orders = []
    # Match patterns like "2 CLF2109", "2x CLF2109", "CLF2109 x2", "CLF2109 - 2"
    patterns = [
        re.compile(r"(?P<quantity>\d+)\s*x?\s*[-:]?\s*(?P<product_id>[A-Z]{3,5}\d{3,})", re.IGNORECASE),
        re.compile(r"(?P<product_id>[A-Z]{3,5}\d{3,})\s*(?:x|-|:)?\s*(?P<quantity>\d+)", re.IGNORECASE),
    ]

    for pattern in patterns:
        for match in pattern.finditer(text):
            quantity = int(match.group("quantity"))
            product_id = match.group("product_id").upper()
            orders.append({"product_id": product_id, "quantity": quantity})

    return orders

for _, row in order_requests.iterrows():
    email_id = row["email_id"]
    message = emails_df.loc[emails_df["email_id"] == email_id, "message"].values[0]
    orders = extract_product_orders(message)

    if not orders:
        continue  # Skip if nothing extracted

    for order in orders:
        product_row = products_df[products_df["product_id"] == order["product_id"]]
        if not product_row.empty:
            available_qty = int(product_row["stock"].values[0])
            status = "Available" if order["quantity"] <= available_qty else "Out of Stock"
        else:
            status = "Product Not Found"

        order_status_records.append({
            "email ID": email_id,
            "product ID": order["product_id"],
            "quantity": order["quantity"],
            "status": status
        })

    # 💬 Generate reply
    if orders:
        response = f"Thank you for your order request.\n"
        for order in orders:
            prod = order["product_id"]
            qty = order["quantity"]
            product_row = products_df[products_df["product_id"] == prod]
            if not product_row.empty:
                available_qty = int(product_row["stock"].values[0])
                if qty <= available_qty:
                    response += f"- {qty} x {prod} is available and has been reserved.\n"
                else:
                    response += f"- {qty} x {prod} is currently out of stock. We have {available_qty} available.\n"
            else:
                response += f"- {qty} x {prod}: Product not found.\n"
        response_messages.append({"email_id": email_id, "response": response})

book = load_workbook("output.xlsx")

order_status_df = pd.DataFrame(order_status_records)
with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    order_status_df.to_excel(writer, sheet_name="order-status", index=False)

order_status_df = pd.read_excel("output.xlsx", sheet_name="order-status")
print("Columns in order_status_df:", order_status_df.columns.tolist())

order_response_df = pd.DataFrame(response_messages)
with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    order_response_df.to_excel(writer, sheet_name="order-response", index=False)
order_response_df = pd.read_excel("output.xlsx", sheet_name="order-response")
print("Columns in order_response_df:", order_response_df.columns.tolist())