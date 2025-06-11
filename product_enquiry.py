import pandas as pd
from openai import OpenAI
import numpy as np
from dotenv import load_dotenv
from sklearn.metrics.pairwise import cosine_similarity
from openai import OpenAI
from email_classifier import classified_df, products_df, emails_df

load_dotenv()
status = "Processed"

def get_embedding(text, model="text-embedding-ada-002"):
    response = client.embeddings.create(
        input=[text],
        model=model
    )
    return response.data[0].embedding


client = OpenAI(
    api_key='APIKEYHERE')

inquiry_emails = classified_df[classified_df["category"] == "product inquiry"]

# Step 1: Create vector embeddings for all product descriptions
product_descriptions = products_df["name"] + " - " + products_df["description"].fillna("")
product_embeddings = [get_embedding(desc, model="text-embedding-ada-002") for desc in product_descriptions]
products_df["embedding"] = product_embeddings

inquiry_response_rows = []

for _, row in inquiry_emails.iterrows():
    email_id = row["email_id"]
    email_body = emails_df.loc[emails_df["email_id"] == email_id, "message"].values[0]

    # Step 2: Embed the inquiry
    inquiry_embedding = get_embedding(email_body, model="text-embedding-ada-002")

    # Step 3: Calculate similarity with products
    sims = cosine_similarity([inquiry_embedding], list(products_df["embedding"]))[0]
    top_indices = sims.argsort()[-3:][::-1]  # top 3 products
    top_products = products_df.iloc[top_indices][["name", "description", "price"]]

    product_context = "\n".join(
        f"- {row['name']}: {row['description']} (Price: ${row['price']})"
        for _, row in top_products.iterrows()
    )

    # Step 4: Generate a response
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a helpful assistant responding to customer product inquiries."},
            {"role": "user",
             "content": f"Customer message:\n{email_body}\n\nRelevant products:\n{product_context}\n\nWrite a helpful and professional reply to the customer."}
        ],
        temperature=0.5
    )

    inquiry_response_rows.append({
        "email ID": email_id,
        "response": response.choices[0].message.content.strip()
    })

# Save to sheet
inquiry_response_df = pd.DataFrame(inquiry_response_rows)


with pd.ExcelWriter("output.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    inquiry_response_df.to_excel(writer, sheet_name="inquiry-response", index=False)

inquiry_response_df = pd.read_excel("output.xlsx", sheet_name="order-response")
print("Columns in order_response_df:", inquiry_response_df.columns.tolist())


