import sqlite3
import pandas as pd

# Create a SQLite connection and write cleaned data into a table
db_path = "C:/Users/busiz/OneDrive/Desktop/sales_data.db"
conn = sqlite3.connect(db_path)

# Save the cleaned dataframe into a SQLite table
df_cleaned = pd.read_csv("C:/Users/busiz/OneDrive/Desktop/electronic_sales_data_cleaned.csv")
df_cleaned.to_sql("fact_sales", conn, if_exists="replace", index=False)

# Define sample dimension tables with CREATE TABLE statements
dim_customer_sql = """
CREATE TABLE IF NOT EXISTS dim_customer AS
SELECT DISTINCT
    "Customer ID" AS customer_id,
    Age,
    "Gender",
    "Age Group",
    "Loyalty Member" AS loyalty_member
FROM fact_sales;
"""

dim_product_sql = """
CREATE TABLE IF NOT EXISTS dim_product AS
SELECT DISTINCT
    SKU AS sku,
    "Product Type" AS product_type,
    "Unit Price" AS unit_price
FROM fact_sales;
"""

dim_date_sql = """
CREATE TABLE IF NOT EXISTS dim_date AS
SELECT DISTINCT
    "Purchase Date" AS purchase_date,
    Year,
    Month
FROM fact_sales;
"""

# Execute the SQL to create dimension tables
cursor = conn.cursor()
cursor.execute(dim_customer_sql)
cursor.execute(dim_product_sql)
cursor.execute(dim_date_sql)
conn.commit()
