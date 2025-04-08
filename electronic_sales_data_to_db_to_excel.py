import sqlite3
import pandas as pd
from tabulate import tabulate

# Define all SQL queries
queries = {
    "Monthly_Sales_Trend": """
        SELECT Year, Month, ROUND(SUM("Total Price"), 2) AS total_sales
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY Year, Month
        ORDER BY Year, Month;
    """,
    "Top_Selling_Products": """
        SELECT SKU, "Product Type", SUM(Quantity) AS total_units_sold
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY SKU, "Product Type"
        ORDER BY total_units_sold DESC
        LIMIT 10;
    """,
    "Sales_by_Age_Group": """
        SELECT "Age Group", ROUND(SUM("Total Price"), 2) AS total_sales
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY "Age Group"
        ORDER BY "Age Group";
    """,
    "Sales_by_Payment_Method": """
        SELECT "Payment Method", ROUND(SUM("Total Price"), 2) AS total_sales
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY "Payment Method"
        ORDER BY total_sales DESC;
    """,
    "Addons_Impact": """
        SELECT
            CASE 
              WHEN "Num Add-ons" = 0 THEN 'No Add-ons'
              ELSE 'With Add-ons'
            END AS addon_group,
            COUNT(*) AS orders,
            ROUND(SUM("Total Price"), 2) AS total_sales
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY addon_group;
    """,
    "Average_Order_Value_by_Loyalty": """
        SELECT
            Year,
            Month,
            "Loyalty Member" AS loyalty_member,
            ROUND(AVG("Total Price"), 2) AS avg_order_value,
            COUNT(*) AS total_orders
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY Year, Month, "Loyalty Member";
    """,
    "Top_3_Products_per_Year": """
        SELECT *
        FROM (
            SELECT
                Year,
                "Product Type",
                ROUND(SUM("Total Price"), 2) AS total_sales,
                RANK() OVER (PARTITION BY Year ORDER BY SUM("Total Price") DESC) AS rank
            FROM fact_sales
            WHERE "Order Status" = 'Completed'
            GROUP BY Year, "Product Type"
        ) AS ranked
        WHERE rank <= 3;
    """,
    "Addon_Ratio_by_Shipping": """
        SELECT
            "Shipping Type",
            SUM(CASE WHEN "Num Add-ons" > 0 THEN 1 ELSE 0 END) AS with_addon,
            SUM(CASE WHEN "Num Add-ons" = 0 THEN 1 ELSE 0 END) AS without_addon,
            ROUND(1.0 * SUM(CASE WHEN "Num Add-ons" > 0 THEN 1 ELSE 0 END) / COUNT(*), 2) AS addon_ratio
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY "Shipping Type";
    """,
    "Monthly_Addon_Revenue": """
        SELECT
            Year,
            Month,
            ROUND(SUM("Add-on Total"), 6) AS total_addon_revenue,
            ROUND(SUM("Add-on Total") / SUM("Total Price"), 6) AS addon_share_of_sales
        FROM fact_sales
        WHERE "Order Status" = 'Completed'
        GROUP BY Year, Month
        ORDER BY Year, Month;
    """
}

# Connect to DB
conn = sqlite3.connect("sales_data.db")

# Create Excel writer
with pd.ExcelWriter("sales_query_results.xlsx", engine="openpyxl") as writer:
    for name, query in queries.items():
        df = pd.read_sql_query(query, conn)
        print(f"\n=== {name.replace('_', ' ')} ===\n")
        print(tabulate(df, headers='keys', tablefmt='fancy_grid', showindex=False))
        df.to_excel(writer, sheet_name=name[:31], index=False)

conn.close()
print("\n Excel file saved as 'sales_query_results.xlsx'")
