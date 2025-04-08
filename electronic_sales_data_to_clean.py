import pandas as pd

# Load data
file_path = "Kaggle Electronic Sales Data.csv"
df = pd.read_csv(file_path)

# Convert 'Purchase Date' to datetime and extract year/month
df['Purchase Date'] = pd.to_datetime(df['Purchase Date'], errors='coerce')
df['Year'] = df['Purchase Date'].dt.year
df['Month'] = df['Purchase Date'].dt.month

# Fill missing 'Gender' with 'Unknown'
df['Gender'] = df['Gender'].fillna('Unknown')

# Convert 'Loyalty Member' to boolean
df['Loyalty Member'] = df['Loyalty Member'].map({'Yes': True, 'No': False})

# Create age group column
bins = [0, 29, 50, 100]
labels = ['<30', '30-50', '>50']
df['Age Group'] = pd.cut(df['Age'], bins=bins, labels=labels, right=True)

# Count number of add-ons
df['Num Add-ons'] = df['Add-ons Purchased'].fillna('').apply(
    lambda x: len([item for item in x.split(',') if item.strip()])
)

# Save cleaned version
df.to_csv('electronic_sales_data_cleaned.csv', index=False)