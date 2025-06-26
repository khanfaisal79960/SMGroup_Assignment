import pandas as pd
from datetime import datetime, timedelta
import os
import dotenv
from openai import OpenAI

# Load environment variables from .env file for local development
dotenv.load_dotenv()

# Configure OpenAI client to use OpenRouter API (still required if you use it elsewhere or just for structure)
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=os.getenv("OPENROUTER_API_KEY")
)

file_path = "Cleaned_Marketing_Performance_Data.xlsx"
print(f"--- Starting Report Generation ---")
print(f"Attempting to load data from: {file_path}")

try:
    df = pd.read_excel(file_path)
    print(f"Successfully loaded '{file_path}'. Total rows: {len(df)}")
    if df.empty:
        print("Error: The loaded DataFrame is empty. Please check the Excel file content.")
        exit()
except FileNotFoundError:
    print(f"Error: The file '{file_path}' was not found. Please ensure it's in the correct directory.")
    exit()
except Exception as e:
    print(f"Error reading Excel file '{file_path}': {e}")
    exit()

df['Date'] = pd.to_datetime(df['Date'])
# Ensure data is sorted by date (still good practice even if not splitting for comparison)
df = df.sort_values(by='Date').reset_index(drop=True)

print(f"Original Date range in data: {df['Date'].min().strftime('%Y-%m-%d')} to {df['Date'].max().strftime('%Y-%m-%d')}")

# --- Analysis over the entire dataset ---

# 1. Calculate Top Campaigns based on the entire dataset
print("\n--- Generating Top Campaigns Report ---")
top_campaigns = (
    df.groupby('Campaign')[['Spend', 'Conversions']]
    .sum().sort_values(by='Conversions', ascending=False).head(5)
    .reset_index() # Converts the 'Campaign' index back to a column
)
print(f"Top campaigns DataFrame size (based on ALL data): {len(top_campaigns)} rows")
if top_campaigns.empty:
    print("Warning: No top campaigns found for the entire dataset. 'Top Campaigns' sheet will be empty or have no data rows.")


# --- Anomaly detection and AI suggestion feature REMOVED ---
# All code blocks related to previous_period_data, current_period_data,
# current_summary, previous_summary, comparison, anomalies, and get_ai_suggestion
# have been removed as per your request.


# 7. Generate the Excel report
print("\n--- Generating Excel Report ---")
output_file_name = "Weekly_Report_With_AI_Suggestions.xlsx"
try:
    with pd.ExcelWriter(output_file_name) as writer:
        top_campaigns.to_excel(writer, sheet_name="Top Campaigns", index=False)
    print(f"'{output_file_name}' has been generated successfully!")
    print(f"File saved to: {os.path.abspath(output_file_name)}")
except Exception as e:
    print(f"Error saving Excel report: {e}")

print("\n--- Report Generation Complete ---")