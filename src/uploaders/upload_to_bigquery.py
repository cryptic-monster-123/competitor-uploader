import os
import argparse
import pandas as pd
from dotenv import load_dotenv
from google.cloud import bigquery
from google.cloud.exceptions import NotFound
import sys
import tempfile
import uuid
from datetime import datetime
import pytz

def upload_csv_to_bigquery(csv_file_path, project_id, dataset_id, table_id):
    """
    Uploads data from a CSV file to a specified BigQuery table.

    Args:
        csv_file_path (str): The path to the input CSV file.
        project_id (str): Google Cloud project ID.
        dataset_id (str): BigQuery dataset ID.
        table_id (str): BigQuery table ID.
    """
    try:
        client = bigquery.Client(project=project_id)
        table_ref = client.dataset(dataset_id).table(table_id)
        full_table_id = f"{project_id}.{dataset_id}.{table_id}"

        print(f"Attempting to load data into BigQuery table: {full_table_id}")
        print(f"Source CSV file: {csv_file_path}")

        print("Defining schema for BigQuery table (all columns as STRING, datePeriod as DATE).")
        schema = [
            bigquery.SchemaField("purpleKey", "STRING"),
            bigquery.SchemaField("storeName", "STRING"),
            bigquery.SchemaField("retailerName", "STRING"),
            bigquery.SchemaField("storeTagging", "STRING"),
            bigquery.SchemaField("dateOpened", "STRING"),
            bigquery.SchemaField("Territory", "STRING"),
            bigquery.SchemaField("TSM", "STRING"),
            bigquery.SchemaField("TSS", "STRING"),
            bigquery.SchemaField("Region", "STRING"),
            bigquery.SchemaField("RSM", "STRING"),
            bigquery.SchemaField("tonikSales", "STRING"),
            bigquery.SchemaField("hcSales", "STRING"),
            bigquery.SchemaField("skyroSales", "STRING"),
            bigquery.SchemaField("salmonSales", "STRING"),
            bigquery.SchemaField("billeaseSales", "STRING"),
            bigquery.SchemaField("inHouseSales", "STRING"),
            bigquery.SchemaField("creditCardSales", "STRING"),
            bigquery.SchemaField("cashSales", "STRING"),
            bigquery.SchemaField("otherSales", "STRING"),
            bigquery.SchemaField("retailerHeadcount", "STRING"),
            bigquery.SchemaField("tonikHeadcount", "STRING"),
            bigquery.SchemaField("hcHeadcount", "STRING"),
            bigquery.SchemaField("skyroHeadcount", "STRING"),
            bigquery.SchemaField("salmonHeadcount", "STRING"),
            bigquery.SchemaField("billeaseHeadcount", "STRING"),
            bigquery.SchemaField("storeHeadcount", "STRING"),
            bigquery.SchemaField("sourceFile", "STRING"),
            bigquery.SchemaField("datePeriod", "DATE"),
            bigquery.SchemaField("upload_timestamp", "TIMESTAMP"),
        ]

        job_config = bigquery.LoadJobConfig(
            schema=schema,
            skip_leading_rows=1,
            source_format=bigquery.SourceFormat.CSV,
            write_disposition=bigquery.WriteDisposition.WRITE_APPEND,
        )

        try:
            df = pd.read_csv(csv_file_path, dtype=str)
            print(f"Successfully read {len(df)} rows from {csv_file_path}")

            # Stamp rows with current Manila time
            manila_tz = pytz.timezone('Asia/Manila')
            upload_timestamp = datetime.now(manila_tz)
            df['upload_timestamp'] = upload_timestamp
            print(f"Added upload timestamp: {upload_timestamp}")

            # --- Date Transformation ---
            print("Processing 'datePeriod' column...")

            if 'datePeriod' in df.columns:
                sample_dates = df['datePeriod'].dropna().head(5)
                already_formatted = all(
                    isinstance(date_val, str) and
                    len(str(date_val)) == 10 and
                    str(date_val).count('-') == 2 and
                    str(date_val)[:4].isdigit()
                    for date_val in sample_dates
                )

                if already_formatted:
                    print("datePeriod column is already in YYYY-MM-DD format, no transformation needed.")
                else:
                    print("Transforming 'datePeriod' column from various formats to YYYY-MM-DD...")

                    def parse_date_flexible(date_str):
                        if pd.isna(date_str) or not date_str:
                            return None

                        date_str = str(date_str).strip()

                        formats_to_try = [
                            '%b-%y',    # May-25
                            '%B %Y',    # May 2025
                            '%b %Y',    # May 2025
                            '%Y-%m',    # 2025-05
                            '%Y-%m-%d', # 2025-05-01
                        ]

                        for fmt in formats_to_try:
                            try:
                                parsed_date = pd.to_datetime(date_str, format=fmt)
                                return parsed_date.strftime('%Y-%m-01')
                            except (ValueError, TypeError):
                                continue

                        try:
                            parsed_date = pd.to_datetime(date_str, errors='coerce')
                            if not pd.isna(parsed_date):
                                return parsed_date.strftime('%Y-%m-01')
                        except Exception:
                            pass

                        print(f"Warning: Could not parse date '{date_str}', setting to null")
                        return None

                    df['datePeriod'] = df['datePeriod'].apply(parse_date_flexible)

                    invalid_dates = df['datePeriod'].isna().sum()
                    if invalid_dates > 0:
                        print(f"Warning: {invalid_dates} rows had invalid date formats in 'datePeriod' and were set to null.")
            else:
                print("Warning: 'datePeriod' column not found in the CSV file.")
            # -------------------------

            if 'upload_timestamp' in df.columns:
                df['upload_timestamp'] = (
                    pd.to_datetime(df['upload_timestamp'], errors='coerce', utc=True)
                    .dt.strftime('%Y-%m-%d %H:%M:%S')
                )

        except FileNotFoundError:
            print(f"Error: CSV file not found at {csv_file_path}")
            sys.exit(1)
        except Exception as e:
            print(f"Error reading CSV file {csv_file_path}: {e}")
            sys.exit(1)

        try:
            print("Loading data directly from DataFrame into BigQuery...")
            job = client.load_table_from_dataframe(df, table_ref, job_config=job_config)
            print("Starting BigQuery load job...")
            job.result()
        except Exception as e:
            print(f"Error loading data from DataFrame to BigQuery: {e}")
            sys.exit(1)

        if job.errors:
            print("BigQuery load job failed:")
            for error in job.errors:
                print(f"- {error['message']}")
            sys.exit(1)
        else:
            table = client.get_table(table_ref)
            print(f"Load job completed successfully. {job.output_rows} rows loaded.")
            print(f"Total rows in table {full_table_id}: {table.num_rows}")

    except NotFound:
        print(f"Error: BigQuery table {full_table_id} not found.")
        print("Please ensure the project, dataset, and table exist and the service account has permissions.")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    load_dotenv()

    project_id = os.getenv("GOOGLE_PROJECT_ID")
    dataset_id = os.getenv("GOOGLE_DATASET_ID")
    table_id = os.getenv("GOOGLE_TABLE_ID")
    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")

    if not all([project_id, dataset_id, table_id, credentials_path]):
        print("Error: Missing required environment variables in .env file.")
        print("Please ensure GOOGLE_PROJECT_ID, GOOGLE_DATASET_ID, GOOGLE_TABLE_ID, and GOOGLE_APPLICATION_CREDENTIALS are set.")
        sys.exit(1)

    if not os.path.exists(credentials_path):
        print(f"Error: Credentials file not found at path specified in .env: {credentials_path}")
        sys.exit(1)

    parser = argparse.ArgumentParser(description="Upload a CSV file to a Google BigQuery table.")
    parser.add_argument("csv_file", help="Path to the CSV file to upload.")

    args = parser.parse_args()

    csv_path = args.csv_file
    if not os.path.isfile(csv_path):
        print(f"Error: CSV file not found: {csv_path}")
        print("Make sure you pass the generated combined CSV, e.g.:")
        print("  output/csv/combined_data_<timestamp>.csv")
        sys.exit(1)

    if not csv_path.lower().endswith(".csv"):
        print(f"Error: Expected a .csv file, got: {csv_path}")
        print("Make sure you pass the combined CSV from `output/csv/`, not this script path.")
        sys.exit(1)

    upload_csv_to_bigquery(csv_path, project_id, dataset_id, table_id)
