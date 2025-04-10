import os
import argparse
import pandas as pd
from dotenv import load_dotenv
from google.cloud import bigquery
from google.cloud.exceptions import NotFound
import sys
import tempfile
import uuid

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
        # Initialize BigQuery client
        # Credentials will be automatically sourced from the
        # GOOGLE_APPLICATION_CREDENTIALS environment variable.
        client = bigquery.Client(project=project_id)
        table_ref = client.dataset(dataset_id).table(table_id)
        full_table_id = f"{project_id}.{dataset_id}.{table_id}"

        print(f"Attempting to load data into BigQuery table: {full_table_id}")
        print(f"Source CSV file: {csv_file_path}")

        # Define the schema to match the existing BigQuery table.
        # The pandas transformation handles the 'Feb-25' format before loading.
        print("Defining schema to match BigQuery table (datePeriod as DATE).")
        schema = [
            bigquery.SchemaField("purpleKey", "STRING"),
            bigquery.SchemaField("storeName", "STRING"),
            bigquery.SchemaField("retailerName", "STRING"),
            bigquery.SchemaField("storeTagging", "STRING"),
            bigquery.SchemaField("dateOpened", "STRING"), # Consider DATE/TIMESTAMP if format allows
            bigquery.SchemaField("Territory", "STRING"),
            bigquery.SchemaField("TSM", "STRING"),
            bigquery.SchemaField("TSS", "STRING"),
            bigquery.SchemaField("Region", "STRING"),
            bigquery.SchemaField("RSM", "STRING"),
            bigquery.SchemaField("tonikSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("hcSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("skyroSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("salmonSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("inHouseSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("creditCardSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("cashSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("otherSales", "STRING"), # Consider NUMERIC/FLOAT/INTEGER
            bigquery.SchemaField("retailerHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("tonikHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("hcHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("skyroHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("salmonHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("storeHeadcount", "STRING"), # Consider INTEGER
            bigquery.SchemaField("sourceFile", "STRING"),
            bigquery.SchemaField("datePeriod", "DATE"), # Matching existing table schema
        ]

        # Configure the load job
        job_config = bigquery.LoadJobConfig(
            schema=schema,
            skip_leading_rows=1,  # Assumes the first row is the header
            source_format=bigquery.SourceFormat.CSV,
            write_disposition=bigquery.WriteDisposition.WRITE_APPEND, # Append to existing table
            # autodetect=False, # Explicit schema is provided
        )

        # Read CSV using pandas
        try:
            # Specify dtype={'datePeriod': str} to ensure pandas reads it as text first
            df = pd.read_csv(csv_file_path, dtype={'datePeriod': str})
            print(f"Successfully read {len(df)} rows from {csv_file_path}")

            # --- Date Transformation ---
            print("Transforming 'datePeriod' column from 'Mon-YY' format to datetime objects...")
            # Convert 'Feb-25' style dates to datetime objects. Errors='coerce' will turn invalid formats into NaT (Not a Time)
            df['datePeriod'] = pd.to_datetime(df['datePeriod'], format='%b-%y', errors='coerce')

            # Check for any dates that failed to parse
            invalid_dates = df['datePeriod'].isna().sum()
            if invalid_dates > 0:
                print(f"Warning: {invalid_dates} rows had invalid date formats in 'datePeriod' and were set to null.")
                # Optionally, handle or log these rows further if needed

            # BigQuery client library usually handles datetime objects correctly when loading from DataFrame.
            # Force formatting to YYYY-MM-DD string to ensure compatibility with BQ DATE type.
            df['datePeriod'] = df['datePeriod'].dt.strftime('%Y-%m-%d')
            # -------------------------

        except FileNotFoundError:
            print(f"Error: CSV file not found at {csv_file_path}")
            sys.exit(1)
        except Exception as e:
            print(f"Error reading CSV file {csv_file_path}: {e}")
            sys.exit(1)

        # --- Load via Temporary File ---
        temp_file_path = None
        try:
            # Create a unique temporary file name
            temp_dir = tempfile.gettempdir()
            temp_file_name = f"bq_upload_{uuid.uuid4()}.csv"
            temp_file_path = os.path.join(temp_dir, temp_file_name)

            print(f"Saving transformed data to temporary file: {temp_file_path}")
            # Save DataFrame to CSV without index and header
            # BigQuery LoadJobConfig handles skipping the original header via skip_leading_rows=1
            # The datePeriod column now contains 'YYYY-MM-DD' strings from the transformation step.
            df.to_csv(temp_file_path, index=False, header=False) # Pandas handles date format correctly here

            # Load data from the temporary file using load_table_from_file
            print(f"Loading data from temporary file into BigQuery...")
            with open(temp_file_path, "rb") as source_file:
                # Pass the job_config which correctly defines datePeriod as DATE
                job = client.load_table_from_file(source_file, table_ref, job_config=job_config)

            print("Starting BigQuery load job...")
            job.result()  # Wait for the job to complete

        finally:
            # Clean up the temporary file
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                    print(f"Removed temporary file: {temp_file_path}")
                except OSError as e:
                    print(f"Error removing temporary file {temp_file_path}: {e}")
        # -----------------------------

        # Check job status
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
    # Load environment variables from .env file
    load_dotenv()

    # Retrieve configuration from environment variables
    project_id = os.getenv("GOOGLE_PROJECT_ID")
    dataset_id = os.getenv("GOOGLE_DATASET_ID")
    table_id = os.getenv("GOOGLE_TABLE_ID")
    credentials_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS") # Used implicitly by client library

    # Basic validation
    if not all([project_id, dataset_id, table_id, credentials_path]):
        print("Error: Missing required environment variables in .env file.")
        print("Please ensure GOOGLE_PROJECT_ID, GOOGLE_DATASET_ID, GOOGLE_TABLE_ID, and GOOGLE_APPLICATION_CREDENTIALS are set.")
        sys.exit(1)

    # Check if credentials file exists
    if not os.path.exists(credentials_path):
         print(f"Error: Credentials file not found at path specified in .env: {credentials_path}")
         sys.exit(1)

    # Set up argument parser
    parser = argparse.ArgumentParser(description="Upload a CSV file to a Google BigQuery table.")
    parser.add_argument("csv_file", help="Path to the CSV file to upload.")

    # Parse arguments
    args = parser.parse_args()

    # Run the upload function
    upload_csv_to_bigquery(args.csv_file, project_id, dataset_id, table_id)
