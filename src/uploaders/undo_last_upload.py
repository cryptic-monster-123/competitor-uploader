import os
from dotenv import load_dotenv
from google.cloud import bigquery
import sys

def undo_last_upload(project_id, dataset_id, table_id):
    """
    Deletes the last batch of uploaded rows from a BigQuery table.
    It identifies the last batch by the 'upload_timestamp' column.
    """
    try:
        client = bigquery.Client(project=project_id)
        full_table_id = f"{project_id}.{dataset_id}.{table_id}"

        # Find the most recent upload_timestamp
        query = f"""
            SELECT MAX(upload_timestamp) as last_upload
            FROM `{full_table_id}`
        """
        print("Finding the timestamp of the last upload...")
        query_job = client.query(query)
        results = query_job.result()

        last_upload_timestamp = None
        for row in results:
            last_upload_timestamp = row.last_upload

        if last_upload_timestamp is None:
            print("No previous uploads found (upload_timestamp column is empty).")
            sys.exit(0)

        print(f"Last upload was at: {last_upload_timestamp}")

        # Delete all rows with that timestamp
        delete_query = f"""
            DELETE FROM `{full_table_id}`
            WHERE upload_timestamp = TIMESTAMP('{last_upload_timestamp.isoformat()}')
        """
        print("Deleting rows from the last upload...")
        delete_job = client.query(delete_query)
        delete_job.result()  # Wait for the job to complete

        print(f"Successfully deleted {delete_job.num_dml_affected_rows} rows from the last upload.")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    load_dotenv()

    project_id = os.getenv("GOOGLE_PROJECT_ID")
    dataset_id = os.getenv("GOOGLE_DATASET_ID")
    table_id = os.getenv("GOOGLE_TABLE_ID")

    if not all([project_id, dataset_id, table_id]):
        print("Error: Missing required environment variables in .env file.")
        print("Please ensure GOOGLE_PROJECT_ID, GOOGLE_DATASET_ID, and GOOGLE_TABLE_ID are set.")
        sys.exit(1)

    undo_last_upload(project_id, dataset_id, table_id)
