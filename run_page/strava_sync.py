import argparse
import json
import sys

from config import JSON_FILE, SQL_FILE, OUTPUT_DIR
from generator import Generator
from csv_exporter import CSVExporter


# for only run type, we use the same logic as garmin_sync
def run_strava_sync(client_id, client_secret, refresh_token, only_run=False):
    generator = Generator(SQL_FILE)
    generator.set_strava_config(client_id, client_secret, refresh_token)
    # if you want to refresh data change False to True
    generator.only_run = only_run
    
    # 执行同步
    generator.sync(False)
    
    # 导出新活动为CSV
    if generator.new_activity_ids:
        print(f"Found {len(generator.new_activity_ids)} new activities, exporting CSV...")
        try:
            csv_exporter = CSVExporter(OUTPUT_DIR)
            csv_exporter.export_activities_to_csv(
                generator.new_activity_ids,
                'strava',
                client=generator.client
            )
        except Exception as e:
            # CSV导出失败不应该中断同步流程
            print(f"Warning: Failed to export CSV files: {e}", file=sys.stderr)

    activities_list = generator.load()
    with open(JSON_FILE, "w") as f:
        json.dump(activities_list, f)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("client_id", help="strava client id")
    parser.add_argument("client_secret", help="strava client secret")
    parser.add_argument("refresh_token", help="strava refresh token")
    parser.add_argument(
        "--only-run",
        dest="only_run",
        action="store_true",
        help="if is only for running",
    )
    options = parser.parse_args()
    run_strava_sync(
        options.client_id,
        options.client_secret,
        options.refresh_token,
        only_run=options.only_run,
    )
