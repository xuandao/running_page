"""
If you do not want bind any account
Only the tcx files in TCX_OUT sync
"""

import sys
from config import JSON_FILE, SQL_FILE, TCX_FOLDER, OUTPUT_DIR
from generator import Generator
from csv_exporter import CSVExporter
import json

if __name__ == "__main__":
    print("only sync tcx files in TCX_OUT")
    
    # 使用Generator进行同步
    generator = Generator(SQL_FILE)
    generator.sync_from_data_dir(TCX_FOLDER, file_suffix="tcx")
    
    # 导出新活动为CSV
    if generator.new_activity_ids:
        print(f"Found {len(generator.new_activity_ids)} new activities, exporting CSV...")
        try:
            csv_exporter = CSVExporter(OUTPUT_DIR)
            csv_exporter.export_activities_from_files(
                generator.new_activity_ids,
                'tcx',
                data_dir=TCX_FOLDER,
                file_suffix="tcx"
            )
        except Exception as e:
            # CSV导出失败不应该中断同步流程
            print(f"Warning: Failed to export CSV files: {e}", file=sys.stderr)
    
    # 生成activities.json
    activities_list = generator.load()
    with open(JSON_FILE, "w") as f:
        json.dump(activities_list, f)
