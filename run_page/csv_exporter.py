"""
CSV导出模块
用于将跑步活动的详细数据导出为CSV格式
"""
import csv
import os
import sys
from datetime import datetime, timedelta
from io import StringIO
from math import floor
from typing import List, Dict, Optional, Any
import xml.etree.ElementTree as ET


class CSVExporter:
    """CSV导出器"""
    
    def __init__(self, activities_dir):
        """
        初始化CSV导出器
        
        Args:
            activities_dir: activities目录路径
        """
        self.activities_dir = activities_dir
        # 确保目录存在
        os.makedirs(activities_dir, exist_ok=True)
    
    def format_time(self, seconds: float) -> str:
        """
        格式化时间为 分:秒.毫秒 格式
        
        Args:
            seconds: 秒数
            
        Returns:
            格式化的时间字符串，如 "4:29.6"
        """
        if not seconds or seconds <= 0:
            return "0:00.0"
        
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        millis = int((seconds % 1) * 10)
        
        return f"{minutes}:{secs:02d}.{millis}"
    
    def format_pace(self, meters_per_second: float) -> str:
        """
        格式化配速为 分:秒/千米 格式
        
        Args:
            meters_per_second: 米每秒的速度
            
        Returns:
            格式化的配速字符串，如 "4:30"
        """
        if not meters_per_second or meters_per_second <= 0:
            return "0:00"
        
        # 计算每千米需要多少分钟
        seconds_per_km = 1000.0 / meters_per_second
        minutes = int(seconds_per_km // 60)
        seconds = int(seconds_per_km % 60)
        
        return f"{minutes}:{seconds:02d}"
    
    def calculate_pace_from_distance_time(self, distance_meters: float, time_seconds: float) -> str:
        """
        根据距离和时间计算配速
        
        Args:
            distance_meters: 距离（米）
            time_seconds: 时间（秒）
            
        Returns:
            配速字符串
        """
        if not distance_meters or distance_meters <= 0 or not time_seconds or time_seconds <= 0:
            return "0:00"
        
        meters_per_second = distance_meters / time_seconds
        return self.format_pace(meters_per_second)
    
    def generate_csv(self, laps_data: List[Dict], summary_data: Dict) -> str:
        """
        生成CSV内容
        
        Args:
            laps_data: 计圈数据列表
            summary_data: 汇总数据
            
        Returns:
            CSV内容字符串
        """
        output = StringIO()
        writer = csv.writer(output, quoting=csv.QUOTE_ALL)
        
        # 写入表头
        headers = [
            "计圈", "时间", "累积时间", "距离公里", "平均配速分钟/千米",
            "平均坡度调整配速分钟/千米", "平均心率bpm", "最大心率bpm",
            "累计爬升米", "累计下降米", "平均功率瓦", "平均 W/kg",
            "最大功率瓦", "最大 W/kg", "平均步频每分钟步数", "平均触地时间毫秒",
            "平均 GCT 平衡%", "平均步长米", "平均垂直摆动cm", "平均垂直步幅比%",
            "热量消耗卡路里", "平均温度", "最佳配速分钟/千米", "最高步频每分钟步数",
            "移动时间", "平均移动配速分钟/千米", "平均步速损失厘米/秒", "平均步速损失百分比%"
        ]
        writer.writerow(headers)
        
        # 写入计圈数据
        cumulative_time = 0.0
        cumulative_distance = 0.0
        
        for i, lap in enumerate(laps_data, 1):
            lap_time = lap.get('time_seconds', 0)
            lap_distance = lap.get('distance_meters', 0) / 1000.0  # 转换为公里
            
            cumulative_time += lap_time
            cumulative_distance += lap_distance
            
            # 计算配速
            pace = self.calculate_pace_from_distance_time(
                lap.get('distance_meters', 0),
                lap_time
            )
            
            row = [
                str(i),  # 计圈
                self.format_time(lap_time),  # 时间
                self.format_time(cumulative_time),  # 累积时间
                f"{lap_distance:.2f}",  # 距离公里
                pace,  # 平均配速
                lap.get('grade_adjusted_pace', pace),  # 坡度调整配速
                str(lap.get('avg_heartrate', '') or ''),  # 平均心率
                str(lap.get('max_heartrate', '') or ''),  # 最大心率
                str(lap.get('elevation_gain', '') or ''),  # 累计爬升
                str(lap.get('elevation_loss', '') or ''),  # 累计下降
                str(lap.get('avg_power', '') or ''),  # 平均功率
                str(lap.get('avg_power_per_kg', '') or ''),  # 平均功率体重比
                str(lap.get('max_power', '') or ''),  # 最大功率
                str(lap.get('max_power_per_kg', '') or ''),  # 最大功率体重比
                str(lap.get('avg_cadence', '') or ''),  # 平均步频
                str(lap.get('avg_ground_contact_time', '') or ''),  # 平均触地时间
                lap.get('gct_balance', '') or '',  # GCT平衡
                str(lap.get('avg_stride_length', '') or ''),  # 平均步长
                str(lap.get('avg_vertical_oscillation', '') or ''),  # 平均垂直摆动
                str(lap.get('avg_vertical_ratio', '') or ''),  # 平均垂直步幅比
                str(lap.get('calories', '') or ''),  # 热量消耗
                str(lap.get('avg_temperature', '') or ''),  # 平均温度
                lap.get('best_pace', pace),  # 最佳配速
                str(lap.get('max_cadence', '') or ''),  # 最高步频
                self.format_time(lap.get('moving_time', lap_time)),  # 移动时间
                lap.get('moving_pace', pace),  # 移动配速
                str(lap.get('pace_loss', '') or ''),  # 步速损失
                str(lap.get('pace_loss_percent', '') or '')  # 步速损失百分比
            ]
            writer.writerow(row)
        
        # 写入统计行
        total_time = summary_data.get('total_time', cumulative_time)
        total_distance = summary_data.get('total_distance', cumulative_distance)
        avg_pace = self.calculate_pace_from_distance_time(
            total_distance * 1000,
            total_time
        ) if total_distance > 0 and total_time > 0 else "0:00"
        
        stats_row = [
            "统计",
            self.format_time(total_time),
            self.format_time(total_time),
            f"{total_distance:.2f}",
            avg_pace,
            summary_data.get('grade_adjusted_pace', avg_pace),
            str(summary_data.get('avg_heartrate', '') or ''),
            str(summary_data.get('max_heartrate', '') or ''),
            str(summary_data.get('total_elevation_gain', '') or ''),
            str(summary_data.get('total_elevation_loss', '') or ''),
            str(summary_data.get('avg_power', '') or ''),
            str(summary_data.get('avg_power_per_kg', '') or ''),
            str(summary_data.get('max_power', '') or ''),
            str(summary_data.get('max_power_per_kg', '') or ''),
            str(summary_data.get('avg_cadence', '') or ''),
            str(summary_data.get('avg_ground_contact_time', '') or ''),
            summary_data.get('gct_balance', '') or '',
            str(summary_data.get('avg_stride_length', '') or ''),
            str(summary_data.get('avg_vertical_oscillation', '') or ''),
            str(summary_data.get('avg_vertical_ratio', '') or ''),
            str(summary_data.get('total_calories', '') or ''),
            str(summary_data.get('avg_temperature', '') or ''),
            summary_data.get('best_pace', avg_pace),
            str(summary_data.get('max_cadence', '') or ''),
            self.format_time(summary_data.get('moving_time', total_time)),
            summary_data.get('moving_pace', avg_pace),
            '--',
            '--'
        ]
        writer.writerow(stats_row)
        
        return output.getvalue()
    
    def save_csv(self, activity_id: int, csv_content: str):
        """
        保存CSV文件
        
        Args:
            activity_id: 活动ID
            csv_content: CSV内容
        """
        filename = f"activity_{activity_id}.csv"
        filepath = os.path.join(self.activities_dir, filename)
        
        try:
            with open(filepath, 'w', encoding='utf-8-sig') as f:  # 使用utf-8-sig以支持Excel打开
                f.write(csv_content)
            print(f"Exported CSV: {filepath}")
        except Exception as e:
            print(f"Error saving CSV file {filepath}: {e}", file=sys.stderr)
            raise
    
    def parse_tcx_data(self, tcx_data: str) -> tuple[List[Dict], Dict]:
        """
        从TCX数据解析计圈数据
        
        Args:
            tcx_data: TCX XML数据字符串
            
        Returns:
            (laps_data, summary_data) 元组
        """
        try:
            root = ET.fromstring(tcx_data)
            
            # TCX命名空间
            ns = {
                'tcx': 'http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2'
            }
            
            laps_data = []
            total_distance = 0.0
            total_time = 0.0
            total_calories = 0
            all_heartrates = []
            all_elevations = []
            
            # 查找所有Lap元素
            for lap in root.findall('.//tcx:Lap', ns):
                lap_data = {}
                
                # 距离（米）
                distance_elem = lap.find('tcx:DistanceMeters', ns)
                if distance_elem is not None:
                    lap_distance = float(distance_elem.text)
                    lap_data['distance_meters'] = lap_distance
                    total_distance += lap_distance
                else:
                    lap_data['distance_meters'] = 0.0
                
                # 时间（秒）
                time_elem = lap.find('tcx:TotalTimeSeconds', ns)
                if time_elem is not None:
                    lap_time = float(time_elem.text)
                    lap_data['time_seconds'] = lap_time
                    total_time += lap_time
                else:
                    lap_data['time_seconds'] = 0.0
                
                # 卡路里
                calories_elem = lap.find('tcx:Calories', ns)
                if calories_elem is not None:
                    calories = int(calories_elem.text)
                    lap_data['calories'] = calories
                    total_calories += calories
                
                # 平均心率
                avg_hr_elem = lap.find('.//tcx:AverageHeartRateBpm/tcx:Value', ns)
                if avg_hr_elem is not None:
                    avg_hr = int(avg_hr_elem.text)
                    lap_data['avg_heartrate'] = avg_hr
                    all_heartrates.append(avg_hr)
                
                # 最大心率
                max_hr_elem = lap.find('.//tcx:MaximumHeartRateBpm/tcx:Value', ns)
                if max_hr_elem is not None:
                    lap_data['max_heartrate'] = int(max_hr_elem.text)
                
                # 平均配速（从距离和时间计算）
                if lap_data['distance_meters'] > 0 and lap_data['time_seconds'] > 0:
                    lap_data['pace'] = self.calculate_pace_from_distance_time(
                        lap_data['distance_meters'],
                        lap_data['time_seconds']
                    )
                
                # 爬升和下降（从TrackPoint计算）
                track_points = lap.findall('.//tcx:TrackPoint', ns)
                elevations = []
                for tp in track_points:
                    alt_elem = tp.find('tcx:AltitudeMeters', ns)
                    if alt_elem is not None:
                        elevations.append(float(alt_elem.text))
                
                if elevations:
                    elevation_gain = 0.0
                    elevation_loss = 0.0
                    for i in range(1, len(elevations)):
                        diff = elevations[i] - elevations[i-1]
                        if diff > 0:
                            elevation_gain += diff
                        else:
                            elevation_loss += abs(diff)
                    
                    lap_data['elevation_gain'] = int(elevation_gain)
                    lap_data['elevation_loss'] = int(elevation_loss)
                    all_elevations.extend(elevations)
                
                laps_data.append(lap_data)
            
            # 生成汇总数据
            summary_data = {
                'total_time': total_time,
                'total_distance': total_distance / 1000.0,  # 转换为公里
                'total_calories': total_calories,
                'avg_heartrate': int(sum(all_heartrates) / len(all_heartrates)) if all_heartrates else None,
                'max_heartrate': max(all_heartrates) if all_heartrates else None,
                'total_elevation_gain': int(sum(lap.get('elevation_gain', 0) for lap in laps_data)),
                'total_elevation_loss': int(sum(lap.get('elevation_loss', 0) for lap in laps_data)),
            }
            
            return laps_data, summary_data
            
        except Exception as e:
            print(f"Error parsing TCX data: {e}", file=sys.stderr)
            raise
    
    def export_activity_from_tcx(self, activity_id: int, tcx_file_path: str):
        """
        从TCX文件导出活动为CSV
        
        Args:
            activity_id: 活动ID
            tcx_file_path: TCX文件路径
        """
        try:
            with open(tcx_file_path, 'r', encoding='utf-8') as f:
                tcx_data = f.read()
            
            laps_data, summary_data = self.parse_tcx_data(tcx_data)
            
            if not laps_data:
                print(f"Warning: No lap data found in TCX file {tcx_file_path}")
                return
            
            csv_content = self.generate_csv(laps_data, summary_data)
            self.save_csv(activity_id, csv_content)
            
        except Exception as e:
            print(f"Error exporting activity {activity_id} from TCX: {e}", file=sys.stderr)
            raise
    
    def export_activities_from_files(
        self,
        activity_ids: List[int],
        platform: str,
        data_dir: str,
        file_suffix: str = "tcx"
    ):
        """
        从文件目录导出多个活动为CSV
        
        Args:
            activity_ids: 活动ID列表
            platform: 平台名称
            data_dir: 数据文件目录
            file_suffix: 文件后缀（tcx, gpx, fit等）
        """
        if not activity_ids:
            return
        
        print(f"Exporting {len(activity_ids)} activities to CSV from {platform}...")
        
        success_count = 0
        fail_count = 0
        
        for activity_id in activity_ids:
            try:
                # 查找对应的文件
                filename = f"{activity_id}.{file_suffix}"
                filepath = os.path.join(data_dir, filename)
                
                if not os.path.exists(filepath):
                    print(f"Warning: File not found: {filepath}")
                    fail_count += 1
                    continue
                
                if file_suffix == "tcx":
                    self.export_activity_from_tcx(activity_id, filepath)
                    success_count += 1
                else:
                    print(f"Warning: Unsupported file format: {file_suffix}")
                    fail_count += 1
                    
            except Exception as e:
                print(f"Error exporting activity {activity_id}: {e}", file=sys.stderr)
                fail_count += 1
        
        print(f"CSV export completed: {success_count} succeeded, {fail_count} failed")
    
    def parse_strava_activity(self, activity_id: int, client) -> tuple[List[Dict], Dict]:
        """
        从Strava API获取并解析活动数据
        
        Args:
            activity_id: 活动ID
            client: Strava客户端对象
            
        Returns:
            (laps_data, summary_data) 元组
        """
        try:
            # 获取活动详情
            activity = client.get_activity(activity_id)
            
            # 尝试获取计圈数据
            laps_data = []
            try:
                laps = client.get_activity_laps(activity_id)
                for lap in laps:
                    lap_data = {
                        'distance_meters': float(lap.distance) if lap.distance else 0.0,
                        'time_seconds': float(lap.elapsed_time.total_seconds()) if lap.elapsed_time else 0.0,
                        'avg_heartrate': int(lap.average_heartrate) if lap.average_heartrate else None,
                        'max_heartrate': int(lap.max_heartrate) if lap.max_heartrate else None,
                        'elevation_gain': float(lap.total_elevation_gain) if lap.total_elevation_gain else 0.0,
                        'elevation_loss': 0.0,  # Strava API不直接提供下降数据
                        'calories': int(lap.calories) if lap.calories else None,
                    }
                    laps_data.append(lap_data)
            except Exception as e:
                # 如果没有计圈数据，尝试从流式数据计算
                print(f"Warning: No lap data for activity {activity_id}, trying streams: {e}")
                try:
                    streams = client.get_activity_streams(
                        activity_id,
                        types=['distance', 'time', 'heartrate', 'altitude', 'cadence', 'watts']
                    )
                    # 从流式数据计算计圈（每1公里一圈）
                    laps_data = self.calculate_laps_from_streams(streams, activity)
                except Exception as e2:
                    print(f"Warning: Failed to get streams for activity {activity_id}: {e2}")
                    # 如果都没有，至少创建一个总的活动数据
                    if activity.distance and activity.elapsed_time:
                        laps_data = [{
                            'distance_meters': float(activity.distance),
                            'time_seconds': float(activity.elapsed_time.total_seconds()),
                            'avg_heartrate': int(activity.average_heartrate) if activity.average_heartrate else None,
                            'max_heartrate': int(activity.max_heartrate) if activity.max_heartrate else None,
                            'elevation_gain': float(activity.total_elevation_gain) if activity.total_elevation_gain else 0.0,
                            'elevation_loss': 0.0,
                            'calories': int(activity.calories) if activity.calories else None,
                        }]
            
            # 生成汇总数据
            summary_data = {
                'total_time': float(activity.elapsed_time.total_seconds()) if activity.elapsed_time else 0.0,
                'total_distance': float(activity.distance) / 1000.0 if activity.distance else 0.0,
                'total_calories': int(activity.calories) if activity.calories else None,
                'avg_heartrate': int(activity.average_heartrate) if activity.average_heartrate else None,
                'max_heartrate': int(activity.max_heartrate) if activity.max_heartrate else None,
                'total_elevation_gain': float(activity.total_elevation_gain) if activity.total_elevation_gain else 0.0,
                'total_elevation_loss': 0.0,
            }
            
            return laps_data, summary_data
            
        except Exception as e:
            print(f"Error parsing Strava activity {activity_id}: {e}", file=sys.stderr)
            raise
    
    def calculate_laps_from_streams(self, streams: Dict, activity, lap_distance: int = 1000) -> List[Dict]:
        """
        从流式数据计算计圈数据（每1公里一圈）
        
        Args:
            streams: Strava流式数据字典
            activity: Strava活动对象
            lap_distance: 每圈距离（米），默认1000米
            
        Returns:
            计圈数据列表
        """
        laps_data = []
        
        if not streams.get('distance') or not streams.get('time'):
            return laps_data
        
        distance_data = streams['distance'].data
        time_data = streams['time'].data
        heartrate_data = streams.get('heartrate', {}).data if streams.get('heartrate') else None
        altitude_data = streams.get('altitude', {}).data if streams.get('altitude') else None
        
        current_lap_distance = 0.0
        lap_start_index = 0
        
        for i, distance in enumerate(distance_data):
            if distance - current_lap_distance >= lap_distance:
                # 计算这一圈的数据
                lap_data = self._calculate_lap_metrics(
                    distance_data[lap_start_index:i+1],
                    time_data[lap_start_index:i+1],
                    heartrate_data[lap_start_index:i+1] if heartrate_data else None,
                    altitude_data[lap_start_index:i+1] if altitude_data else None
                )
                laps_data.append(lap_data)
                current_lap_distance = distance
                lap_start_index = i + 1
        
        # 处理最后一圈（不足1公里的部分）
        if lap_start_index < len(distance_data):
            lap_data = self._calculate_lap_metrics(
                distance_data[lap_start_index:],
                time_data[lap_start_index:],
                heartrate_data[lap_start_index:] if heartrate_data else None,
                altitude_data[lap_start_index:] if altitude_data else None
            )
            laps_data.append(lap_data)
        
        return laps_data
    
    def _calculate_lap_metrics(
        self,
        distance_list: List[float],
        time_list: List[int],
        heartrate_list: Optional[List[int]] = None,
        altitude_list: Optional[List[float]] = None
    ) -> Dict:
        """
        计算单圈的指标
        
        Args:
            distance_list: 距离数据列表
            time_list: 时间数据列表
            heartrate_list: 心率数据列表（可选）
            altitude_list: 海拔数据列表（可选）
            
        Returns:
            单圈数据字典
        """
        if not distance_list or not time_list:
            return {}
        
        lap_distance = distance_list[-1] - distance_list[0]
        lap_time = time_list[-1] - time_list[0]
        
        lap_data = {
            'distance_meters': lap_distance,
            'time_seconds': float(lap_time),
        }
        
        # 计算平均和最大心率
        if heartrate_list and len(heartrate_list) > 0:
            valid_hr = [hr for hr in heartrate_list if hr and hr > 0]
            if valid_hr:
                lap_data['avg_heartrate'] = int(sum(valid_hr) / len(valid_hr))
                lap_data['max_heartrate'] = max(valid_hr)
        
        # 计算爬升和下降
        if altitude_list and len(altitude_list) > 1:
            elevation_gain = 0.0
            elevation_loss = 0.0
            for i in range(1, len(altitude_list)):
                diff = altitude_list[i] - altitude_list[i-1]
                if diff > 0:
                    elevation_gain += diff
                else:
                    elevation_loss += abs(diff)
            lap_data['elevation_gain'] = int(elevation_gain)
            lap_data['elevation_loss'] = int(elevation_loss)
        
        return lap_data
    
    def export_activity_from_strava(self, activity_id: int, client):
        """
        从Strava API导出活动为CSV
        
        Args:
            activity_id: 活动ID
            client: Strava客户端对象
        """
        try:
            laps_data, summary_data = self.parse_strava_activity(activity_id, client)
            
            if not laps_data:
                print(f"Warning: No lap data found for Strava activity {activity_id}")
                return
            
            csv_content = self.generate_csv(laps_data, summary_data)
            self.save_csv(activity_id, csv_content)
            
        except Exception as e:
            print(f"Error exporting Strava activity {activity_id}: {e}", file=sys.stderr)
            raise
    
    def export_activities_to_csv(
        self,
        activity_ids: List[int],
        platform: str,
        client=None,
        data_dir: str = None,
        file_suffix: str = None
    ):
        """
        导出多个活动为CSV（统一入口）
        
        Args:
            activity_ids: 活动ID列表
            platform: 平台名称（strava, garmin, tcx等）
            client: 平台客户端对象（对于API平台）
            data_dir: 数据文件目录（对于文件平台）
            file_suffix: 文件后缀（对于文件平台）
        """
        if not activity_ids:
            return
        
        print(f"Exporting {len(activity_ids)} activities to CSV from {platform}...")
        
        success_count = 0
        fail_count = 0
        
        for activity_id in activity_ids:
            try:
                if platform == 'strava' and client:
                    self.export_activity_from_strava(activity_id, client)
                    success_count += 1
                elif platform in ['tcx', 'gpx', 'fit'] and data_dir:
                    filename = f"{activity_id}.{file_suffix or platform}"
                    filepath = os.path.join(data_dir, filename)
                    
                    if not os.path.exists(filepath):
                        print(f"Warning: File not found: {filepath}")
                        fail_count += 1
                        continue
                    
                    if platform == 'tcx':
                        self.export_activity_from_tcx(activity_id, filepath)
                        success_count += 1
                    else:
                        print(f"Warning: Unsupported file format: {platform}")
                        fail_count += 1
                else:
                    print(f"Warning: Unsupported platform or missing parameters: {platform}")
                    fail_count += 1
                    
            except Exception as e:
                print(f"Error exporting activity {activity_id}: {e}", file=sys.stderr)
                fail_count += 1
        
        print(f"CSV export completed: {success_count} succeeded, {fail_count} failed")
