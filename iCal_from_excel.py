import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
import re
from openpyxl import load_workbook
import warnings
from typing import Dict, Optional, List

# 忽略openpyxl样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl.styles.stylesheet')

def read_excel_raw(file_path: str) -> pd.DataFrame:
    """读取Excel文件，保留原始格式"""
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    return pd.DataFrame(list(ws.values))

def parse_course_info(cell_content: str) -> Optional[Dict[str, str]]:
    """
    增强版课程信息解析，支持多种格式：
    - 课程名\n[教师]\n[班级]\n[周次][地点][时间]
    - 课程名\n[教师]\n[班级]\n[时间][地点][周次]
    - 以及其他变体
    """
    if not cell_content or not isinstance(cell_content, str) or cell_content.strip() == "":
        return None
    
    try:
        # 统一处理换行符
        content = cell_content.replace('\r\n', '\n').strip()
        lines = [line.strip() for line in content.split('\n') if line.strip()]
        
        if len(lines) < 4:  # 至少需要课程名、教师、班级和详细信息
            return None
        
        result = {
            'course_name': lines[0].split('[')[0].strip(),
            'teacher': re.sub(r'^\[|\]$', '', lines[1]) if len(lines) > 1 else "",
            'class_info': re.sub(r'^\[|\]$', '', lines[2]) if len(lines) > 2 else "",
            'week_info': "",
            'location': "",
            'time_info': ""
        }
        
        # 解析最后一行中的[]内容
        last_line = lines[3]
        bracket_contents = re.findall(r'\[(.*?)\]', last_line)
        
        # 智能分配周次、地点和时间
        for content in bracket_contents:
            # 识别周次
            if re.match(r'^(\d+-\d+|\d+)(单周|双周|周)?$', content):
                result['week_info'] = content
            # 识别时间
            elif re.match(r'^\d+-\d+节$', content):
                result['time_info'] = content
            # 识别地点（包含特定关键词）
            elif any(keyword in content for keyword in ['栋', '教', '楼', '园', '场', '实验室']):
                result['location'] = content
            # 默认分配
            elif not result['week_info'] and re.match(r'^\d', content):
                result['week_info'] = content
            elif not result['time_info'] and re.match(r'^\d+-\d+', content):
                result['time_info'] = content
            else:
                result['location'] = content
        
        return result
    except Exception as e:
        print(f"解析课程信息出错: {e}\n内容: {cell_content}")
        return None

def parse_weeks(week_str: str) -> List[int]:
    """
    增强版周次解析，支持：
    - 单周/双周 (1-15单周)
    - 连续范围 (2-16周)
    - 不连续范围 (1-8,10-16周)
    - 单一周次 (15周)
    """
    if not week_str or not isinstance(week_str, str):
        return []
    
    try:
        week_str = week_str.strip()
        is_odd = '单周' in week_str
        is_even = '双周' in week_str
        week_str = week_str.replace('单周', '').replace('双周', '').replace('周', '')
        
        weeks = []
        for part in week_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                for week in range(start, end + 1):
                    if (is_odd and week % 2 == 1) or (is_even and week % 2 == 0) or (not is_odd and not is_even):
                        weeks.append(week)
            elif part.isdigit():
                week = int(part)
                if (is_odd and week % 2 == 1) or (is_even and week % 2 == 0) or (not is_odd and not is_even):
                    weeks.append(week)
        
        return sorted(list(set(weeks)))
    except Exception as e:
        print(f"解析周次出错: {e}\n内容: {week_str}")
        return []

def create_ical_from_excel(file_path: str, output_file: str = 'course_schedule.ics') -> int:
    """从Excel创建iCal文件"""
    try:
        df = read_excel_raw(file_path)
        print(f"成功读取Excel文件，共{len(df)}行数据")
    except Exception as e:
        print(f"读取Excel失败: {e}")
        return 0
    
    # 初始化日历
    cal = Calendar()
    cal.add('prodid', '-//Course Schedule//mxm.dk//')
    cal.add('version', '2.0')
    
    # 配置参数
    semester_start = datetime(2025, 2, 17)  # 学期开始日期
    weekday_map = {'星期一': 0, '星期二': 1, '星期三': 2, '星期四': 3, 
                  '星期五': 4, '星期六': 5, '星期日': 6}
    class_time_map = {
        '1-2节': (8, 0, 9, 50),
        '3-4节': (10, 10, 12, 0),
        '5-6节': (14, 0, 15, 50),
        '7-8节': (16, 0, 17, 50),
        '9-10节': (19, 0, 20, 50)
    }
    
    event_count = 0
    
    # 解析数据
    for row_idx in range(3, len(df)):  # 从第4行开始
        time_slot = df.iloc[row_idx, 0]
        if not time_slot or '第' not in str(time_slot):
            continue
        
        # 解析时间段
        time_match = re.search(r'第(\d+)-(\d+)节', str(time_slot))
        if not time_match or time_match.group(1) == time_match.group(2):
            continue
        
        time_key = f"{time_match.group(1)}-{time_match.group(2)}节"
        if time_key not in class_time_map:
            print(f"忽略未知时间段: {time_slot}")
            continue
        
        # 处理每个weekday
        for col_idx in range(1, 8):
            weekday = df.iloc[2, col_idx]
            if not weekday or weekday not in weekday_map:
                continue
            
            cell_content = df.iloc[row_idx, col_idx]
            if not cell_content or str(cell_content).strip() == "":
                continue
            
            # 解析课程信息
            course_info = parse_course_info(cell_content)
            if not course_info:
                print(f"解析失败，内容: {cell_content}")
                continue
            
            # 解析周次
            weeks = parse_weeks(course_info['week_info'])
            if not weeks:
                print(f"无有效周次: {course_info['week_info']}")
                continue
            
            # 创建事件
            for week in weeks:
                try:
                    # 计算日期
                    course_date = semester_start + timedelta(
                        weeks=week-1,
                        days=weekday_map[weekday]
                    )
                    
                    # 设置时间
                    start_h, start_m, end_h, end_m = class_time_map[time_key]
                    start_time = course_date.replace(hour=start_h, minute=start_m)
                    end_time = course_date.replace(hour=end_h, minute=end_m)
                    
                    # 创建事件
                    event = Event()
                    event.add('summary', course_info['course_name'])
                    event.add('description',
                            f"教师: {course_info['teacher']}\n"
                            f"班级: {course_info['class_info']}\n"
                            f"周次: 第{week}周\n"
                            f"时间: {time_key}\n"
                            f"地点: {course_info['location']}")
                    event.add('location', course_info['location'])
                    event.add('dtstart', start_time)
                    event.add('dtend', end_time)
                    event.add('dtstamp', datetime.now())
                    
                    cal.add_component(event)
                    event_count += 1
                    
                except Exception as e:
                    print(f"创建事件失败: {e}\n课程: {course_info}")
    
    # 保存文件
    if event_count > 0:
        with open(output_file, 'wb') as f:
            f.write(cal.to_ical())
        print(f"成功创建 {event_count} 个日历事件，已保存到 {output_file}")
    else:
        print("未创建任何日历事件，请检查输入文件格式")
    
    return event_count

# 使用示例
if __name__ == "__main__":
    input_file = r'C:\我的大学\课表\export.xlsx'
    create_ical_from_excel(input_file)