import csv
from icalendar import Calendar, Event, vRecur
from datetime import datetime, timedelta
import re

def parse_time_slot(slot):
    """解析节次时间"""
    match = re.search(r'(\d+):(\d+)-(\d+):(\d+)', slot)
    if not match:
        return None, None
    start_h, start_m, end_h, end_m = map(int, match.groups())
    return (start_h, start_m), (end_h, end_m)

def parse_weeks(week_str):
    """解析周数规则，返回周数列表"""
    is_odd = "单周" in week_str
    is_even = "双周" in week_str
    week_str = week_str.replace("单周", "").replace("双周", "").strip()
    start, end = map(int, week_str.split('-'))
    weeks = list(range(start, end+1))
    
    if is_odd:
        return [w for w in weeks if w % 2 == 1]
    if is_even:
        return [w for w in weeks if w % 2 == 0]
    return weeks

def get_course_info(cell):
    """解析课程单元格信息"""
    match = re.match(r'(.*?)\s+([^(]+?)\s*\((.*?)\)', cell.strip())
    if not match:
        return None
    
    course_name = match.group(1).strip()
    teacher = match.group(2).strip()
    details = match.group(3).split()
    
    weeks = []
    location = ""
    for item in details:
        if any(c.isdigit() for c in item):
            weeks.extend(parse_weeks(item))
        else:
            location += item
    
    return {
        "course": course_name,
        "teacher": teacher,
        "weeks": sorted(list(set(weeks))),
        "location": location
    }

# 配置信息
start_date = datetime(2025, 2, 17)  # 第一周周一
time_slots = {
    '1': ('08:00', '08:45'),
    '2': ('08:50', '09:35'),
    '3': ('09:55', '10:40'),
    '4': ('10:45', '11:30'),
    '5': ('11:35', '12:20'),
    '6': ('14:00', '14:45'),
    '7': ('14:50', '15:35'),
    '8': ('15:55', '16:40'),
    '9': ('16:45', '17:30'),
    '10': ('18:40', '19:25'),
    '11': ('19:30', '20:15'),
    '12': ('20:20', '21:05')
}

cal = Calendar()
cal.add('prodid', '-//Course Schedule//mxm.dk//')
cal.add('version', '2.0')

# 修改点1：使用utf-8-sig编码读取文件（处理BOM头）
with open('courses.csv', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile)
    
    # 修改点2：清洗字段名（处理可能的空格或不可见字符）
    fieldnames = [name.strip() for name in reader.fieldnames]
    reader = csv.DictReader(csvfile, fieldnames=fieldnames)
    next(reader)  # 跳过原始标题行
    
    for row in reader:
        # 修改点3：添加空值过滤
        if not row.get('节次'):
            continue
            
        slot = row['节次']
        if not slot:
            continue
            
        # 获取节次数字（处理类似"1(08:00-08:45)"的格式）
        slot_num = slot[0]
        
        # 处理每一天的课程
        for weekday in range(1, 8):
            day_key = ['星期一','星期二','星期三','星期四','星期五','星期六','星期日'][weekday-1]
            cell = row.get(day_key, '').strip()
            if not cell or cell == ' ':
                continue
                
            info = get_course_info(cell)
            if not info or not info['weeks']:
                continue
                
            # 创建事件
            event = Event()
            event.add('summary', f"{info['course']} - {info['teacher']}")
            event.add('location', info['location'])
            
            # 设置时间
            start_time_str, end_time_str = time_slots[slot_num]
            
            for week in info['weeks']:
                # 计算具体日期
                delta_days = (weekday - 1) + (week - 1) * 7
                event_date = start_date + timedelta(days=delta_days)
                
                # 合并时间
                start_time = datetime.strptime(start_time_str, "%H:%M").time()
                end_time = datetime.strptime(end_time_str, "%H:%M").time()
                
                event.add('dtstart', datetime.combine(event_date, start_time))
                event.add('dtend', datetime.combine(event_date, end_time))
                event.add('dtstamp', datetime.now())
                
                cal.add_component(event)

# 保存ICS文件
with open('course_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())

print("ICS文件已生成：course_schedule.ics")
