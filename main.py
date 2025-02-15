import csv
import re
import uuid
from datetime import datetime, timedelta
from icalendar import Calendar, Event, Alarm

def parse_time_slot(slot):
    """解析节次时间"""
    match = re.search(r'(\d+):(\d+)-(\d+):(\d+)', slot)
    if not match:
        return None, None
    start_h, start_m, end_h, end_m = map(int, match.groups())
    return (start_h, start_m), (end_h, end_m)

def parse_weeks(week_str):
    """解析周数规则的健壮版本"""
    try:
        is_odd = "单" in week_str
        is_even = "双" in week_str
        
        # 提取纯数字部分
        numbers = re.findall(r'\d+', week_str)
        if len(numbers) < 2:
            return []
            
        start = int(numbers[0])
        end = int(numbers[1])
        weeks = list(range(start, end+1))
        
        if is_odd:
            return [w for w in weeks if w % 2 == 1]
        if is_even:
            return [w for w in weeks if w % 2 == 0]
        return weeks
    except:
        return []


def get_course_info(cell):
    """解析课程单元格信息的优化版本"""
    # 分离课程主体和括号内容
    main_part, _, bracket_part = cell.strip().partition(')')
    if '(' in main_part:
        main_part, bracket_content = main_part.split('(', 1)
        bracket_part = bracket_content + ')' + bracket_part
    
    # 提取课程名称和教师
    parts = main_part.strip().rsplit(' ', 1)
    course_name = parts[0].strip()
    teacher = parts[1].strip() if len(parts) > 1 else ""
    
    # 解析括号内容
    weeks = []
    location = []
    if bracket_part:
        # 使用更精确的周数匹配模式
        week_matches = re.findall(r'([\d\-]+[单双]?周?)', bracket_part)
        for match in week_matches:
            weeks.extend(parse_weeks(match))
        
        # 移除已匹配的周数信息，剩余部分作为地点
        location_str = re.sub(r'([\d\-]+[单双]?周?)', '', bracket_part)
        location = [s.strip() for s in location_str.split() if s.strip()]
    
    return {
        "course": course_name,
        "teacher": teacher,
        "weeks": sorted(list(set(weeks))),
        "location": ' '.join(location)
    }

def parse_slot_number(slot_str):
    """精确提取节次数字"""
    match = re.search(r'^(\d{1,2})', slot_str)
    return match.group(1) if match else None

def generate_event(info, slot_num, week, weekday):
    """生成独立事件的核心函数"""
    event = Event()
    
    # 计算精确时间
    start_time_str, end_time_str = time_slots[slot_num]
    delta_days = (weekday - 1) + (week - 1) * 7
    event_date = start_date + timedelta(days=delta_days)
    
    # 时间格式标准化
    start_dt = datetime.combine(
        event_date, 
        datetime.strptime(start_time_str, "%H:%M").time()
    )
    end_dt = datetime.combine(
        event_date,
        datetime.strptime(end_time_str, "%H:%M").time()
    )
    
    # 设置事件属性
    event.add('summary', f"{info['course']} - {info['teacher']}")
    event.add('dtstart', start_dt)
    event.add('dtend', end_dt)
    event.add('location', info['location'])
    event.add('dtstamp', datetime.now())
    event.add('uid', f"{info['course']}_{slot_num}_{week}_{weekday}_{uuid.uuid4()}@courseschedule")
    
    return event

# 配置信息
start_date = datetime(2025, 2, 17)  # 第一周周一
# 修改后的时间配置（完整节次对应）
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

# 主处理循环修改部分
with open('courses.csv', encoding='utf-8-sig') as csvfile:
    reader = csv.DictReader(csvfile)
    # 处理列名中的异常空格
    reader.fieldnames = [name.strip().replace('\ufeff', '') for name in reader.fieldnames]  
    
    for row in reader:
        # 列名检查
        if '节次' not in row:
            continue
            
        # 精确获取节次数字
        slot_num = parse_slot_number(row['节次'])
        if not slot_num or slot_num not in time_slots:
            continue
            
        # 遍历每一天
        for weekday in range(1, 8):  # 1=周一,...,7=周日
            day_key = [
                '星期一','星期二','星期三',
                '星期四','星期五','星期六','星期日'
            ][weekday-1]
            
            cell = row.get(day_key, '').strip()
            if not cell or cell == ' ':
                continue
                
            info = get_course_info(cell)
            if not info or not info['weeks']:
                continue
                
            # 为每个周次生成独立事件
            for week in info['weeks']:
                # 跳过无效周次
                if week < 1 or week > 16:
                    continue
                    
                event = generate_event(info, slot_num, week, weekday)
                cal.add_component(event)


# 保存ICS文件
with open('course_schedule.ics', 'wb') as f:
    f.write(cal.to_ical())

print("ICS文件已生成：course_schedule.ics")
