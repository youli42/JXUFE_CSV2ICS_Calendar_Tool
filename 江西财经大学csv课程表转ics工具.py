import csv
import re
import uuid
import os
from datetime import datetime, timedelta
from tkinter import Tk, filedialog, messagebox
from tkinter.ttk import Progressbar
from icalendar import Calendar, Event

# ---------------------------- 核心解析逻辑 ----------------------------
def parse_time_slot(slot):
    """解析节次时间"""
    match = re.search(r'(\d+):(\d+)-(\d+):(\d+)', slot)
    if not match:
        return None, None
    start_h, start_m, end_h, end_m = map(int, match.groups())
    return (start_h, start_m), (end_h, end_m)

def parse_weeks(week_str):
    """解析周数规则"""
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
    """解析课程单元格信息"""
    cell = cell.strip()
    if not cell or cell == ' ':
        return None
    
    # 分离主体和括号内容
    main_part, bracket_part = None, None
    if '(' in cell:
        parts = cell.split('(', 1)
        main_part = parts[0].strip()
        bracket_part = parts[1].split(')')[0].strip()
    else:
        main_part = cell
    
    # 解析课程名称和教师
    name_teacher = main_part.rsplit(' ', 1)
    course_name = name_teacher[0].strip()
    teacher = name_teacher[1].strip() if len(name_teacher) > 1 else ""
    
    # 解析周次和地点
    weeks = []
    location = []
    if bracket_part:
        # 提取周次信息
        week_matches = re.findall(r'([\d\-]+[单双]?周?)', bracket_part)
        for match in week_matches:
            weeks.extend(parse_weeks(match))
        
        # 剩余部分作为地点
        location = re.sub(r'([\d\-]+[单双]?周?)', '', bracket_part).strip()
    
    return {
        "course": course_name,
        "teacher": teacher,
        "weeks": sorted(list(set(weeks))),
        "location": location
    }

# ---------------------------- 图形界面组件 ----------------------------
class CourseConverterGUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("课程表转换工具 v2.0")
        self.root.withdraw()  # 隐藏主窗口
        
        self.time_slots = {
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
        
        self.start_date = datetime(2025, 2, 17)
        self.select_file()

    def select_file(self):
        """文件选择对话框"""
        file_path = filedialog.askopenfilename(
            title="选择课程表CSV文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            self.convert_file(file_path)
        else:
            if messagebox.askretrycancel("提示", "您未选择文件，是否重试？"):
                self.select_file()
            else:
                self.root.destroy()

    def create_event(self, course_info, slot_num, week, weekday):
        """创建日历事件"""
        event = Event()
        
        # 计算时间
        start_str, end_str = self.time_slots.get(slot_num, ('08:00', '08:45'))
        delta_days = (weekday - 1) + (week - 1) * 7
        event_date = self.start_date + timedelta(days=delta_days)
        
        # 设置时间
        start_time = datetime.strptime(start_str, "%H:%M").time()
        end_time = datetime.strptime(end_str, "%H:%M").time()
        
        event.add('summary', f"{course_info['course']} - {course_info['teacher']}")
        event.add('dtstart', datetime.combine(event_date, start_time))
        event.add('dtend', datetime.combine(event_date, end_time))
        event.add('location', course_info['location'])
        event.add('uid', f"{uuid.uuid4()}@courseschedule")
        return event

    def convert_file(self, csv_path):
        """执行转换的核心方法"""
        try:
            # 创建进度窗口
            progress_win = Tk()
            progress_win.title("转换进度")
            progress_win.geometry("300x60")
            Progressbar(progress_win, mode='indeterminate').pack(pady=15)
            progress_win.update()
            
            # 初始化日历
            cal = Calendar()
            cal.add('prodid', '-//Course Schedule//mxm.dk//')
            cal.add('version', '2.0')
            
            # 处理CSV内容
            with open(csv_path, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    slot = row.get('节次', '')
                    if not slot:
                        continue
                    
                    # 提取节次数字
                    slot_num = re.search(r'^\d+', slot)
                    if not slot_num:
                        continue
                    slot_num = slot_num.group()
                    
                    # 处理每一天的课程
                    for weekday in range(1, 8):
                        day_key = ['星期一','星期二','星期三','星期四','星期五','星期六','星期日'][weekday-1]
                        cell = row.get(day_key, '')
                        if not cell:
                            continue
                        
                        course_info = get_course_info(cell)
                        if not course_info or not course_info['weeks']:
                            continue
                        
                        # 生成每个周次的事件
                        for week in course_info['weeks']:
                            event = self.create_event(course_info, slot_num, week, weekday)
                            cal.add_component(event)
            
            # 保存文件
            output_path = os.path.splitext(csv_path)[0] + ".ics"
            with open(output_path, 'wb') as f:
                f.write(cal.to_ical())
            
            # 关闭进度窗口
            progress_win.destroy()
            
            # 显示完成提示
            messagebox.showinfo(
                "转换成功",
                f"日历文件已保存至：\n{output_path}",
                parent=self.root
            )
            
        except Exception as e:
            messagebox.showerror(
                "发生错误",
                f"错误信息：{str(e)}\n请检查文件格式是否符合要求",
                parent=self.root
            )
        finally:
            self.root.destroy()

if __name__ == "__main__":
    app = CourseConverterGUI()
