import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# 加载 JSON 数据
try:
    with open('grades.json', 'r', encoding='utf-8') as file:
        data = json.load(file)
    print(f"Loaded JSON data with {len(data)} semesters")
except FileNotFoundError:
    print("Error: grades.json not found")
    exit(1)
except json.JSONDecodeError:
    print("Error: Invalid JSON format in grades.json")
    exit(1)

# 定义字段及其对应的中文标签
fields = [
    ("semesterName", "学期名称"),
    ("courseCode", "课程代码"),
    ("courseName", "课程名称"),
    ("courseNameEn", "课程英文名称"),
    ("lessonCode", "课程序号"),
    ("credits", "学分"),
    ("courseType", "课程类型"),
    ("courseProperty", "课程属性"),
    ("gaGrade", "成绩"),
    ("passed", "是否通过"),
    ("gp", "绩点"),
    ("gradeDetail", "成绩详情"),
    ("published", "是否公布"),
    ("fillAGrace", "补考成绩"),
    ("compulsory", "是否必修"),
    ("courseModuleTypeName", "课程模块类型")
]

# 提取英文字段名作为 DataFrame 的初始列名
english_fields = [field for field, _ in fields]

# 创建工作簿和工作表
wb = Workbook()
ws = wb.active
ws.title = "成绩总览"

# 准备数据
rows = []
for semester_data in data:
    try:
        # 检查 semesterId2studentGrades 是否为空
        if not semester_data["semesterId2studentGrades"]:
            continue
        semester_id = list(semester_data["semesterId2studentGrades"].keys())[0]
        # 跳过没有成绩的学期
        if not semester_data["semesterId2studentGrades"][semester_id]:
            continue

        semester_name = semester_data["semesters"][0]["nameZh"]
        grades = semester_data["semesterId2studentGrades"][semester_id]
        print(f"Processing semester {semester_name} (ID: {semester_id}), {len(grades)} grades found")

        for grade in grades:
            # 使用英文字段名作为字典的键
            row = {field: grade.get(field, None) for field, _ in fields}
            row["gradeDetail"] = "; ".join(filter(None, row["gradeDetail"])) if row.get("gradeDetail") else ""
            rows.append(row)
            print(f"Added row for course: {row['courseCode']}")
    except (KeyError, IndexError) as e:
        print(f"Error processing semester: {e}")
        continue

# 检查 rows 是否为空
if not rows:
    print("Warning: No data rows were added. Check JSON data for grades.")
    ws.append(["无数据"])  # 添加提示信息到 Excel
else:
    # 1. 使用英文字段名创建 DataFrame
    df = pd.DataFrame(rows, columns=english_fields)

    # 2. 将列名重命名为中文标签
    df.columns = [chinese_label for _, chinese_label in fields]

    # 写入 DataFrame 到工作表
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

# 保存工作簿
wb.save("grades.xlsx")
print("Excel file saved as grades.xlsx")