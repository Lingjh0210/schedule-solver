"""
排课求解器 Web UI
基于 Streamlit 框架
"""

import streamlit as st
import pandas as pd
import re
from pathlib import Path
import io
import time
import threading
try:
    from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx
except ImportError:
    from streamlit.scriptrunner import add_script_run_ctx, get_script_run_ctx
from ortools.sat.python import cp_model
from collections import defaultdict
from openpyxl.utils import get_column_letter
from datetime import datetime

st.set_page_config(
    page_title="智能排课求解器",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        color: #1f77b4;
        padding: 1rem 0;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #ff7f0e;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        margin: 1rem 0;
    }
    .error-box {
        padding: 1rem;
        background-color: #f8d7da;
        border-left: 5px solid #dc3545;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border-left: 5px solid #17a2b8;
        margin: 1rem 0;
    }
    .save-section {
        padding: 1rem;
        background-color: #e7f3ff;
        border: 1px solid #2196F3;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def natural_sort_key(s):
    """自然排序的key函数，用于正确排序包含数字的字符串
    例如: S1, S2, S3, ..., S9, S10, S11 (而不是 S1, S10, S11, S2)
    """
    import re
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split(r'(\d+)', str(s))]

def parse_subject_string(subject_str):
    """解析科目字符串（支持中英文括号及空格）"""
    subjects = {}
    # 增加 \s* 允许括号周围有空格
    pattern = r'([^,\(（]+)\s*[\(（]\s*(\d+)\s*[\)）]'
    matches = re.findall(pattern, subject_str)
    for subject, hours in matches:
        subject = subject.strip()
        subjects[subject] = int(hours)
    return subjects

# 初始化 session_state 用于保存方案
if 'saved_solutions' not in st.session_state:
    st.session_state['saved_solutions'] = {}

def save_solution_to_storage(sol, save_name):
    """保存方案到存储"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state['saved_solutions'][save_name] = {
        'solution': sol,
        'timestamp': timestamp,
        'original_name': sol['name']
    }

def delete_saved_solution(save_name):
    """删除已保存的方案"""
    if save_name in st.session_state['saved_solutions']:
        del st.session_state['saved_solutions'][save_name]

# Read Excel File
def parse_uploaded_file(uploaded_file):
    """解析上传的Excel/CSV文件"""
    try:
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
        else:
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'cp936', 'utf-8-sig']
            df = None
            last_error = None
            
            for encoding in encodings:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding=encoding)
                    st.success(f"✅ 成功读取文件（编码：{encoding}）")
                    break
                except (UnicodeDecodeError, Exception) as e:
                    last_error = e
                    continue
            
            if df is None:
                raise Exception(f"无法识别文件编码，请确保文件是有效的CSV格式。最后错误：{last_error}")
        
        packages = {}
        subject_hours = {}
        total_hours_stats = []
        
        for _, row in df.iterrows():
            package_name = str(row['配套']).strip()
            student_count = int(row['人数'])
            subject_str = str(row['科目'])
            
            subjects = parse_subject_string(subject_str)
            
            total_hours = sum(subjects.values())
            total_hours_stats.append({
                '配套': package_name,
                '总课时': total_hours
            })
            
            packages[package_name] = {
                '人数': student_count,
                '科目': subjects
            }
            
            for subject, hours in subjects.items():
                if subject not in subject_hours:
                    subject_hours[subject] = hours
                elif subject_hours[subject] != hours:
                    st.error(f"❌ **数据错误：科目'{subject}'的课时不一致！**")
                    st.error(f"   • 在某些配套中是 **{subject_hours[subject]}小时**")
                    st.error(f"   • 在'{package_name}'配套中是 **{hours}小时**")
                    st.markdown("---")
                    st.markdown("""
                    ### 🔍 为什么会导致错误？
                    
                    系统会为每个科目创建**统一长度**的班级（如6小时的会计班）。
                    所有学生都会被分配到这些统一的班级中。
                    
                    如果配套A需要6小时会计，配套B需要4小时会计：
                    - ❌ 无法用6小时的班满足4小时的需求
                    - ❌ 也无法用4小时的班满足6小时的需求
                    - ❌ 导致求解器找不到可行解
                    
                    ### ✅ 解决方案：
                    
                    **方案1：统一课时（推荐）**
                    - 将所有配套的'{subject}'课时改为相同值（如都改为6小时或都改为4小时）
                    
                    **方案2：分离科目**
                    - 将4小时的会计命名为"会计1"
                    - 将6小时的会计命名为"会计2"
                    - 这样系统会将它们视为不同科目
                    """)
                    return None, None, None
        
        min_hours = min(s['总课时'] for s in total_hours_stats)
        max_hours = max(s['总课时'] for s in total_hours_stats)
        
        if min_hours < 21:
            st.info(f"ℹ️ 检测到部分配套总课时少于21小时（范围：{min_hours}-{max_hours}小时）")
            st.success("✅ 系统支持总课时不足的配套，这些配套将在某些时段不上课")
            
            short_packages = [s for s in total_hours_stats if s['总课时'] < 21]
            if short_packages:
                with st.expander("查看总课时不足21的配套"):
                    for pkg in short_packages:
                        st.write(f"  • {pkg['配套']}: {pkg['总课时']}小时")
        
        return packages, subject_hours, total_hours_stats
        
    except Exception as e:
        st.error(f"❌ 文件解析错误：{str(e)}")
        return None, None, None

def split_large_packages(packages, max_students_per_package=25):
    """
    将人数过多的配套拆分为 A/B 班
    """
    new_packages = {}
    split_log = []
    
    for pkg_name, pkg_info in packages.items():
        student_count = pkg_info['人数']
        
        if student_count > max_students_per_package:
            # 需要拆分
            num_splits = (student_count + max_students_per_package - 1) // max_students_per_package
            students_per_split = student_count // num_splits
            remainder = student_count % num_splits
            
            parts = []
            for i in range(num_splits):
                split_size = students_per_split + (1 if i < remainder else 0)
                suffix = chr(65 + i)  # A, B, C...
                split_name = f"{pkg_name}_{suffix}"
                new_packages[split_name] = {
                    '人数': split_size,
                    '科目': pkg_info['科目'].copy()
                }
                parts.append(f"{split_name}({split_size}人)")
            
            split_log.append({
                'original': pkg_name,
                'total': student_count,
                'parts': parts
            })
        else:
            # 不需要拆分
            new_packages[pkg_name] = pkg_info
    
    return new_packages, split_log

def build_schedule_model(packages, subject_hours, num_time_slots=7, 
                         min_class_size=None, max_class_size=None,
                         allow_split=False,
                         force_open_classes=None):
    """
    Build CP model
    """
    model = cp_model.CpModel()
    
    packages_sorted = sorted(packages.keys(), key=natural_sort_key)
    subjects_sorted = sorted(subject_hours.keys(), key=natural_sort_key)
    slots_range = range(num_time_slots)
    
    # 动态推断班额范围
    all_sizes = [pkg['人数'] for pkg in packages.values()]
    if not min_class_size:
        min_class_size = min(all_sizes)
    if not max_class_size:
        max_class_size = max(max(all_sizes), 35)
    
    # 估算最大可能班级数
    total_students = sum(pkg['人数'] for pkg in packages.values())
    max_classes_per_subject = (total_students // min_class_size) + 5
    
    # Variables
    class_exists = {}
    class_size = {}
    class_slot = {}
    
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject):
            class_exists[(subject, c)] = model.NewBoolVar(f'exists_{subject}_{c}')
            class_size[(subject, c)] = model.NewIntVar(0, max_class_size, f'size_{subject}_{c}')
            class_slot[(subject, c)] = model.NewIntVar(0, num_time_slots - 1, f'slot_{subject}_{c}')
    
    package_assignment = {}
    for package in packages_sorted:
        for subject in packages[package]['科目'].keys():
            for c in range(max_classes_per_subject):
                package_assignment[(package, subject, c)] = model.NewBoolVar(
                    f'assign_{package}_{subject}_{c}'
                )
    
    # Constraints
    # 1. 每个配套的每个科目只能分配到一个班级
    for package in packages_sorted:
        for subject in packages[package]['科目'].keys():
            model.Add(sum(package_assignment[(package, subject, c)] 
                         for c in range(max_classes_per_subject)) == 1)
    
    # 2. 班级存在性
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject):
            for package in packages_sorted:
                if subject in packages[package]['科目']:
                    model.Add(package_assignment[(package, subject, c)] <= class_exists[(subject, c)])
    
    # 3. 班级人数计算
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject):
            model.Add(class_size[(subject, c)] == sum(
                package_assignment[(package, subject, c)] * packages[package]['人数']
                for package in packages_sorted if subject in packages[package]['科目']
            ))
    
    # 4. 班额约束
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject):
            model.Add(class_size[(subject, c)] >= min_class_size).OnlyEnforceIf(class_exists[(subject, c)])
            model.Add(class_size[(subject, c)] <= max_class_size).OnlyEnforceIf(class_exists[(subject, c)])
            model.Add(class_size[(subject, c)] == 0).OnlyEnforceIf(class_exists[(subject, c)].Not())
    
    # 5. 强制开班数约束
    if force_open_classes:
        for subject, num_classes in force_open_classes.items():
            if subject in subjects_sorted:
                model.Add(sum(class_exists[(subject, c)] for c in range(max_classes_per_subject)) == num_classes)
    
    # 6. 同一配套的不同科目不能在同一时段
    for package in packages_sorted:
        subjects_in_package = list(packages[package]['科目'].keys())
        for i, subj1 in enumerate(subjects_in_package):
            for subj2 in subjects_in_package[i+1:]:
                for c1 in range(max_classes_per_subject):
                    for c2 in range(max_classes_per_subject):
                        b1 = package_assignment[(package, subj1, c1)]
                        b2 = package_assignment[(package, subj2, c2)]
                        
                        same_slot = model.NewBoolVar(f'same_slot_{package}_{subj1}_{c1}_{subj2}_{c2}')
                        model.Add(class_slot[(subj1, c1)] == class_slot[(subj2, c2)]).OnlyEnforceIf(same_slot)
                        model.Add(class_slot[(subj1, c1)] != class_slot[(subj2, c2)]).OnlyEnforceIf(same_slot.Not())
                        
                        both_assigned = model.NewBoolVar(f'both_{package}_{subj1}_{c1}_{subj2}_{c2}')
                        model.AddBoolAnd([b1, b2]).OnlyEnforceIf(both_assigned)
                        model.AddBoolOr([b1.Not(), b2.Not()]).OnlyEnforceIf(both_assigned.Not())
                        
                        model.AddBoolOr([both_assigned.Not(), same_slot.Not()])
    
    # 7. 班级编号连续性
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject - 1):
            model.Add(class_exists[(subject, c)] >= class_exists[(subject, c + 1)])
    
    # 优化目标
    total_classes = sum(class_exists[(subject, c)] 
                       for subject in subjects_sorted 
                       for c in range(max_classes_per_subject))
    
    variance_terms = []
    for subject in subjects_sorted:
        for c in range(max_classes_per_subject):
            deviation = model.NewIntVar(-max_class_size, max_class_size, f'dev_{subject}_{c}')
            abs_deviation = model.NewIntVar(0, max_class_size, f'abs_dev_{subject}_{c}')
            target_size = (min_class_size + max_class_size) // 2
            
            model.Add(deviation == class_size[(subject, c)] - target_size).OnlyEnforceIf(class_exists[(subject, c)])
            model.AddAbsEquality(abs_deviation, deviation)
            
            variance_terms.append(abs_deviation)
    
    total_variance = sum(variance_terms)
    
    model.Minimize(total_classes * 10000 + total_variance)
    
    return model, {
        'class_exists': class_exists,
        'class_size': class_size,
        'class_slot': class_slot,
        'package_assignment': package_assignment,
        'subjects': subjects_sorted,
        'packages': packages_sorted,
        'slots_range': slots_range,
        'max_classes': max_classes_per_subject
    }

def solve_schedule(packages, subject_hours, num_time_slots=7, 
                   min_class_size=None, max_class_size=None,
                   allow_split=False, force_open_classes=None, time_limit=120):
    """
    Solve model
    """
    model, variables = build_schedule_model(
        packages, subject_hours, num_time_slots,
        min_class_size, max_class_size, allow_split, force_open_classes
    )
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = 8
    
    start_time = time.time()
    status = solver.Solve(model)
    solve_time = time.time() - start_time
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        class_details = []
        slot_details = defaultdict(list)
        
        for subject in variables['subjects']:
            for c in range(variables['max_classes']):
                if solver.Value(variables['class_exists'][(subject, c)]):
                    size = solver.Value(variables['class_size'][(subject, c)])
                    slot = solver.Value(variables['class_slot'][(subject, c)])
                    
                    packages_in_class = []
                    for package in variables['packages']:
                        if subject in packages[package]['科目']:
                            if solver.Value(variables['package_assignment'][(package, subject, c)]):
                                packages_in_class.append(package)
                    
                    class_details.append({
                        '科目': subject,
                        '班级': f"{c+1}班",
                        '人数': size,
                        '时段': f"S{slot+1}",
                        '学生配套': ', '.join(sorted(packages_in_class, key=natural_sort_key))
                    })
                    
                    slot_details[slot].append({
                        '科目': subject,
                        '班级': f"{c+1}班",
                        '人数': size,
                        '配套': packages_in_class
                    })
        
        return {
            'status': 'success',
            'class_details': class_details,
            'slot_details': slot_details,
            'solve_time': solve_time
        }
    else:
        return {
            'status': 'failed',
            'solve_time': solve_time
        }

def analyze_teacher_needs(slot_schedule):
    """分析每个科目需要的老师数（最大并发数）"""
    teacher_needs = defaultdict(int)
    
    for slot_data in slot_schedule:
        slot = slot_data['时段']
        subject_count = defaultdict(int)
        
        for item in slot_data.get('display_items', []):
            if not item.get('is_gap', False):
                subject = item.get('subject', '')
                if subject:
                    subject_count[subject] += 1
        
        for subject, count in subject_count.items():
            teacher_needs[subject] = max(teacher_needs[subject], count)
    
    return teacher_needs

def save_history_to_disk(solutions):
    """保存求解历史到本地（占位函数）"""
    pass

def analyze_solution(class_details):
    """分析方案统计信息"""
    if not class_details:
        return {
            'total_classes': 0,
            'avg_size': 0,
            'min_size': 0,
            'max_size': 0,
            'split_count': 0,
            'split_details': []
        }
    
    sizes = [c['人数'] for c in class_details]
    
    # 检测时段分割
    slot_groups = defaultdict(list)
    for detail in class_details:
        slot = detail['时段']
        packages = detail['学生配套'].split(', ')
        for pkg in packages:
            slot_groups[pkg].append(slot)
    
    split_count = 0
    split_details = []
    for pkg, slots in slot_groups.items():
        unique_slots = set(slots)
        if len(slots) > len(unique_slots):
            split_count += len(slots) - len(unique_slots)
            split_details.append(f"{pkg} 在 {', '.join(sorted(unique_slots))} 有重复")
    
    return {
        'total_classes': len(class_details),
        'avg_size': round(sum(sizes) / len(sizes), 1) if sizes else 0,
        'min_size': min(sizes) if sizes else 0,
        'max_size': max(sizes) if sizes else 0,
        'split_count': split_count,
        'split_details': split_details
    }

def format_slot_schedule(slot_details, packages, subject_hours):
    """格式化时段表为展示格式"""
    schedule_data = []
    
    for slot in sorted(slot_details.keys()):
        classes = slot_details[slot]
        
        # 按科目分组
        subject_groups = defaultdict(list)
        for cls in classes:
            subject_groups[cls['科目']].append(cls)
        
        for subject in sorted(subject_groups.keys(), key=natural_sort_key):
            subject_classes = subject_groups[subject]
            hours = subject_hours[subject]
            
            for cls in subject_classes:
                display_items = []
                packages_in_class = cls['配套']
                
                # 为每个配套创建时间轴项
                for pkg in sorted(packages_in_class, key=natural_sort_key):
                    display_items.append({
                        'subject': subject,
                        'class_name': cls['班级'],
                        'package': pkg,
                        'duration': f"{hours}h",
                        'start_offset': 0,
                        'relative_slots': list(range(hours)),
                        'is_gap': False,
                        'packages_str': pkg
                    })
                
                schedule_data.append({
                    '时段': f"S{slot+1}",
                    '时长': f"{hours}h",
                    '科目': subject,
                    '班级': cls['班级'],
                    '涉及配套': ', '.join(sorted(packages_in_class, key=natural_sort_key)),
                    'display_items': display_items,
                    'sort_key_subject': natural_sort_key(subject)
                })
    
    # 排序
    schedule_data.sort(key=lambda x: (x['时段'], x['sort_key_subject']))
    
    return schedule_data

def main():
    st.markdown('<div class="main-header">📚 智能排课求解器</div>', unsafe_allow_html=True)
    
    # 侧边栏 - 已保存的方案
    with st.sidebar:
        st.markdown("### 💾 已保存的方案")
        
        if st.session_state['saved_solutions']:
            st.markdown(f"**共有 {len(st.session_state['saved_solutions'])} 个已保存方案**")
            
            for save_name, saved_data in st.session_state['saved_solutions'].items():
                with st.expander(f"📁 {save_name}"):
                    st.markdown(f"**原方案名称:** {saved_data['original_name']}")
                    st.markdown(f"**保存时间:** {saved_data['timestamp']}")
                    
                    col1, col2 = st.columns(2)
                    
                    # 下载按钮
                    with col1:
                        sol = saved_data['solution']
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            raw_class_data = sol['class_details']
                            raw_slot_data = sol['slot_schedule']
                            
                            df_class = pd.DataFrame(raw_class_data)
                            
                            def format_subject_class_col(row):
                                suffix = row['班级'].replace('班', '')
                                if suffix:
                                    return f"{row['科目']} {suffix}"
                                else:
                                    return row['科目']
                            
                            df_class = df_class.sort_values(by=['科目', '班级'])
                            df_class['科目 & 班级'] = df_class.apply(format_subject_class_col, axis=1)
                            df_class_export = df_class[['科目 & 班级', '人数', '时段', '学生配套']]
                            df_class_export.to_excel(writer, sheet_name='开班详情', index=False)
                            
                            df_slot = pd.DataFrame(raw_slot_data)
                            p1_list, p2_list, p3_list = [], [], []
                            
                            for item in raw_slot_data:
                                current_pkg_slots = ["-", "-", "-"]
                                d_items = item.get('display_items', [])
                                
                                if isinstance(d_items, list):
                                    for sub_item in d_items:
                                        pkg_str = sub_item.get('packages_str', '-')
                                        if not pkg_str or sub_item.get('is_gap', False):
                                            pkg_str = "-"
                                        
                                        rel_slots = sub_item.get('relative_slots', [])
                                        if not rel_slots and 'start_offset' in sub_item:
                                            try: dur = int(sub_item['duration'].replace('h',''))
                                            except: dur = 1
                                            start = sub_item['start_offset']
                                            rel_slots = range(start, start + dur)
                                            
                                        for idx in rel_slots:
                                            if 0 <= idx < 3:
                                                current_pkg_slots[idx] = pkg_str
                                
                                p1_list.append(current_pkg_slots[0])
                                p2_list.append(current_pkg_slots[1])
                                p3_list.append(current_pkg_slots[2])
                            
                            df_slot['配套 (第1小时)'] = p1_list
                            df_slot['配套 (第2小时)'] = p2_list
                            df_slot['配套 (第3小时)'] = p3_list
                            
                            drops = ['display_items', 'sort_key_subject', '涉及配套']
                            df_slot = df_slot.drop(columns=[c for c in drops if c in df_slot.columns])
                            
                            base_cols = [c for c in df_slot.columns if '配套' not in c]
                            new_cols = ['配套 (第1小时)', '配套 (第2小时)', '配套 (第3小时)']
                            df_slot = df_slot[base_cols + new_cols]
                            
                            df_slot.to_excel(writer, sheet_name='时段总表', index=False)
                            
                            from openpyxl.styles import Alignment, Border, Side
                            
                            ws_slot = writer.sheets['时段总表']
                            col_pkg_start = 5
                            
                            thick_border = Border(bottom=Side(style='thick', color='000000'))
                            thin_border = Border(bottom=Side(style='thin', color='D3D3D3'))
                            center_align = Alignment(horizontal='center', vertical='center')
                            
                            max_row = len(df_slot) + 1
                            slot_merge_start = 2
                            
                            for r_idx in range(2, max_row + 2):
                                cell1 = ws_slot.cell(row=r_idx, column=col_pkg_start)
                                cell2 = ws_slot.cell(row=r_idx, column=col_pkg_start+1)
                                cell3 = ws_slot.cell(row=r_idx, column=col_pkg_start+2)
                                
                                val1, val2, val3 = cell1.value, cell2.value, cell3.value
                                
                                if val1 == val2 == val3 and val1 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start, end_row=r_idx, end_column=col_pkg_start+2)
                                    cell1.alignment = center_align
                                elif val1 == val2 and val1 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start, end_row=r_idx, end_column=col_pkg_start+1)
                                    cell1.alignment = center_align
                                    cell3.alignment = center_align
                                elif val2 == val3 and val2 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start+1, end_row=r_idx, end_column=col_pkg_start+2)
                                    cell2.alignment = center_align
                                    cell1.alignment = center_align
                                else:
                                    cell1.alignment = center_align
                                    cell2.alignment = center_align
                                    cell3.alignment = center_align
                                
                                current_slot = ws_slot.cell(row=r_idx, column=1).value
                                next_slot = None
                                if r_idx < max_row + 1:
                                    next_slot = ws_slot.cell(row=r_idx+1, column=1).value
                                
                                if current_slot != next_slot:
                                    ws_slot.merge_cells(start_row=slot_merge_start, start_column=1, end_row=r_idx, end_column=1)
                                    ws_slot.merge_cells(start_row=slot_merge_start, start_column=2, end_row=r_idx, end_column=2)
                                    
                                    ws_slot.cell(row=slot_merge_start, column=1).alignment = center_align
                                    ws_slot.cell(row=slot_merge_start, column=2).alignment = center_align
                                    
                                    for c_idx in range(1, 8):
                                        cell = ws_slot.cell(row=r_idx, column=c_idx)
                                        cell.border = thick_border
                                    
                                    slot_merge_start = r_idx + 1
                                else:
                                    for c_idx in range(1, 8):
                                        ws_slot.cell(row=r_idx, column=c_idx).border = thin_border
                            
                            df_overview = df_class_export[['科目 & 班级', '学生配套', '人数']].copy()
                            df_overview.columns = ['科目 SUBJECT', '配套 PACKAGE', '人数']
                            df_overview.to_excel(writer, sheet_name='导入', index=False)
                            
                            workbook = writer.book
                            for sheet_name in writer.sheets:
                                worksheet = writer.sheets[sheet_name]
                                if sheet_name == '时段总表':
                                    df_to_measure = df_slot
                                elif sheet_name == '导入':
                                    df_to_measure = df_overview
                                else:
                                    df_to_measure = df_class_export
                                    
                                for idx, col in enumerate(df_to_measure.columns):
                                    max_len = max(
                                        len(str(col)),
                                        df_to_measure[col].astype(str).str.len().max() if not df_to_measure[col].empty else 0
                                    )
                                    adjusted_width = min(max_len + 4, 60)
                                    worksheet.column_dimensions[get_column_letter(idx + 1)].width = adjusted_width
                        
                        st.download_button(
                            label="📥 下载",
                            data=output.getvalue(),
                            file_name=f"{save_name}_排课结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{save_name}"
                        )
                    
                    # 删除按钮
                    with col2:
                        if st.button("🗑️ 删除", key=f"delete_{save_name}"):
                            delete_saved_solution(save_name)
                            st.rerun()
        else:
            st.info("暂无已保存的方案")
    
    # 主界面
    st.markdown("### 📂 上传数据文件")
    
    uploaded_files = st.file_uploader(
        "选择Excel或CSV文件",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        help="上传包含配套、人数、科目信息的文件"
    )
    
    if uploaded_files:
        all_uploaded_files = uploaded_files
        
        st.markdown("---")
        st.markdown("### ⚙️ 求解参数设置")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            num_time_slots = st.number_input("时段组数", min_value=1, max_value=21, value=7, step=1)
            min_class_size = st.number_input("班额下限", min_value=1, max_value=100, value=15, step=1)
        
        with col2:
            max_class_size = st.number_input("班额上限", min_value=1, max_value=100, value=35, step=1)
            time_limit = st.number_input("求解时限(秒)", min_value=10, max_value=600, value=120, step=10)
        
        with col3:
            allow_split = st.checkbox("启用时段分割", value=False, help="允许同一配套在不同时段上同一科目的不同班级")
        
        # 强制开班设置
        st.markdown("#### 🎯 强制开班设置（可选）")
        use_force_open = st.checkbox("启用强制开班约束")
        
        force_open_classes = None
        if use_force_open:
            force_open_input = st.text_area(
                "输入格式：科目名称:班级数，每行一个",
                placeholder="例如：\n会计:2\n经济法:3",
                help="强制指定某些科目必须开设的班级数"
            )
            
            if force_open_input:
                force_open_classes = {}
                for line in force_open_input.strip().split('\n'):
                    if ':' in line:
                        subject, num = line.split(':')
                        force_open_classes[subject.strip()] = int(num.strip())
                
                if force_open_classes:
                    st.success(f"✅ 已设置强制开班：{force_open_classes}")
        
        # 求解按钮
        if st.button("🚀 开始求解", type="primary"):
            all_packages = {}
            all_subject_hours = {}
            file_errors = []
            
            # 解析所有文件
            for uploaded_file in all_uploaded_files:
                st.info(f"📄 正在处理文件：{uploaded_file.name}")
                packages, subject_hours, stats = parse_uploaded_file(uploaded_file)
                
                if packages is None:
                    file_errors.append(uploaded_file.name)
                    continue
                
                # 合并数据
                for pkg_name, pkg_data in packages.items():
                    if pkg_name in all_packages:
                        st.warning(f"⚠️ 配套 '{pkg_name}' 在多个文件中出现，将使用最后一次的数据")
                    all_packages[pkg_name] = pkg_data
                
                for subject, hours in subject_hours.items():
                    if subject in all_subject_hours and all_subject_hours[subject] != hours:
                        st.error(f"❌ 科目 '{subject}' 在不同文件中的课时不一致！")
                        return
                    all_subject_hours[subject] = hours
            
            if file_errors:
                st.error(f"❌ 以下文件解析失败：{', '.join(file_errors)}")
                return
            
            if not all_packages:
                st.error("❌ 没有成功解析任何数据！")
                return
            
            st.success(f"✅ 成功加载 {len(all_packages)} 个配套，{len(all_subject_hours)} 个科目")
            
            # 开始求解
            solutions = []
            
            # 方案A：基础方案
            st.info("🔄 正在求解方案A：基础方案...")
            progress_bar = st.progress(0.0)
            percentage_text = st.empty()
            status_text = st.empty()
            
            percentage_text.markdown("**10%**")
            status_text.markdown("🔍 方案A求解中...")
            
            result_a = solve_schedule(
                all_packages, all_subject_hours, num_time_slots,
                min_class_size, max_class_size, allow_split, force_open_classes, time_limit
            )
            
            if result_a['status'] == 'success':
                slot_schedule = format_slot_schedule(result_a['slot_details'], all_packages, all_subject_hours)
                analysis = analyze_solution(result_a['class_details'])
                
                solutions.append({
                    'name': '方案A：基础方案',
                    'status': 'success',
                    'icon': '✅',
                    'class_details': result_a['class_details'],
                    'slot_schedule': slot_schedule,
                    'analysis': analysis,
                    'solve_time': result_a['solve_time']
                })
                st.success("✅ 方案A求解成功！")
            else:
                solutions.append({
                    'name': '方案A：基础方案',
                    'status': 'failed',
                    'icon': '❌',
                    'solve_time': result_a['solve_time']
                })
                st.error("❌ 方案A无解")
            
            progress_bar.progress(0.33)
            time.sleep(0.5)
            
            # 方案B：优化班额
            percentage_text.markdown("**40%**")
            status_text.markdown("🔍 方案B求解中...")
            
            adjusted_min = max(min_class_size - 5, 1)
            adjusted_max = max_class_size + 5
            
            result_b = solve_schedule(
                all_packages, all_subject_hours, num_time_slots,
                adjusted_min, adjusted_max, allow_split, force_open_classes, time_limit
            )
            
            if result_b['status'] == 'success':
                slot_schedule = format_slot_schedule(result_b['slot_details'], all_packages, all_subject_hours)
                analysis = analyze_solution(result_b['class_details'])
                
                solutions.append({
                    'name': f'方案B：优化班额（{adjusted_min}-{adjusted_max}人）',
                    'status': 'success',
                    'icon': '✅',
                    'class_details': result_b['class_details'],
                    'slot_schedule': slot_schedule,
                    'analysis': analysis,
                    'solve_time': result_b['solve_time']
                })
                st.success("✅ 方案B求解成功！")
            else:
                solutions.append({
                    'name': f'方案B：优化班额（{adjusted_min}-{adjusted_max}人）',
                    'status': 'failed',
                    'icon': '❌',
                    'solve_time': result_b['solve_time']
                })
                st.error("❌ 方案B无解")
            
            progress_bar.progress(0.66)
            time.sleep(0.5)
            
            # 方案C：增加时段
            percentage_text.markdown("**70%**")
            status_text.markdown("🔍 方案C求解中...")
            
            extended_slots = num_time_slots + 2
            
            result_c = solve_schedule(
                all_packages, all_subject_hours, extended_slots,
                min_class_size, max_class_size, allow_split, force_open_classes, time_limit
            )
            
            if result_c['status'] == 'success':
                slot_schedule = format_slot_schedule(result_c['slot_details'], all_packages, all_subject_hours)
                analysis = analyze_solution(result_c['class_details'])
                
                solutions.append({
                    'name': f'方案C：增加时段（{extended_slots}组）',
                    'status': 'success',
                    'icon': '✅',
                    'class_details': result_c['class_details'],
                    'slot_schedule': slot_schedule,
                    'analysis': analysis,
                    'solve_time': result_c['solve_time']
                })
                st.success("✅ 方案C求解成功！")
            else:
                solutions.append({
                    'name': f'方案C：增加时段（{extended_slots}组）',
                    'status': 'failed',
                    'icon': '❌',
                    'solve_time': result_c['solve_time']
                })
                st.error("❌ 方案C无解")
            
            progress_bar.progress(1.0)
            percentage_text.markdown("**100%**")
            status_text.markdown("🎉 **所有方案求解完成！**")
            time.sleep(0.5)
            
            progress_bar.empty()
            status_text.empty()
            percentage_text.empty()
            
            if not solutions:
                st.markdown('<div class="error-box">', unsafe_allow_html=True)
                st.error("❌ 所有方案均无解！")
                st.markdown('</div>', unsafe_allow_html=True)
                return
            
            st.session_state['solutions'] = solutions
            
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.success(f"✅ 成功生成 {len(solutions)} 个方案！")
            st.markdown('</div>', unsafe_allow_html=True)
            save_history_to_disk(solutions)
        
        if 'solutions' in st.session_state:
            st.markdown("---")
            st.markdown('<div class="sub-header">📊 方案对比</div>', unsafe_allow_html=True)
            
            comparison_data = []
            for sol in st.session_state['solutions']:
                if sol['status'] == 'success':
                    analysis = sol['analysis']
                    comparison_data.append({
                        '方案': sol['name'],
                        '开班数': analysis['total_classes'],
                        '平均班额': f"{analysis['avg_size']}人",
                        '班额范围': f"{analysis['min_size']}-{analysis['max_size']}人",
                        '时段分割次数': analysis['split_count'],
                        '求解时间': f"{sol['solve_time']:.1f}秒",
                        '状态': sol['icon']
                    })
                else:
                    comparison_data.append({
                        '方案': sol['name'],
                        '开班数': '-',
                        '平均班额': '-',
                        '班额范围': '-',
                        '时段分割次数': '-',
                        '求解时间': f"{sol['solve_time']:.1f}秒",
                        '状态': sol['icon']
                    })
            
            df_comparison = pd.DataFrame(comparison_data)
            st.dataframe(df_comparison, use_container_width=True)
            
            for sol in st.session_state['solutions']:
                if sol['status'] == 'failed':
                    continue
                    
                with st.expander(f"📋 {sol['name']} - 详细结果"):
                    st.markdown("---")
                    
                    tab1, tab2, tab3 = st.tabs(["开班详情", "时段总表", "数据导出"])
                    
                    with tab1:
                        df_class = pd.DataFrame(sol['class_details'])
                        st.dataframe(df_class, use_container_width=True)
                    
                    with tab2:
                        st.markdown("### 🕐 时段总表")
                        schedule_data = sol['slot_schedule']
                        
                        if not schedule_data:
                            st.info("暂无数据")
                        else:
                            # 简化显示
                            df_slot = pd.DataFrame(schedule_data)
                            cols_to_drop = ['display_items', 'sort_key_subject']
                            df_slot_export = df_slot.drop(columns=[c for c in cols_to_drop if c in df_slot.columns])
                            st.dataframe(df_slot_export, use_container_width=True)
                    
                    with tab3:
                        st.markdown("### 📥 导出选项")
                        
                        # 生成Excel文件
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            raw_class_data = sol['class_details']
                            raw_slot_data = sol['slot_schedule']
                            
                            df_class = pd.DataFrame(raw_class_data)
                            
                            def format_subject_class_col(row):
                                suffix = row['班级'].replace('班', '')
                                if suffix:
                                    return f"{row['科目']} {suffix}"
                                else:
                                    return row['科目']
                            
                            df_class = df_class.sort_values(by=['科目', '班级'])
                            df_class['科目 & 班级'] = df_class.apply(format_subject_class_col, axis=1)
                            df_class_export = df_class[['科目 & 班级', '人数', '时段', '学生配套']]
                            df_class_export.to_excel(writer, sheet_name='开班详情', index=False)
                            
                            df_slot = pd.DataFrame(raw_slot_data)
                            p1_list, p2_list, p3_list = [], [], []
                            
                            for item in raw_slot_data:
                                current_pkg_slots = ["-", "-", "-"]
                                d_items = item.get('display_items', [])
                                
                                if isinstance(d_items, list):
                                    for sub_item in d_items:
                                        pkg_str = sub_item.get('packages_str', '-')
                                        if not pkg_str or sub_item.get('is_gap', False):
                                            pkg_str = "-"
                                        
                                        rel_slots = sub_item.get('relative_slots', [])
                                        if not rel_slots and 'start_offset' in sub_item:
                                            try: dur = int(sub_item['duration'].replace('h',''))
                                            except: dur = 1
                                            start = sub_item['start_offset']
                                            rel_slots = range(start, start + dur)
                                            
                                        for idx in rel_slots:
                                            if 0 <= idx < 3:
                                                current_pkg_slots[idx] = pkg_str
                                
                                p1_list.append(current_pkg_slots[0])
                                p2_list.append(current_pkg_slots[1])
                                p3_list.append(current_pkg_slots[2])
                            
                            df_slot['配套 (第1小时)'] = p1_list
                            df_slot['配套 (第2小时)'] = p2_list
                            df_slot['配套 (第3小时)'] = p3_list
                            
                            drops = ['display_items', 'sort_key_subject', '涉及配套']
                            df_slot = df_slot.drop(columns=[c for c in drops if c in df_slot.columns])
                            
                            base_cols = [c for c in df_slot.columns if '配套' not in c]
                            new_cols = ['配套 (第1小时)', '配套 (第2小时)', '配套 (第3小时)']
                            df_slot = df_slot[base_cols + new_cols]
                            
                            df_slot.to_excel(writer, sheet_name='时段总表', index=False)
                            
                            from openpyxl.styles import Alignment, Border, Side
                            
                            ws_slot = writer.sheets['时段总表']
                            col_pkg_start = 5
                            
                            thick_border = Border(bottom=Side(style='thick', color='000000'))
                            thin_border = Border(bottom=Side(style='thin', color='D3D3D3'))
                            center_align = Alignment(horizontal='center', vertical='center')
                            
                            max_row = len(df_slot) + 1
                            slot_merge_start = 2
                            
                            for r_idx in range(2, max_row + 2):
                                cell1 = ws_slot.cell(row=r_idx, column=col_pkg_start)
                                cell2 = ws_slot.cell(row=r_idx, column=col_pkg_start+1)
                                cell3 = ws_slot.cell(row=r_idx, column=col_pkg_start+2)
                                
                                val1, val2, val3 = cell1.value, cell2.value, cell3.value
                                
                                if val1 == val2 == val3 and val1 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start, end_row=r_idx, end_column=col_pkg_start+2)
                                    cell1.alignment = center_align
                                elif val1 == val2 and val1 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start, end_row=r_idx, end_column=col_pkg_start+1)
                                    cell1.alignment = center_align
                                    cell3.alignment = center_align
                                elif val2 == val3 and val2 != '-':
                                    ws_slot.merge_cells(start_row=r_idx, start_column=col_pkg_start+1, end_row=r_idx, end_column=col_pkg_start+2)
                                    cell2.alignment = center_align
                                    cell1.alignment = center_align
                                else:
                                    cell1.alignment = center_align
                                    cell2.alignment = center_align
                                    cell3.alignment = center_align
                                
                                current_slot = ws_slot.cell(row=r_idx, column=1).value
                                next_slot = None
                                if r_idx < max_row + 1:
                                    next_slot = ws_slot.cell(row=r_idx+1, column=1).value
                                
                                if current_slot != next_slot:
                                    ws_slot.merge_cells(start_row=slot_merge_start, start_column=1, end_row=r_idx, end_column=1)
                                    ws_slot.merge_cells(start_row=slot_merge_start, start_column=2, end_row=r_idx, end_column=2)
                                    
                                    ws_slot.cell(row=slot_merge_start, column=1).alignment = center_align
                                    ws_slot.cell(row=slot_merge_start, column=2).alignment = center_align
                                    
                                    for c_idx in range(1, 8):
                                        cell = ws_slot.cell(row=r_idx, column=c_idx)
                                        cell.border = thick_border
                                    
                                    slot_merge_start = r_idx + 1
                                else:
                                    for c_idx in range(1, 8):
                                        ws_slot.cell(row=r_idx, column=c_idx).border = thin_border
                            
                            df_overview = df_class_export[['科目 & 班级', '学生配套', '人数']].copy()
                            df_overview.columns = ['科目 SUBJECT', '配套 PACKAGE', '人数']
                            df_overview.to_excel(writer, sheet_name='导入', index=False)
                            
                            workbook = writer.book
                            for sheet_name in writer.sheets:
                                worksheet = writer.sheets[sheet_name]
                                if sheet_name == '时段总表':
                                    df_to_measure = df_slot
                                elif sheet_name == '导入':
                                    df_to_measure = df_overview
                                else:
                                    df_to_measure = df_class_export
                                    
                                for idx, col in enumerate(df_to_measure.columns):
                                    max_len = max(
                                        len(str(col)),
                                        df_to_measure[col].astype(str).str.len().max() if not df_to_measure[col].empty else 0
                                    )
                                    adjusted_width = min(max_len + 4, 60)
                                    worksheet.column_dimensions[get_column_letter(idx + 1)].width = adjusted_width
                        
                        # 下载按钮
                        st.download_button(
                            label="📥 下载Excel文件",
                            data=output.getvalue(),
                            file_name=f"{sol['name'].replace('：', '_')}_排课结果.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_main_{sol['name']}"
                        )
                        
                        st.markdown("---")
                        
                        # 保存到存储区域
                        st.markdown('<div class="save-section">', unsafe_allow_html=True)
                        st.markdown("#### 💾 保存方案到存储")
                        
                        col1, col2 = st.columns([3, 1])
                        
                        with col1:
                            save_name = st.text_input(
                                "输入存储名称",
                                placeholder="例如：2024秋季排课_最终版",
                                key=f"save_name_{sol['name']}"
                            )
                        
                        with col2:
                            st.markdown("<br>", unsafe_allow_html=True)  # 对齐按钮
                            if st.button("💾 保存方案", key=f"save_btn_{sol['name']}"):
                                if save_name:
                                    if save_name in st.session_state['saved_solutions']:
                                        st.warning(f"⚠️ 名称 '{save_name}' 已存在，是否覆盖？")
                                        if st.button("确认覆盖", key=f"confirm_{sol['name']}"):
                                            save_solution_to_storage(sol, save_name)
                                            st.success(f"✅ 方案已保存为：{save_name}")
                                            st.rerun()
                                    else:
                                        save_solution_to_storage(sol, save_name)
                                        st.success(f"✅ 方案已保存为：{save_name}")
                                        time.sleep(1)
                                        st.rerun()
                                else:
                                    st.error("❌ 请输入存储名称！")
                        
                        st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
