#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
排课求解器 Web UI (升级版)
功能更新：
1. 班级按人数自动命名为 A, B, C...
2. 时段总表自动合并同一拨学生的分割课程
"""

import streamlit as st
import pandas as pd
import re
from pathlib import Path
import io
import time
from ortools.sat.python import cp_model
from collections import defaultdict
from openpyxl.utils import get_column_letter

# 页面配置
st.set_page_config(
    page_title="智能排课求解器",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== 全局样式 ==========
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
</style>
""", unsafe_allow_html=True)

# ========== 工具函数 ==========
def natural_sort_key(s):
    """自然排序的key函数"""
    import re
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split(r'(\d+)', str(s))]

def parse_subject_string(subject_str):
    """解析科目字符串"""
    subjects = {}
    pattern = r'([^,\(（]+)[\(（](\d+)[\)）]'
    matches = re.findall(pattern, subject_str)
    for subject, hours in matches:
        subject = subject.strip()
        subjects[subject] = int(hours)
    return subjects

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
                except Exception as e:
                    last_error = e
                    continue
            if df is None:
                raise Exception(f"无法识别文件编码。最后错误：{last_error}")
        
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
                    return None, None, None
        
        min_hours = min(s['总课时'] for s in total_hours_stats)
        max_hours = max(s['总课时'] for s in total_hours_stats)
        
        if min_hours < 21:
            st.info(f"ℹ️ 检测到部分配套总课时少于21小时（范围：{min_hours}-{max_hours}小时）")
        
        return packages, subject_hours, max_hours
    
    except Exception as e:
        st.error(f"❌ 文件解析失败: {str(e)}")
        return None, None, None

def calculate_subject_enrollment(packages):
    enrollment = defaultdict(int)
    for p_data in packages.values():
        for subject in p_data['科目'].keys():
            enrollment[subject] += p_data['人数']
    return dict(enrollment)

def calculate_recommended_slots(max_total_hours):
    import math
    if max_total_hours <= 3:
        return 1
    recommended = math.ceil((max_total_hours - 1) / 2)
    return max(2, min(recommended, 20))

# ========== 排课求解器核心 ==========
class ScheduleSolver:
    def __init__(self, packages, subject_hours, config):
        self.packages = packages
        self.subject_hours = subject_hours
        self.config = config
        self.subjects = list(subject_hours.keys())
        self.package_names = list(packages.keys())
        
        # 时段定义
        self.TIME_SLOTS_1H = list(range(1, config['num_slots'] * 2 + 2))
        self.SLOT_GROUPS = {}
        for i in range(1, config['num_slots'] + 1):
            if i < config['num_slots']:
                self.SLOT_GROUPS[f'S{i}'] = [i*2-1, i*2]
            else:
                self.SLOT_GROUPS[f'S{i}'] = [i*2-1, i*2, i*2+1]
        
        self.SLOT_TO_GROUP = {}
        for group_name, slots in self.SLOT_GROUPS.items():
            for slot in slots:
                self.SLOT_TO_GROUP[slot] = group_name
        
        self.subject_enrollment = calculate_subject_enrollment(packages)
    
    def build_model(self, objective_type='min_classes'):
        model = cp_model.CpModel()
        
        # 变量定义
        u_r = {}
        y_rt = {}
        u_pkr = {}
        x_prt = {}
        
        for k in self.subjects:
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                u_r[(k, r)] = model.NewBoolVar(f'u_{k}_{r}')
                for t in self.TIME_SLOTS_1H:
                    y_rt[(k, r, t)] = model.NewBoolVar(f'y_{k}_{r}_{t}')
        
        for p in self.package_names:
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    u_pkr[(p, k, r)] = model.NewBoolVar(f'u_{p}_{k}_{r}')
                    for t in self.TIME_SLOTS_1H:
                        x_prt[(p, k, r, t)] = model.NewBoolVar(f'x_{p}_{k}_{r}_{t}')
        
        # 约束定义
        # HA: 精确学时
        for k in self.subjects:
            H_k = self.subject_hours[k]
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                model.Add(sum(y_rt[(k, r, t)] for t in self.TIME_SLOTS_1H) == H_k).OnlyEnforceIf(u_r[(k, r)])
                model.Add(sum(y_rt[(k, r, t)] for t in self.TIME_SLOTS_1H) == 0).OnlyEnforceIf(u_r[(k, r)].Not())
        
        # HB: 同师全修
        for p in self.package_names:
            for k in self.subjects:
                if k in self.packages[p]['科目']:
                    model.Add(sum(u_pkr[(p, k, r)] for r in range(1, self.config['max_classes_per_subject'] + 1)) == 1)
                else:
                    for r in range(1, self.config['max_classes_per_subject'] + 1):
                        model.Add(u_pkr[(p, k, r)] == 0)
        
        # HC: 班额限制
        for k in self.subjects:
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                class_size = sum(self.packages[p]['人数'] * u_pkr[(p, k, r)] for p in self.package_names)
                model.Add(class_size >= self.config['min_class_size']).OnlyEnforceIf(u_r[(k, r)])
                model.Add(class_size <= self.config['max_class_size']).OnlyEnforceIf(u_r[(k, r)])
                model.Add(class_size == 0).OnlyEnforceIf(u_r[(k, r)].Not())
        
        # H2 & H2': 逻辑关联与互斥
        for p in self.package_names:
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    for t in self.TIME_SLOTS_1H:
                        model.Add(x_prt[(p, k, r, t)] <= u_pkr[(p, k, r)])
                        model.Add(x_prt[(p, k, r, t)] <= y_rt[(k, r, t)])
                        model.Add(x_prt[(p, k, r, t)] >= u_pkr[(p, k, r)] + y_rt[(k, r, t)] - 1)
        
        for p in self.package_names:
            for t in self.TIME_SLOTS_1H:
                model.Add(sum(x_prt[(p, k, r, t)] for k in self.subjects for r in range(1, self.config['max_classes_per_subject'] + 1)) <= 1)
        
        # H6: 教师资源约束
        for k in self.subjects:
            for t in self.TIME_SLOTS_1H:
                model.Add(sum(y_rt[(k, r, t)] for r in range(1, self.config['max_classes_per_subject'] + 1)) <= 1)
        
        # H1: 覆盖需求
        for p in self.package_names:
            for k in self.subjects:
                if k in self.packages[p]['科目']:
                    required_hours = self.packages[p]['科目'][k]
                    total_hours_pk = sum(x_prt[(p, k, r, t)] for r in range(1, self.config['max_classes_per_subject'] + 1) for t in self.TIME_SLOTS_1H)
                    model.Add(total_hours_pk == required_hours)
        
        # 开班限制
        for k in self.subjects:
            model.Add(sum(u_r[(k, r)] for r in range(1, self.config['max_classes_per_subject'] + 1)) <= self.config['max_classes_per_subject'])
        
        for k, count in self.config['forced_class_count'].items():
            if k in self.subjects:
                model.Add(sum(u_r[(k, r)] for r in range(1, self.config['max_classes_per_subject'] + 1)) == count)
        
        # 时段分割处理
        slot_split_penalty = 0
        if not self.config['allow_slot_split']:
            for p in self.package_names:
                for group_name, group_slots in self.SLOT_GROUPS.items():
                    subjects_in_group = []
                    for k in self.subjects:
                        for r in range(1, self.config['max_classes_per_subject'] + 1):
                            has_subject = model.NewBoolVar(f'has_{p}_{k}_{r}_{group_name}')
                            model.AddMaxEquality(has_subject, [x_prt[(p, k, r, t)] for t in group_slots])
                            subjects_in_group.append(has_subject)
                    model.Add(sum(subjects_in_group) <= 1)
        else:
            split_vars = []
            for p in self.package_names:
                for group_name, group_slots in self.SLOT_GROUPS.items():
                    subjects_in_group = []
                    for k in self.subjects:
                        for r in range(1, self.config['max_classes_per_subject'] + 1):
                            has_subject = model.NewBoolVar(f'has_{p}_{k}_{r}_{group_name}')
                            model.AddMaxEquality(has_subject, [x_prt[(p, k, r, t)] for t in group_slots])
                            subjects_in_group.append(has_subject)
                    
                    num_subjects = sum(subjects_in_group)
                    is_split = model.NewBoolVar(f'split_{p}_{group_name}')
                    model.Add(num_subjects >= 2).OnlyEnforceIf(is_split)
                    model.Add(num_subjects <= 1).OnlyEnforceIf(is_split.Not())
                    split_vars.append(is_split)
            slot_split_penalty = sum(split_vars) * self.config['slot_split_penalty']
        
        # 目标函数
        total_classes = sum(u_r[(k, r)] for k in self.subjects for r in range(1, self.config['max_classes_per_subject'] + 1))
        priority_penalty = sum(u_r[(k, r)] * r * max(0, 100 - self.subject_enrollment[k]) for k in self.subjects for r in range(1, self.config['max_classes_per_subject'] + 1))
        
        if objective_type == 'min_classes':
            model.Minimize(total_classes * 100000 + slot_split_penalty + priority_penalty)
        elif objective_type == 'balanced':
            effective_sizes_for_max = []
            effective_sizes_for_min = []
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    actual_size = sum(self.packages[p]['人数'] * u_pkr[(p, k, r)] for p in self.package_names)
                    eff_size_max = model.NewIntVar(0, 200, f'eff_max_{k}_{r}')
                    model.Add(eff_size_max == actual_size).OnlyEnforceIf(u_r[(k, r)])
                    model.Add(eff_size_max == 0).OnlyEnforceIf(u_r[(k, r)].Not())
                    effective_sizes_for_max.append(eff_size_max)
                    
                    eff_size_min = model.NewIntVar(0, 200, f'eff_min_{k}_{r}')
                    model.Add(eff_size_min == actual_size).OnlyEnforceIf(u_r[(k, r)])
                    model.Add(eff_size_min == 200).OnlyEnforceIf(u_r[(k, r)].Not())
                    effective_sizes_for_min.append(eff_size_min)
            
            max_size = model.NewIntVar(0, 200, 'max_size')
            min_size = model.NewIntVar(0, 200, 'min_size')
            model.AddMaxEquality(max_size, effective_sizes_for_max)
            model.AddMinEquality(min_size, effective_sizes_for_min)
            model.Minimize(total_classes * 1000000 + slot_split_penalty * 100 + (max_size - min_size) * 1000 + priority_penalty)
        
        return model, {'u_r': u_r, 'y_rt': y_rt, 'u_pkr': u_pkr, 'x_prt': x_prt}
    
    def solve(self, model, variables, timeout):
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = timeout
        solver.parameters.log_search_progress = False
        solver.parameters.num_search_workers = 8
        
        start_time = time.time()
        status = solver.Solve(model)
        solve_time = time.time() - start_time
        
        status_map = {
            cp_model.OPTIMAL: ('最优解', '✅'),
            cp_model.FEASIBLE: ('可行解', '✅'),
            cp_model.INFEASIBLE: ('无解', '❌'),
            cp_model.MODEL_INVALID: ('模型无效', '⚠️'),
            cp_model.UNKNOWN: ('超时/未知', '⏱️')
        }
        status_name, icon = status_map.get(status, ('未知状态', '❓'))
        
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            return {'status': 'success', 'solver': solver, 'variables': variables, 'solve_status': status_name, 'icon': icon, 'solve_time': solve_time}
        else:
            return {'status': 'failed', 'solve_status': status_name, 'icon': icon, 'solve_time': solve_time}
    
    def analyze_solution(self, result):
        solver = result['solver']
        u_r = result['variables']['u_r']
        u_pkr = result['variables']['u_pkr']
        x_prt = result['variables']['x_prt']
        
        total_classes = sum(1 for (k, r) in u_r if solver.Value(u_r[(k, r)]) == 1)
        class_sizes = []
        for k in self.subjects:
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                if solver.Value(u_r[(k, r)]) == 1:
                    size = sum(self.packages[p]['人数'] for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1)
                    class_sizes.append(size)
        
        split_count = 0
        split_details = []
        for p in self.package_names:
            for group_name, group_slots in self.SLOT_GROUPS.items():
                subjects_in_group = set()
                for t in group_slots:
                    for k in self.subjects:
                        for r in range(1, self.config['max_classes_per_subject'] + 1):
                            if solver.Value(x_prt[(p, k, r, t)]) == 1:
                                subjects_in_group.add(k)
                if len(subjects_in_group) >= 2:
                    split_count += 1
                    split_details.append(f"{p}-{group_name}: {', '.join(sorted(subjects_in_group))}")
        
        return {
            'total_classes': total_classes,
            'avg_size': round(sum(class_sizes) / len(class_sizes), 1) if class_sizes else 0,
            'min_size': min(class_sizes) if class_sizes else 0,
            'max_size': max(class_sizes) if class_sizes else 0,
            'split_count': split_count,
            'split_details': split_details
        }

    # ================= 修改后的核心提取函数 =================
    def extract_timetable(self, result):
        """
        提取课表数据
        包含功能：
        1. 班级自动命名为 A, B, C... (按人数大小)
        2. 时段总表自动合并同一拨学生的分割课程 (如: 化学(1h)+商业(1h))
        """
        solver = result['solver']
        u_r = result['variables']['u_r']
        y_rt = result['variables']['y_rt']
        u_pkr = result['variables']['u_pkr']
        
        # --- 步骤 1: 预计算班级大小并分配名称 (A, B, C...) ---
        class_name_mapping = {} # {科目: {原ID_r: '班A'}}
        
        for k in self.subjects:
            active_classes = []
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                if solver.Value(u_r[(k, r)]) == 1:
                    students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                    size = sum(self.packages[p]['人数'] for p in students)
                    active_classes.append({'r': r, 'size': size})
            
            # 按人数降序排列 (人数多的叫班A)
            active_classes.sort(key=lambda x: x['size'], reverse=True)
            
            mapping = {}
            for idx, item in enumerate(active_classes):
                new_name = f"班{chr(65 + idx)}" # 班A, 班B...
                mapping[item['r']] = new_name
            class_name_mapping[k] = mapping

        # --- 步骤 2: 生成开班详情 (按科目列表) ---
        class_details = []
        for k in self.subjects:
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                if solver.Value(u_r[(k, r)]) == 1:
                    students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                    size = sum(self.packages[p]['人数'] for p in students)
                    
                    time_slots = [t for t in self.TIME_SLOTS_1H if solver.Value(y_rt[(k, r, t)]) == 1]
                    slot_groups_used = defaultdict(list)
                    for t in time_slots:
                        group = self.SLOT_TO_GROUP[t]
                        slot_groups_used[group].append(t)
                    
                    slot_str = ', '.join([f"{g}({len(slots)}h)" 
                                         for g, slots in sorted(slot_groups_used.items(), key=lambda x: natural_sort_key(x[0]))])
                    students_sorted = sorted(students, key=natural_sort_key)
                    class_name = class_name_mapping[k].get(r, f"班{r}")
                    
                    class_details.append({
                        '科目': k,
                        '班级': class_name,
                        '人数': size,
                        '时段': slot_str,
                        '学生配套': ', '.join(students_sorted)
                    })
        class_details.sort(key=lambda x: x['科目'])

        # --- 步骤 3: 生成时段总表 (合并分割课程逻辑) ---
        slot_schedule_data = []
        
        # 遍历每个时段组 (S1, S2...)
        for group_name in sorted(self.SLOT_GROUPS.keys(), key=natural_sort_key):
            group_slots = self.SLOT_GROUPS[group_name]
            
            # 临时存储桶：key=学生配套集合(frozenset), value=该时段内的课程列表
            # 目的是把同一拨学生在同一时段上的不同课归类到一起
            student_group_batches = defaultdict(list)
            
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    # 检查该班级在这个时段组内是否有课
                    active_sub_slots = [t for t in group_slots if solver.Value(y_rt[(k, r, t)]) == 1]
                    
                    if active_sub_slots:
                        # 获取上这门课的学生配套
                        students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                        if not students: continue
                        
                        size = sum(self.packages[p]['人数'] for p in students)
                        class_name = class_name_mapping[k].get(r, f"班{r}")
                        
                        # 生成唯一Key：根据学生配套名单 (排序后转元组，保证唯一性)
                        students_key = tuple(sorted(students))
                        
                        student_group_batches[students_key].append({
                            'subject': k,
                            'class_name': class_name,
                            'duration': len(active_sub_slots),
                            'students': students,
                            'size': size,
                            'first_slot': min(active_sub_slots) # 用于内部排序，先上的课排前面
                        })
            
            # 处理聚合后的数据，生成表格行
            for students_tuple, class_list in student_group_batches.items():
                # 按实际上课时间排序 (例如先上化学再上商业)
                class_list.sort(key=lambda x: x['first_slot'])
                
                students_str = ', '.join(sorted(list(students_tuple), key=natural_sort_key))
                total_size = class_list[0]['size'] # 同一拨学生，人数是一样的
                
                if len(class_list) == 1:
                    # 情况A: 没有分割，只有一门课
                    item = class_list[0]
                    slot_schedule_data.append({
                        '时段': group_name,
                        '时长': f"{item['duration']}h",
                        '科目': item['subject'],
                        '班级': item['class_name'],
                        '人数': total_size,
                        '涉及配套': students_str
                    })
                else:
                    # 情况B: 出现分割，合并显示！
                    # 格式: 化学(1h) + 商业(1h)
                    combined_subject = " + ".join([f"{item['subject']}({item['duration']}h)" for item in class_list])
                    # 格式: 班A + 班B
                    combined_class = " + ".join([item['class_name'] for item in class_list])
                    # 总时长
                    total_duration = sum(item['duration'] for item in class_list)
                    
                    slot_schedule_data.append({
                        '时段': group_name,
                        '时长': f"{total_duration}h", # 显示总时长
                        '科目': combined_subject,     # 合并后的科目名
                        '班级': combined_class,       # 合并后的班级名
                        '人数': total_size,
                        '涉及配套': students_str
                    })
        
        return class_details, slot_schedule_data
    # ================= 结束修改 =================

# ========== 主应用 ==========
def main():
    st.markdown('<div class="main-header">📚 智能排课求解器 (升级版)</div>', unsafe_allow_html=True)
    
    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        st.subheader("📁 数据导入")
        
        # 下载模板功能
        st.markdown("##### 📥 下载数据模板")
        template_data = """配套,科目,人数,总学点
P12,"会计学（4）,经济（4）,商业（3）,历史（4）,AI应用（2）,AI编程（2）",5,19
P13,"物理（6）,经济（4）,历史（4）,地理（4）,AI应用（2）",6,20
P14,"物理（6）,会计学（4）,经济（4）,商业（3）,AI应用（2）,AI编程（2）",4,21"""
        
        col1, col2 = st.columns([1, 1])
        with col1:
            st.download_button("📄 CSV模板", template_data.encode('utf-8-sig'), "排课数据模板.csv", "text/csv")
        with col2:
            template_df = pd.DataFrame([
                {'配套': 'P12', '科目': '会计学（4）,经济（4）,商业（3）,历史（4）,AI应用（2）,AI编程（2）', '人数': 5},
                {'配套': 'P13', '科目': '物理（6）,经济（4）,历史（4）,地理（4）,AI应用（2）', '人数': 6},
            ])
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False)
            st.download_button("📊 Excel模板", excel_buffer.getvalue(), "排课数据模板.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.markdown("---")
        
        # 文件上传
        uploaded_file = st.file_uploader("选择文件", type=['xlsx', 'xls', 'csv'], label_visibility="collapsed")
        
        if uploaded_file:
            with st.spinner("正在解析文件..."):
                packages, subject_hours, max_hours = parse_uploaded_file(uploaded_file)
            
            if packages and subject_hours:
                st.success(f"✅ 成功加载 {len(packages)} 个配套，{len(subject_hours)} 个科目")
                st.session_state['packages'] = packages
                st.session_state['subject_hours'] = subject_hours
                st.session_state['max_total_hours'] = max_hours
        
        st.markdown("---")
        
        st.subheader("🔧 求解参数")
        
        min_class_size = st.number_input("最小班额", min_value=1, max_value=100, value=5, step=1)
        max_class_size = st.number_input("最大班额", min_value=1, max_value=200, value=60, step=1)
        max_classes_per_subject = st.number_input("每科目最大班数", min_value=1, max_value=10, value=3, step=1)
        
        if 'max_total_hours' in st.session_state:
            max_hours = st.session_state['max_total_hours']
            recommended_slots = calculate_recommended_slots(max_hours)
            default_slots = recommended_slots
        else:
            default_slots = 10
        
        num_slots = st.number_input("时段组数量", min_value=1, max_value=20, value=default_slots, step=1)
        solver_timeout = st.number_input("求解超时(秒)", min_value=10, max_value=600, value=120, step=10)
        
        st.markdown("---")
        
        st.subheader("🔀 时段分割")
        allow_slot_split = st.checkbox("允许时段分割", value=True)
        if allow_slot_split:
            slot_split_penalty = st.slider("分割惩罚系数", 0, 5000, 1000, 100)
        else:
            slot_split_penalty = 0
        
        st.markdown("---")
        
        st.subheader("🔒 强制开班")
        forced_class_count = {}
        if 'subject_hours' in st.session_state:
            for subject in st.session_state['subject_hours'].keys():
                count = st.number_input(f"{subject}", min_value=0, max_value=10, value=0, key=f"forced_{subject}")
                if count > 0:
                    forced_class_count[subject] = count
    
    # 主内容区
    if 'packages' not in st.session_state:
        st.markdown('<div class="info-box"><h3>👋 欢迎使用智能排课系统</h3>请在左侧上传数据文件开始使用。</div>', unsafe_allow_html=True)
        return
    
    st.markdown('<div class="sub-header">📊 数据概览</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1: st.metric("配套数量", len(st.session_state['packages']))
    with col2: st.metric("科目数量", len(st.session_state['subject_hours']))
    with col3: st.metric("学生总数", sum(p['人数'] for p in st.session_state['packages'].values()))
    
    st.markdown("---")
    st.markdown('<div class="sub-header">🚀 开始求解</div>', unsafe_allow_html=True)
    
    if st.button("🎯 生成排课方案", type="primary", use_container_width=True):
        config = {
            'min_class_size': min_class_size,
            'max_class_size': max_class_size,
            'max_classes_per_subject': max_classes_per_subject,
            'num_slots': num_slots,
            'allow_slot_split': allow_slot_split,
            'slot_split_penalty': slot_split_penalty,
            'forced_class_count': forced_class_count
        }
        
        solver_instance = ScheduleSolver(st.session_state['packages'], st.session_state['subject_hours'], config)
        solution_configs = [{'type': 'min_classes', 'name': '方案A：最少开班'}, {'type': 'balanced', 'name': '方案B：均衡班额'}]
        
        progress_bar = st.progress(0)
        solutions = []
        
        for i, sol_config in enumerate(solution_configs):
            progress_bar.progress((i + 1) / len(solution_configs))
            model, variables = solver_instance.build_model(sol_config['type'])
            result = solver_instance.solve(model, variables, solver_timeout)
            
            if result['status'] == 'success':
                result['name'] = sol_config['name']
                result['analysis'] = solver_instance.analyze_solution(result)
                result['class_details'], result['slot_schedule'] = solver_instance.extract_timetable(result)
                solutions.append(result)
        
        progress_bar.empty()
        st.session_state['solutions'] = solutions
        
        if not solutions:
            st.error("❌ 所有方案均无解！请尝试增加时段数或放宽班额限制。")
        else:
            st.success(f"✅ 成功生成 {len(solutions)} 个方案！")
    
    if 'solutions' in st.session_state:
        st.markdown("---")
        st.markdown('<div class="sub-header">📊 方案对比</div>', unsafe_allow_html=True)
        
        comparison_data = []
        for sol in st.session_state['solutions']:
            a = sol['analysis']
            comparison_data.append({
                '方案': sol['name'],
                '开班数': a['total_classes'],
                '平均班额': f"{a['avg_size']}人",
                '班额范围': f"{a['min_size']}-{a['max_size']}人",
                '分割次数': a['split_count'],
                '状态': sol['icon']
            })
        st.dataframe(pd.DataFrame(comparison_data), use_container_width=True)
        
        for sol in st.session_state['solutions']:
            with st.expander(f"📋 {sol['name']} - 详细结果"):
                tab1, tab2, tab3 = st.tabs(["开班详情", "时段总表", "数据导出"])
                
                with tab1:
                    st.dataframe(pd.DataFrame(sol['class_details']), use_container_width=True)
                
                with tab2:
                    st.markdown("**说明：** 如果某时段显示为 `科目A(1h) + 科目B(1h)`，表示该配套在该时段先后上这两门课。")
                    st.dataframe(pd.DataFrame(sol['slot_schedule']), use_container_width=True, height=600)
                
                with tab3:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_class = pd.DataFrame(sol['class_details'])
                        df_slot = pd.DataFrame(sol['slot_schedule'])
                        df_class.to_excel(writer, sheet_name='开班详情', index=False)
                        df_slot.to_excel(writer, sheet_name='时段总表', index=False)
                        
                        # 调整列宽
                        for sheet_name in ['开班详情', '时段总表']:
                            ws = writer.sheets[sheet_name]
                            df = df_class if sheet_name == '开班详情' else df_slot
                            for idx, col in enumerate(df.columns):
                                max_len = max(len(str(col)), df[col].astype(str).str.len().max())
                                ws.column_dimensions[get_column_letter(idx + 1)].width = min(max_len + 2, 50)
                                
                    st.download_button("📥 下载Excel文件", output.getvalue(), f"{sol['name']}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
