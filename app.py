#!/usr/bin/env python3
# -*- coding: utf-8 -*-
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
    """自然排序的key函数，用于正确排序包含数字的字符串
    例如: S1, S2, S3, ..., S9, S10, S11 (而不是 S1, S10, S11, S2)
    """
    import re
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split(r'(\d+)', str(s))]

def parse_subject_string(subject_str):
    """解析科目字符串（支持中英文括号）
    输入: "会计(6),历史(4),地理(4),商业(3)" 或 "会计（6）,历史（4）"
    输出: {'会计': 6, '历史': 4, '地理': 4, '商业': 3}
    """
    subjects = {}
    # 匹配格式：科目名(数字) 或 科目名（数字）
    # 同时支持英文括号() 和中文括号（）
    pattern = r'([^,\(（]+)[\(（](\d+)[\)）]'
    matches = re.findall(pattern, subject_str)
    for subject, hours in matches:
        subject = subject.strip()
        subjects[subject] = int(hours)
    return subjects

def parse_uploaded_file(uploaded_file):
    """解析上传的Excel/CSV文件"""
    try:
        # 尝试读取Excel
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file)
        else:
            # 尝试多种编码方式读取CSV
            encodings = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'cp936', 'utf-8-sig']
            df = None
            last_error = None
            
            for encoding in encodings:
                try:
                    uploaded_file.seek(0)  # 重置文件指针
                    df = pd.read_csv(uploaded_file, encoding=encoding)
                    st.success(f"✅ 成功读取文件（编码：{encoding}）")
                    break
                except (UnicodeDecodeError, Exception) as e:
                    last_error = e
                    continue
            
            if df is None:
                raise Exception(f"无法识别文件编码，请确保文件是有效的CSV格式。最后错误：{last_error}")
        
        # 解析数据
        packages = {}
        subject_hours = {}
        total_hours_stats = []
        
        for _, row in df.iterrows():
            package_name = str(row['配套']).strip()
            student_count = int(row['人数'])
            subject_str = str(row['科目'])
            
            # 解析科目字符串
            subjects = parse_subject_string(subject_str)
            
            # 计算该配套的总课时
            total_hours = sum(subjects.values())
            total_hours_stats.append({
                '配套': package_name,
                '总课时': total_hours
            })
            
            packages[package_name] = {
                '人数': student_count,
                '科目': subjects
            }
            
            # 收集所有科目的课时（强制要求一致性）
            for subject, hours in subjects.items():
                if subject not in subject_hours:
                    subject_hours[subject] = hours
                elif subject_hours[subject] != hours:
                    # 严重错误：课时不一致会导致约束冲突
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
                    - 将4小时的会计命名为"会计基础"
                    - 将6小时的会计命名为"会计进阶"
                    - 这样系统会将它们视为不同科目
                    """)
                    return None, None, None
        
        # 显示总课时统计
        min_hours = min(s['总课时'] for s in total_hours_stats)
        max_hours = max(s['总课时'] for s in total_hours_stats)
        
        if min_hours < 21:
            st.info(f"ℹ️ 检测到部分配套总课时少于21小时（范围：{min_hours}-{max_hours}小时）")
            st.success("✅ 系统支持总课时不足的配套，这些配套将在某些时段不上课")
            
            # 显示总课时不足的配套
            short_packages = [s for s in total_hours_stats if s['总课时'] < 21]
            if short_packages:
                with st.expander("查看总课时不足21的配套"):
                    for pkg in short_packages:
                        st.text(f"  {pkg['配套']}: {pkg['总课时']}小时")
        
        return packages, subject_hours, max_hours
    
    except Exception as e:
        st.error(f"❌ 文件解析失败: {str(e)}")
        return None, None, None

def calculate_subject_enrollment(packages):
    """计算每个科目的总选修人数"""
    enrollment = defaultdict(int)
    for p_data in packages.values():
        for subject in p_data['科目'].keys():
            enrollment[subject] += p_data['人数']
    return dict(enrollment)

def calculate_recommended_slots(max_total_hours):
    """根据最大总课时计算推荐的时段组数
    
    时段组结构：前(n-1)个时段组各2小时，最后1个时段组3小时
    总容量 = (n-1)*2 + 3 = 2n+1 小时
    
    参数:
        max_total_hours: 所有配套中的最大总课时
    
    返回:
        推荐的时段组数
    """
    import math
    # 如果最大课时<=3，至少需要1个时段组（3小时）
    if max_total_hours <= 3:
        return 1
    # 否则计算需要的时段组数：n = ceil((max_hours - 1) / 2)
    # 这样总容量 2n+1 >= max_hours
    recommended = math.ceil((max_total_hours - 1) / 2)
    # 至少2个时段组，最多20个
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
                # 最后一个是3h时段
                self.SLOT_GROUPS[f'S{i}'] = [i*2-1, i*2, i*2+1]
        
        self.SLOT_TO_GROUP = {}
        for group_name, slots in self.SLOT_GROUPS.items():
            for slot in slots:
                self.SLOT_TO_GROUP[slot] = group_name
        
        self.subject_enrollment = calculate_subject_enrollment(packages)
    
    def build_model(self, objective_type='min_classes'):
        """构建模型"""
        model = cp_model.CpModel()
        
        # 决策变量
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
        
        # 添加约束
        # HA: 精确学时
        for k in self.subjects:
            H_k = self.subject_hours[k]
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                total_hours = sum(y_rt[(k, r, t)] for t in self.TIME_SLOTS_1H)
                model.Add(total_hours == H_k).OnlyEnforceIf(u_r[(k, r)])
                model.Add(total_hours == 0).OnlyEnforceIf(u_r[(k, r)].Not())
        
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
        
        # H2: x_prt逻辑
        for p in self.package_names:
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    for t in self.TIME_SLOTS_1H:
                        model.Add(x_prt[(p, k, r, t)] <= u_pkr[(p, k, r)])
                        model.Add(x_prt[(p, k, r, t)] <= y_rt[(k, r, t)])
                        model.Add(x_prt[(p, k, r, t)] >= u_pkr[(p, k, r)] + y_rt[(k, r, t)] - 1)
        
        # H2': 配套时段互斥
        for p in self.package_names:
            for t in self.TIME_SLOTS_1H:
                model.Add(sum(x_prt[(p, k, r, t)] 
                            for k in self.subjects 
                            for r in range(1, self.config['max_classes_per_subject'] + 1)) <= 1)
        
        # H6: 教师资源约束
        for k in self.subjects:
            for t in self.TIME_SLOTS_1H:
                model.Add(sum(y_rt[(k, r, t)] for r in range(1, self.config['max_classes_per_subject'] + 1)) <= 1)
        
        # H1: 覆盖需求
        for p in self.package_names:
            for k in self.subjects:
                if k in self.packages[p]['科目']:
                    required_hours = self.packages[p]['科目'][k]
                    total_hours_pk = sum(
                        x_prt[(p, k, r, t)]
                        for r in range(1, self.config['max_classes_per_subject'] + 1)
                        for t in self.TIME_SLOTS_1H
                    )
                    model.Add(total_hours_pk == required_hours)
        
        # H4: 开班上限
        for k in self.subjects:
            model.Add(sum(u_r[(k, r)] for r in range(1, self.config['max_classes_per_subject'] + 1)) <= self.config['max_classes_per_subject'])
        
        # H5: 强制开班数
        for k, count in self.config['forced_class_count'].items():
            if k in self.subjects:
                model.Add(sum(u_r[(k, r)] for r in range(1, self.config['max_classes_per_subject'] + 1)) == count)
        
        # 时段分割处理
        slot_split_penalty = 0
        
        if not self.config['allow_slot_split']:
            # 不允许时段分割：添加硬约束
            # 每个配套在每个时段组最多只能上一门课
            for p in self.package_names:
                for group_name, group_slots in self.SLOT_GROUPS.items():
                    subjects_in_group = []
                    for k in self.subjects:
                        for r in range(1, self.config['max_classes_per_subject'] + 1):
                            has_subject = model.NewBoolVar(f'has_{p}_{k}_{r}_{group_name}')
                            model.AddMaxEquality(has_subject, [x_prt[(p, k, r, t)] for t in group_slots])
                            subjects_in_group.append(has_subject)
                    
                    # 硬约束：每个时段组最多一门课
                    model.Add(sum(subjects_in_group) <= 1)
        
        else:
            # 允许时段分割：添加软惩罚，尽量减少分割
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
            
            # 软惩罚：减少分割次数
            slot_split_penalty = sum(split_vars) * self.config['slot_split_penalty']
        
        # 目标函数
        total_classes = sum(u_r[(k, r)] for k in self.subjects for r in range(1, self.config['max_classes_per_subject'] + 1))
        # 修复：使用max(0, ...)避免负惩罚，当选修人数>100时惩罚为0而非负数
        priority_penalty = sum(
            u_r[(k, r)] * r * max(0, 100 - self.subject_enrollment[k])
            for k in self.subjects 
            for r in range(1, self.config['max_classes_per_subject'] + 1)
        )
        
        if objective_type == 'min_classes':
            # 方案A：铁血节省资源。每开一个班罚分极大，绝不多开班。
            model.Minimize(total_classes * 100000 + slot_split_penalty + priority_penalty)
            
        elif objective_type == 'balanced':
            # ========== 1. 定义 max_size 和 min_size (补回缺失的逻辑) ==========
            # 为每个班级创建"有效大小"变量
            # 用于计算 max_size 和 min_size，忽略未开启的班级
            effective_sizes_for_max = []
            effective_sizes_for_min = []
            
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    # 计算该班级的实际人数
                    actual_size = sum(self.packages[p]['人数'] * u_pkr[(p, k, r)] for p in self.package_names)
                    
                    # 有效大小（用于max）：如果开班则=实际大小，否则=0（不影响max计算）
                    eff_size_max = model.NewIntVar(0, 200, f'eff_max_{k}_{r}')
                    model.Add(eff_size_max == actual_size).OnlyEnforceIf(u_r[(k, r)])
                    model.Add(eff_size_max == 0).OnlyEnforceIf(u_r[(k, r)].Not())
                    effective_sizes_for_max.append(eff_size_max)
                    
                    # 有效大小（用于min）：如果开班则=实际大小，否则=200（不影响min计算）
                    # 注意：设为200是因为班额上限通常小于200，这样未开班的200不会成为最小值
                    eff_size_min = model.NewIntVar(0, 200, f'eff_min_{k}_{r}')
                    model.Add(eff_size_min == actual_size).OnlyEnforceIf(u_r[(k, r)])
                    model.Add(eff_size_min == 200).OnlyEnforceIf(u_r[(k, r)].Not())
                    effective_sizes_for_min.append(eff_size_min)
            
            # 定义决策变量：所有开班班级中的最大值和最小值
            max_size = model.NewIntVar(0, 200, 'max_size')
            min_size = model.NewIntVar(0, 200, 'min_size')
            
            # 绑定约束
            model.AddMaxEquality(max_size, effective_sizes_for_max)
            model.AddMinEquality(min_size, effective_sizes_for_min)

            # ========== 2. 定义目标函数 (调整权重) ==========
            # 降低开班权重的绝对统治力，提升均衡权重
            weight_class = 5000  # 开一个班的“成本” (降低，原为100万)
            weight_balance = 200 # 不均衡的“罚款” (提升，每差1人罚200)
            weight_split = self.config.get('slot_split_penalty', 1000) # 时段分割罚款
            
            model.Minimize(
                total_classes * weight_class + 
                (max_size - min_size) * weight_balance + 
                slot_split_penalty * (weight_split / 100) + 
                priority_penalty
            )
        
        return model, {'u_r': u_r, 'y_rt': y_rt, 'u_pkr': u_pkr, 'x_prt': x_prt}
    
    def solve(self, model, variables, timeout):
        """求解模型"""
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
            return {
                'status': 'success',
                'solver': solver,
                'variables': variables,
                'solve_status': status_name,
                'icon': icon,
                'solve_time': solve_time
            }
        else:
            return {
                'status': 'failed',
                'solve_status': status_name,
                'icon': icon,
                'solve_time': solve_time
            }
    
    def analyze_solution(self, result):
        """分析方案"""
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
        
        # 统计时段分割
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
    
    def extract_timetable(self, result):
        """
        提取课表数据（格式调整版）
        1. 班级命名：A, B, C...
        2. 时段总表：
           - [核心修改] 格式调整为：科目+班号+(时长)
           - 例如：从 "化学(1h)A" 改为 "化学A(1h)"
           - 空缺显示为 "0(1h)"
        """
        solver = result['solver']
        u_r = result['variables']['u_r']
        y_rt = result['variables']['y_rt']
        u_pkr = result['variables']['u_pkr']
        
        # ========== 1. 班级命名映射 ==========
        class_name_map = {} 
        for k in self.subjects:
            active_classes = []
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                if solver.Value(u_r[(k, r)]) == 1:
                    students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                    size = sum(self.packages[p]['人数'] for p in students)
                    active_classes.append({'r': r, 'size': size})
            active_classes.sort(key=lambda x: (-x['size'], x['r']))
            for index, item in enumerate(active_classes):
                class_name_map[(k, item['r'])] = f"班{chr(65 + index)}"

        # ========== 2. 开班详情 ==========
        class_details = []
        for k in self.subjects:
            for r in range(1, self.config['max_classes_per_subject'] + 1):
                if solver.Value(u_r[(k, r)]) == 1:
                    students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                    size = sum(self.packages[p]['人数'] for p in students)
                    time_slots = [t for t in self.TIME_SLOTS_1H if solver.Value(y_rt[(k, r, t)]) == 1]
                    slot_groups_used = defaultdict(list)
                    for t in time_slots:
                        slot_groups_used[self.SLOT_TO_GROUP[t]].append(t)
                    slot_str = ', '.join([f"{g}({len(slots)}h)" for g, slots in sorted(slot_groups_used.items(), key=lambda x: natural_sort_key(x[0]))])
                    class_details.append({
                        '科目': k,
                        '班级': class_name_map.get((k, r), f'班{r}'),
                        '人数': size,
                        '时段': slot_str,
                        '学生配套': ', '.join(sorted(students, key=natural_sort_key))
                    })
        class_details.sort(key=lambda x: (x['科目'], x['班级']))

        # ========== 3. 时段总表 ==========
        slot_schedule_data = []
        
        for group_name in sorted(self.SLOT_GROUPS.keys(), key=natural_sort_key):
            group_slots = self.SLOT_GROUPS[group_name]
            group_slots_set = set(group_slots)
            
            # 3.1 收集碎片
            fragments = []
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    active_slots = [t for t in group_slots if solver.Value(y_rt[(k, r, t)]) == 1]
                    actual_hours = len(active_slots)
                    if actual_hours == 0: continue
                    students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                    if not students: continue
                    
                    fragments.append({
                        'subject': f"{k}",
                        'duration_str': f"{actual_hours}h",
                        'class_name': class_name_map.get((k, r), f'班{r}'),
                        'packages_str': ', '.join(sorted(students, key=natural_sort_key)),
                        'raw_packages': students,
                        'size': sum(self.packages[p]['人数'] for p in students),
                        'raw_hours': actual_hours,
                        'active_slots': set(active_slots),
                        'start_time': min(active_slots),
                        'is_gap': False
                    })
            
            # 3.2 贪心拼图
            fragments.sort(key=lambda x: -x['size'])
            visual_rows = []
            for frag in fragments:
                placed = False
                for row in visual_rows:
                    conflict = False
                    for existing in row:
                        if not frag['active_slots'].isdisjoint(existing['active_slots']):
                            conflict = True; break
                    if not conflict:
                        row.append(frag); placed = True; break
                if not placed: visual_rows.append([frag])
            
            # 3.3 填空 & 格式化
            for row_items in visual_rows:
                # Gap Filling
                occupied_slots = set()
                for item in row_items: occupied_slots.update(item['active_slots'])
                missing_slots = sorted(list(group_slots_set - occupied_slots))
                
                if missing_slots:
                    import itertools
                    for _, g in itertools.groupby(enumerate(missing_slots), lambda ix: ix[0] - ix[1]):
                        gap_group = list(map(lambda ix: ix[1], g))
                        row_items.append({
                            'subject': '0',
                            'duration_str': f"{len(gap_group)}h",
                            'class_name': '-',
                            'packages_str': '-',
                            'raw_packages': [],
                            'size': 0,
                            'raw_hours': 0,
                            'active_slots': set(gap_group),
                            'start_time': min(gap_group),
                            'is_gap': True
                        })
                
                row_items.sort(key=lambda x: x['start_time'])
                
                # [核心修改] 字符串拼接顺序：科目 + 班号 + (时长)
                merged_items_str = []
                for i in row_items:
                    if i['is_gap']:
                        # Gap 显示: 0(1h)
                        item_str = f"{i['subject']}({i['duration_str']})"
                    else:
                        # 正常课程: 化学A(1h)
                        cls_short = i['class_name'].replace('班', '') # 把"班A"变成"A"
                        item_str = f"{i['subject']}{cls_short}({i['duration_str']})"
                    
                    merged_items_str.append(item_str)
                
                merged_info = " + ".join(merged_items_str)
                merged_packages = " + ".join([i['packages_str'] for i in row_items])
                
                # 去重人数
                unique_pkgs = set()
                for i in row_items:
                    for p in i['raw_packages']: unique_pkgs.add(p)
                unique_count = sum(self.packages[p]['人数'] for p in unique_pkgs)
                
                # UI Display Items
                display_list = []
                for idx, item in enumerate(row_items):
                    display_list.append({
                        'seq': idx + 1,
                        'subject': item['subject'],
                        'duration': item['duration_str'],
                        'class': item['class_name'],
                        'color_seed': item['subject'] if not item['is_gap'] else 'gap',
                        'is_gap': item['is_gap']
                    })

                slot_schedule_data.append({
                    '时段': group_name,
                    '时长': f"{sum(i['raw_hours'] for i in row_items)}h",
                    '科目 & 班级': merged_info,
                    '人数': unique_count,
                    '涉及配套': merged_packages,
                    'display_items': display_list
                })
        
        return class_details, slot_schedule_data

# ========== 主应用 ==========
def main():
    st.markdown('<div class="main-header">📚 智能排课求解器</div>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #666;">走班制排课搜索系统</p>', unsafe_allow_html=True)
    
    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        st.subheader("📁 数据导入")
        
        # 下载模板功能
        st.markdown("##### 📥 下载数据模板")
        st.markdown("""
        <div style="padding: 0.5rem; border-radius: 0.3rem; margin-bottom: 0.5rem; font-size: 0.85rem;">
        💡 首次使用？下载示例模板了解数据格式
        </div>
        """, unsafe_allow_html=True)
        
        # 创建示例CSV数据
        template_data = """配套,科目,人数,总学点
P12,"会计学（4）,经济（4）,商业（3）,历史（4）,AI应用（2）,AI编程（2）",5,19
P13,"物理（6）,经济（4）,历史（4）,地理（4）,AI应用（2）",6,20
P14,"物理（6）,会计学（4）,经济（4）,商业（3）,AI应用（2）,AI编程（2）",4,21
P15,"生物（4）,化学（5）,物理（6）,会计学（4）,AI应用（2）",9,21
P16,"生物（4）,化学（5）,物理（6）,商业（3）,AI应用（2）",3,20
P17,"生物（4）,化学（5）,会计学（4）,地理（4）,AI应用（2）,AI编程（2）",8,21
P18,"生物（4）,化学（5）,经济（4）,历史（4）,AI应用（2）,AI编程（2）",11,21
P19,"物理（6）,经济（4）,商业（3）,历史（4）,AI应用（2）,AI编程（2）",7,21
P20,"物理（6）,生物（4）,化学（5）,经济（4）,AI应用（2）",10,21
P21,"物理（6）,生物（4）,化学（5）,地理（4）,AI应用（2）",2,21
P22,"生物（4）,化学（5）,经济（4）,地理（4）,AI应用（2）,AI编程（2）",12,21"""
        
        # 下载按钮
        col1, col2 = st.columns([1, 1])
        with col1:
            st.download_button(
                label="📄 CSV模板",
                data=template_data.encode('utf-8-sig'),  # 使用BOM确保Excel正确识别UTF-8
                file_name="排课数据模板.csv",
                mime="text/csv",
                help="下载CSV格式的示例模板",
                use_container_width=True
            )
        with col2:
            # 创建Excel格式的模板
            template_df = pd.DataFrame([
                {'配套': 'P12', '科目': '会计学（4）,经济（4）,商业（3）,历史（4）,AI应用（2）,AI编程（2）', '人数': 5, '总学点': 19},
                {'配套': 'P13', '科目': '物理（6）,经济（4）,历史（4）,地理（4）,AI应用（2）', '人数': 6, '总学点': 20},
                {'配套': 'P14', '科目': '物理（6）,会计学（4）,经济（4）,商业（3）,AI应用（2）,AI编程（2）', '人数': 4, '总学点': 21},
                {'配套': 'P15', '科目': '生物（4）,化学（5）,物理（6）,会计学（4）,AI应用（2）', '人数': 9, '总学点': 21},
            ])
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                template_df.to_excel(writer, index=False, sheet_name='配套数据')
            
            st.download_button(
                label="📊 Excel模板",
                data=excel_buffer.getvalue(),
                file_name="排课数据模板.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="下载Excel格式的示例模板",
                use_container_width=True
            )
        
        st.markdown("---")
        
        # 文件上传
        st.markdown("##### 📤 上传数据文件")
        uploaded_file = st.file_uploader(
            "选择文件",
            type=['xlsx', 'xls', 'csv'],
            help="支持Excel和CSV格式，需包含'配套'、'科目'、'人数'列",
            label_visibility="collapsed"
        )
        
        if uploaded_file:
            with st.spinner("正在解析文件..."):
                packages, subject_hours, max_hours = parse_uploaded_file(uploaded_file)
            
            if packages and subject_hours:
                st.success(f"✅ 成功加载 {len(packages)} 个配套，{len(subject_hours)} 个科目")
                st.session_state['packages'] = packages
                st.session_state['subject_hours'] = subject_hours
                st.session_state['max_total_hours'] = max_hours  # 保存最大总课时
        
        st.markdown("---")
        
        st.subheader("🔧 求解参数")
        
        min_class_size = st.number_input("最小班额", min_value=1, max_value=100, value=5, step=1)
        max_class_size = st.number_input("最大班额", min_value=1, max_value=200, value=60, step=1)
        max_classes_per_subject = st.number_input("每科目最大班数", min_value=1, max_value=10, value=3, step=1)
        
        # 智能推荐时段组数
        if 'max_total_hours' in st.session_state:
            max_hours = st.session_state['max_total_hours']
            recommended_slots = calculate_recommended_slots(max_hours)
            total_capacity = (recommended_slots - 1) * 2 + 3
            
            st.markdown(f"""
            <div style="padding: 1rem; border-radius: 0.5rem; border-left: 4px solid #2196f3; margin: 1rem 0;">
                <strong>📊 智能分析</strong><br>
                • 最大总课时：<strong>{max_hours}小时</strong><br>
                • 推荐时段组数：<strong>{recommended_slots}组</strong> (总容量{total_capacity}小时)<br>
                • 说明：{recommended_slots-1}组×2小时 + 1组×3小时 = {total_capacity}小时
            </div>
            """, unsafe_allow_html=True)
            
            default_slots = recommended_slots
        else:
            default_slots = 10
            st.info("💡 上传数据后将自动推荐时段组数")
        
        num_slots = st.number_input(
            "时段组数量", 
            min_value=1, 
            max_value=20, 
            value=default_slots, 
            step=1,
            help="系统会根据数据自动推荐，也可手动调整。最后一个时段组为3小时，其余为2小时"
        )
        
        solver_timeout = st.number_input("求解超时(秒)", min_value=10, max_value=600, value=120, step=10)
        
        st.markdown("---")
        
        st.subheader("🔀 时段分割")
        allow_slot_split = st.checkbox("允许时段分割", value=True,
                                      help="允许一个时段内上不同科目的课")
        if allow_slot_split:
            slot_split_penalty = st.slider("分割惩罚系数", min_value=0, max_value=5000, value=1000, step=100,
                                          help="越大越不愿意分割")
        else:
            slot_split_penalty = 0
        
        st.markdown("---")
        
        st.subheader("🔒 强制开班")
        if 'subject_hours' in st.session_state:
            forced_class_count = {}
            for subject in st.session_state['subject_hours'].keys():
                count = st.number_input(f"{subject}", min_value=0, max_value=10, value=0, key=f"forced_{subject}")
                if count > 0:
                    forced_class_count[subject] = count
        else:
            forced_class_count = {}
            st.info("请先上传数据文件")
    
    # 主内容区
    if 'packages' not in st.session_state:
        st.markdown('<div class="info-box">', unsafe_allow_html=True)
        st.markdown("""
        ### 智能排课搜索器！
        
        **使用步骤：**
        1. 📁 在左侧上传配套数据文件（Excel或CSV格式）
        2. ⚙️ 调整求解参数（可选）
        3. 🚀 点击"开始求解"按钮
        4. 📊 查看并下载结果
        
        **数据格式要求：**
        - 必须包含列：`配套`、`科目`、`人数`
        - 科目格式：`会计(6),历史(4),地理(4)` 或 `会计（6）,历史（4）`

    
    
        **功能：**
        - 🎯 自动生成多个优化方案
        - 🔀 支持时段分割（一个时段上不同科目）
        - 👨‍🏫 教师资源约束（同科目不同班不冲突）
        - 📊 时段总表（查看每个时段的全局安排）
        - ⏰ 灵活课时
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        return
    
    # 显示数据概览
    st.markdown('<div class="sub-header">📊 数据概览</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("配套数量", len(st.session_state['packages']))
    with col2:
        st.metric("科目数量", len(st.session_state['subject_hours']))
    with col3:
        total_students = sum(p['人数'] for p in st.session_state['packages'].values())
        st.metric("学生总数", total_students)
    
    # 配套详情
    with st.expander("查看配套详情"):
        df_packages = []
        for name, data in st.session_state['packages'].items():
            subjects_str = ', '.join([f"{k}({v}h)" for k, v in data['科目'].items()])
            df_packages.append({
                '配套': name,
                '人数': data['人数'],
                '科目': subjects_str
            })
        st.dataframe(pd.DataFrame(df_packages), use_container_width=True)
    
    # 科目选修统计
    with st.expander("查看科目选修统计"):
        enrollment = calculate_subject_enrollment(st.session_state['packages'])
        df_enrollment = pd.DataFrame([
            {'科目': k, '课时': st.session_state['subject_hours'][k], '选修人数': enrollment[k]}
            for k in sorted(enrollment.keys(), key=lambda x: enrollment[x], reverse=True)
        ])
        st.dataframe(df_enrollment, use_container_width=True)
    
    st.markdown("---")
    
    # 求解按钮
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
        
        solver_instance = ScheduleSolver(
            st.session_state['packages'],
            st.session_state['subject_hours'],
            config
        )
        
        # 生成多个方案
        solution_configs = [
            {'type': 'min_classes', 'name': '方案A：最少开班'},
            {'type': 'balanced', 'name': '方案B：均衡班额'}
        ]
        
        # 创建进度条容器
        progress_container = st.container()
        with progress_container:
            progress_bar = st.progress(0)
            col1, col2 = st.columns([3, 1])
            with col1:
                status_text = st.empty()
            with col2:
                percentage_text = st.empty()
        
        solutions = []
        total_steps = len(solution_configs) * 3  # 每个方案3个步骤：准备、建模、求解
        current_step = 0
        
        for i, sol_config in enumerate(solution_configs):
            # 步骤1: 准备阶段
            current_step += 1
            progress = current_step / total_steps
            progress_bar.progress(progress)
            status_text.markdown(f"🔄 **{sol_config['name']}** - 准备中...")
            percentage_text.markdown(f"**{int(progress * 100)}%**")
            time.sleep(0.1)  # 短暂延迟使进度可见
            
            # 步骤2: 建模阶段
            current_step += 1
            progress = current_step / total_steps
            progress_bar.progress(progress)
            status_text.markdown(f"🏗️ **{sol_config['name']}** - 构建模型...")
            percentage_text.markdown(f"**{int(progress * 100)}%**")
            
            model, variables = solver_instance.build_model(sol_config['type'])
            
            # 步骤3: 求解阶段
            current_step += 1
            progress = current_step / total_steps
            progress_bar.progress(progress)
            status_text.markdown(f"⚙️ **{sol_config['name']}** - 正在求解...")
            percentage_text.markdown(f"**{int(progress * 100)}%**")
            
            result = solver_instance.solve(model, variables, solver_timeout)
            
            if result['status'] == 'success':
                result['name'] = sol_config['name']
                result['analysis'] = solver_instance.analyze_solution(result)
                result['class_details'], result['slot_schedule'] = solver_instance.extract_timetable(result)
                solutions.append(result)
                status_text.markdown(f"✅ **{sol_config['name']}** - 完成")
        
        # 完成后显示100%
        progress_bar.progress(1.0)
        percentage_text.markdown("**100%**")
        status_text.markdown("🎉 **所有方案求解完成！**")
        time.sleep(0.5)
        
        # 清空进度条
        progress_bar.empty()
        status_text.empty()
        percentage_text.empty()
        
        if not solutions:
            st.markdown('<div class="error-box">', unsafe_allow_html=True)
            st.error("❌ 所有方案均无解！")
            st.markdown("""
            **可能原因：**
            - 时段数量不足
            - 班额限制过严
            - 强制开班数设置不合理
            
            **建议解决方案：**
            1. 增加时段组数量
            2. 放宽班额上限
            3. 取消强制开班限制
            4. 启用时段分割功能
            """)
            st.markdown('</div>', unsafe_allow_html=True)
            return
        
        st.session_state['solutions'] = solutions
        
        # 显示结果
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.success(f"✅ 成功生成 {len(solutions)} 个方案！")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # 显示方案结果
    if 'solutions' in st.session_state:
        st.markdown("---")
        st.markdown('<div class="sub-header">📊 方案对比</div>', unsafe_allow_html=True)
        
        comparison_data = []
        for sol in st.session_state['solutions']:
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
        
        df_comparison = pd.DataFrame(comparison_data)
        st.dataframe(df_comparison, use_container_width=True)
        
        # 方案详情
        for sol in st.session_state['solutions']:
            with st.expander(f"📋 {sol['name']} - 详细结果"):
                tab1, tab2, tab3 = st.tabs(["开班详情", "时段总表", "数据导出"])
                
                with tab1:
                    df_class = pd.DataFrame(sol['class_details'])
                    st.dataframe(df_class, use_container_width=True)
                    
                    if sol['analysis']['split_count'] > 0:
                        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                        st.warning(f"⚠️ 检测到 {sol['analysis']['split_count']} 处时段分割")
                        for detail in sol['analysis']['split_details']:
                            st.text(f"  • {detail}")
                        st.markdown('</div>', unsafe_allow_html=True)
                
                with tab2:
                    st.markdown("### 🕐 时段总表")
                    
                    schedule_data = sol['slot_schedule']
                    if not schedule_data:
                        st.info("暂无数据")
                    else:
                        # ========== HTML 表格 (流程卡片优化版) ==========
                        
                        table_css = """
                        <style>
                            .schedule-table {
                                width: 100%;
                                border-collapse: collapse;
                                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                                margin-bottom: 1rem;
                                font-size: 15px;
                                color: #ffffff; 
                            }
                            .schedule-table th {
                                background-color: #262730;
                                color: #ffffff;
                                font-weight: 700;
                                padding: 12px;
                                text-align: left;
                                border-bottom: 2px solid #4a4a4a;
                                border-top: 1px solid #4a4a4a;
                                white-space: nowrap;
                            }
                            .schedule-table td {
                                padding: 8px 12px; /* 减小内边距 */
                                text-align: left;
                                border-right: 1px solid #333333;
                                color: #e0e0e0;
                                vertical-align: middle;
                            }
                            .group-border-bottom { border-bottom: 3px solid #666666 !important; }
                            .normal-border-bottom { border-bottom: 1px solid #333333; }
                            
                            .slot-column {
                                font-weight: 800; font-size: 1.2rem;
                                text-align: center !important;
                                background-color: #1a1c24; color: #4fc3f7;
                                width: 80px; border-right: 2px solid #4a4a4a !important;
                            }
                            .duration-column {
                                text-align: center !important; width: 60px;
                                font-weight: 600; color: #90caf9;
                            }
                            
                            /* === 核心优化：流程卡片样式 === */
                            .timeline-container {
                                display: flex;
                                align-items: center;
                                flex-wrap: wrap; /* 如果屏幕太窄允许换行 */
                                gap: 6px;
                            }
                            .timeline-card {
                                background-color: #333333;
                                border: 1px solid #444;
                                border-radius: 6px;
                                padding: 4px 8px;
                                display: flex;
                                flex-direction: column;
                                min-width: 110px;
                            }
                            .card-header {
                                display: flex;
                                align-items: center;
                                margin-bottom: 2px;
                            }
                            .seq-badge {
                                background-color: #0288d1; /* 序号底色 */
                                color: white;
                                font-size: 0.75rem;
                                font-weight: bold;
                                width: 16px; height: 16px;
                                border-radius: 50%;
                                display: flex; align-items: center; justify-content: center;
                                margin-right: 6px;
                            }
                            .subject-name {
                                font-weight: 800; color: #fff; font-size: 0.95rem;
                            }
                            .card-footer {
                                display: flex; justify-content: space-between;
                                font-size: 0.8rem; color: #aaa;
                            }
                            .duration-tag {
                                background-color: #424242; padding: 0 4px; border-radius: 3px;
                            }
                            .arrow-icon {
                                color: #666; font-size: 1.2rem; font-weight: bold;
                                margin: 0 2px;
                            }
                            
                            .count-cell { font-weight: bold; font-size: 1.1rem; text-align: center; color: #fff; }
                            .package-cell { color: #b0bec5; font-size: 0.9rem; }
                        </style>
                        """
                        
                        html_rows = []
                        from itertools import groupby
                        schedule_data.sort(key=lambda x: natural_sort_key(x['时段']))
                        
                        for slot_name, items in groupby(schedule_data, key=lambda x: x['时段']):
                            group_items = list(items)
                            row_count = len(group_items)
                            for i, item in enumerate(group_items):
                                border_class = "group-border-bottom" if i == row_count - 1 else "normal-border-bottom"
                                row_html = f"<tr class='{border_class}'>"
                                
                                if i == 0:
                                    row_html += f"<td class='slot-column' rowspan='{row_count}'>{item['时段']}</td>"
                                    row_html += f"<td class='duration-column' rowspan='{row_count}'>{item['时长']}</td>"
                                
                                # === [核心] 构建流程图 ===
                                flow_html = '<div class="timeline-container">'
                                display_items = item.get('display_items', [])
                                
                                for idx, d_item in enumerate(display_items):
                                    # 卡片 HTML
                                    card = f"""
                                    <div class="timeline-card">
                                        <div class="card-header">
                                            <span class="seq-badge">{d_item['seq']}</span>
                                            <span class="subject-name">{d_item['subject']}</span>
                                        </div>
                                        <div class="card-footer">
                                            <span>{d_item['class']}</span>
                                            <span class="duration-tag">{d_item['duration']}</span>
                                        </div>
                                    </div>
                                    """
                                    flow_html += card
                                    
                                    # 如果不是最后一个，添加箭头
                                    if idx < len(display_items) - 1:
                                        flow_html += '<div class="arrow-icon">➜</div>'
                                
                                flow_html += '</div>'
                                row_html += f"<td>{flow_html}</td>"
                                
                                row_html += f"<td class='count-cell'>{item['人数']}</td>"
                                row_html += f"<td class='package-cell'>{item['涉及配套']}</td>"
                                row_html += "</tr>"
                                html_rows.append(row_html)
                        
                        full_html = f"""
                        {table_css}
                        <table class="schedule-table">
                            <thead>
                                <tr>
                                    <th>时段</th>
                                    <th>时长</th>
                                    <th>课程流程 (顺序)</th> <th style="text-align: center;">人数</th>
                                    <th>涉及配套</th>
                                </tr>
                            </thead>
                            <tbody>{''.join(html_rows)}</tbody>
                        </table>
                        """
                        st.markdown(full_html, unsafe_allow_html=True)

                    # ========== 3. 统计信息 (保持不变) ==========
                    st.markdown("### 📊 统计信息")
                    df_slot = pd.DataFrame(schedule_data)
                    # 导出Excel时，要把 display_items 这个辅助字段去掉，防止报错
                    if 'display_items' in df_slot.columns:
                        df_slot = df_slot.drop(columns=['display_items'])
                        
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总时段数", df_slot['时段'].nunique() if not df_slot.empty else 0)
                    with col2:
                        st.metric("总条目数", len(df_slot))
                    with col3:
                        unique = df_slot['时段'].nunique() if not df_slot.empty else 0
                        avg = len(df_slot) / unique if unique > 0 else 0
                        st.metric("平均每时段条目", f"{avg:.1f}")
                                
                with tab3:
                    # 导出为Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # 获取数据
                        df_class = pd.DataFrame(sol['class_details'])
                        df_slot = pd.DataFrame(sol['slot_schedule'])
                        
                        # [修复] 剔除只用于网页显示的辅助列 'display_items'
                        if 'display_items' in df_slot.columns:
                            df_slot = df_slot.drop(columns=['display_items'])
                        
                        # [重要] 确保开班详情按 科目 -> 班级(A,B) 排序
                        df_class = df_class.sort_values(by=['科目', '班级'])
                        
                        # 1. 写入 "开班详情" Sheet
                        df_class.to_excel(writer, sheet_name='开班详情', index=False)
                        
                        # 2. 写入 "时段总表" Sheet
                        df_slot.to_excel(writer, sheet_name='时段总表', index=False)
                        
                        # 3. 写入 "所有班级及涉及的配套" Sheet
                        #  - 先复制一份数据
                        df_overview = df_class.copy()
                        
                        #  - 合并列并去除"班"字：例如 "化学" + "班A" -> "化学A"
                        df_overview['科目 & 班级'] = df_overview['科目'] + df_overview['班级'].str.replace('班', '')
                        
                        #  - 只保留合并后的列、人数和配套
                        df_overview = df_overview[['科目 & 班级', '学生配套']]
                        #  - 重命名配套列
                        df_overview.columns = ['科目 SUBJECT', '配套 PACKAGE']
                        
                        df_overview.to_excel(writer, sheet_name='导入', index=False)
                        
                        # === 自动调整列宽逻辑 ===
                        workbook = writer.book
                        
                        for sheet_name in writer.sheets:
                            worksheet = writer.sheets[sheet_name]
                            
                            # 根据当前Sheet选择对应的DataFrame来计算列宽
                            if sheet_name == '时段总表':
                                df_to_measure = df_slot
                            elif sheet_name == '导入':
                                df_to_measure = df_overview
                            else:
                                df_to_measure = df_class
                                
                            for idx, col in enumerate(df_to_measure.columns):
                                # 计算最大长度
                                max_len = max(
                                    len(str(col)),
                                    df_to_measure[col].astype(str).str.len().max() if not df_to_measure[col].empty else 0
                                )
                                # 设置宽度 (限制最大 60)
                                adjusted_width = min(max_len + 4, 60)
                                worksheet.column_dimensions[get_column_letter(idx + 1)].width = adjusted_width
                    
                    st.download_button(
                        label="📥 下载Excel文件",
                        data=output.getvalue(),
                        file_name=f"{sol['name'].replace('：', '_')}_排课结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
if __name__ == "__main__":
    main()
