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
            
            # 收集所有科目的课时
            for subject, hours in subjects.items():
                if subject not in subject_hours:
                    subject_hours[subject] = hours
                elif subject_hours[subject] != hours:
                    st.warning(f"⚠️ 科目'{subject}'的课时不一致: {subject_hours[subject]} vs {hours}")
        
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
        
        return packages, subject_hours
    
    except Exception as e:
        st.error(f"❌ 文件解析失败: {str(e)}")
        return None, None

def calculate_subject_enrollment(packages):
    """计算每个科目的总选修人数"""
    enrollment = defaultdict(int)
    for p_data in packages.values():
        for subject in p_data['科目'].keys():
            enrollment[subject] += p_data['人数']
    return dict(enrollment)

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
        
        # 时段分割惩罚
        slot_split_penalty = 0
        if self.config['allow_slot_split']:
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
        priority_penalty = sum(
            u_r[(k, r)] * r * (100 - self.subject_enrollment[k])
            for k in self.subjects 
            for r in range(1, self.config['max_classes_per_subject'] + 1)
        )
        
        if objective_type == 'min_classes':
            model.Minimize(total_classes * 100000 + slot_split_penalty + priority_penalty)
        elif objective_type == 'balanced':
            class_sizes = []
            for k in self.subjects:
                for r in range(1, self.config['max_classes_per_subject'] + 1):
                    size = sum(self.packages[p]['人数'] * u_pkr[(p, k, r)] for p in self.package_names)
                    class_sizes.append(size)
            max_size = model.NewIntVar(0, 200, 'max_size')
            min_size = model.NewIntVar(0, 200, 'min_size')
            for size in class_sizes:
                model.Add(max_size >= size)
            model.Minimize(total_classes * 1000000 + slot_split_penalty * 100 + (max_size - min_size) * 1000 + priority_penalty)
        
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
        """提取课表数据"""
        solver = result['solver']
        u_r = result['variables']['u_r']
        y_rt = result['variables']['y_rt']
        u_pkr = result['variables']['u_pkr']
        x_prt = result['variables']['x_prt']
        
        # 开班详情
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
                    
                    slot_str = ', '.join([f"{g}({len(slots)}h)" for g, slots in sorted(slot_groups_used.items())])
                    
                    class_details.append({
                        '科目': k,
                        '班级': f'班{r}',
                        '人数': size,
                        '时段': slot_str,
                        '学生配套': ', '.join(students)
                    })
        
        # 时段总表
        slot_schedule_data = []
        for group_name in sorted(self.SLOT_GROUPS.keys()):
            group_slots = self.SLOT_GROUPS[group_name]
            row = {'时段': group_name, '时长': f'{len(group_slots)}h'}
            
            # 找出该时段所有上课的班级
            classes_in_slot = []
            packages_in_slot = set()
            
            for t in group_slots:
                for k in self.subjects:
                    for r in range(1, self.config['max_classes_per_subject'] + 1):
                        if solver.Value(y_rt[(k, r, t)]) == 1:
                            # 该班在这个时段上课
                            students = [p for p in self.package_names if solver.Value(u_pkr[(p, k, r)]) == 1]
                            size = sum(self.packages[p]['人数'] for p in students)
                            class_info = f"{k}班{r}({size}人)"
                            if class_info not in classes_in_slot:  # 避免重复
                                classes_in_slot.append(class_info)
                                packages_in_slot.update(students)
            
            # 找出空闲的配套（在这个时段没有课的配套）
            all_packages = set(self.package_names)
            free_packages = all_packages - packages_in_slot
            
            row['上课班级'] = ', '.join(classes_in_slot) if classes_in_slot else '-'
            row['涉及配套'] = ', '.join(sorted(packages_in_slot)) if packages_in_slot else '-'
            row['空闲配套'] = ', '.join(sorted(free_packages)) if free_packages else '-'
            row['班级数'] = len(classes_in_slot)
            row['上课配套数'] = len(packages_in_slot)
            row['空闲配套数'] = len(free_packages)
            
            slot_schedule_data.append(row)
        
        return class_details, slot_schedule_data

# ========== 主应用 ==========
def main():
    st.markdown('<div class="main-header">📚 智能排课求解器 v3.6</div>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #666;">基于约束编程的走班制排课优化系统</p>', unsafe_allow_html=True)
    
    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 系统配置")
        
        st.subheader("📁 数据导入")
        uploaded_file = st.file_uploader(
            "上传配套数据文件",
            type=['xlsx', 'xls', 'csv'],
            help="支持Excel和CSV格式，需包含'配套'、'科目'、'人数'列"
        )
        
        if uploaded_file:
            with st.spinner("正在解析文件..."):
                packages, subject_hours = parse_uploaded_file(uploaded_file)
            
            if packages and subject_hours:
                st.success(f"✅ 成功加载 {len(packages)} 个配套，{len(subject_hours)} 个科目")
                st.session_state['packages'] = packages
                st.session_state['subject_hours'] = subject_hours
        
        st.markdown("---")
        
        st.subheader("🔧 求解参数")
        
        min_class_size = st.number_input("最小班额", min_value=1, max_value=100, value=5, step=1)
        max_class_size = st.number_input("最大班额", min_value=1, max_value=200, value=60, step=1)
        max_classes_per_subject = st.number_input("每科目最大班数", min_value=1, max_value=10, value=3, step=1)
        num_slots = st.number_input("时段组数量", min_value=5, max_value=20, value=10, step=1, 
                                   help="最后一个时段组为3小时，其余为2小时")
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
        ### 👋 欢迎使用智能排课求解器！
        
        **使用步骤：**
        1. 📁 在左侧上传配套数据文件（Excel或CSV格式）
        2. ⚙️ 调整求解参数（可选）
        3. 🚀 点击"开始求解"按钮
        4. 📊 查看并下载结果
        
        **数据格式要求：**
        - 必须包含列：`配套`、`科目`、`人数`
        - 科目格式：`会计(6),历史(4),地理(4)` 或 `会计（6）,历史（4）`
        - ✅ **同时支持英文括号()和中文括号（）**
        - ✅ **支持总课时不足21的配套**（这些配套在某些时段不上课）
        
        **示例：**
        ```
        配套 | 科目                              | 人数
        P1  | 会计(6),历史(4),地理(4),商业(3)    | 24  (总17h)
        P2  | 生物(4),会计(6),历史(4),商业(3)    | 5   (总17h)
        P3  | 物理(6),化学(5)                    | 10  (总11h) ← 总课时少也没问题！
        ```
        
        **特色功能：**
        - 🎯 自动生成多个优化方案
        - 🔀 支持时段分割（一个时段上不同科目）
        - 👨‍🏫 教师资源约束（同科目不同班不冲突）
        - 📊 时段总表（查看每个时段的全局安排）
        - ⏰ 灵活课时（配套总课时可以小于21）
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
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        solutions = []
        for i, sol_config in enumerate(solution_configs):
            status_text.text(f"正在生成{sol_config['name']}...")
            progress_bar.progress((i + 1) / len(solution_configs))
            
            model, variables = solver_instance.build_model(sol_config['type'])
            result = solver_instance.solve(model, variables, solver_timeout)
            
            if result['status'] == 'success':
                result['name'] = sol_config['name']
                result['analysis'] = solver_instance.analyze_solution(result)
                result['class_details'], result['slot_schedule'] = solver_instance.extract_timetable(result)
                solutions.append(result)
        
        progress_bar.empty()
        status_text.empty()
        
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
                    st.markdown("### 🕐 时段总表（全局视图）")
                    st.markdown("*显示每个时段有哪些班级在上课，哪些配套是空闲的*")
                    df_slot = pd.DataFrame(sol['slot_schedule'])
                    st.dataframe(df_slot, use_container_width=True)
                    
                    # 统计空闲情况
                    total_slots = len(df_slot)
                    slots_with_free = sum(1 for row in sol['slot_schedule'] if row['空闲配套'] != '-')
                    avg_free = sum(row['空闲配套数'] for row in sol['slot_schedule']) / total_slots if total_slots > 0 else 0
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总时段数", total_slots)
                    with col2:
                        st.metric("有空闲配套的时段", slots_with_free)
                    with col3:
                        st.metric("平均每时段空闲配套数", f"{avg_free:.1f}")
                    
                    if avg_free > 0:
                        st.info(f"💡 提示：平均每个时段有{avg_free:.1f}个配套是空闲的，这些时段可以用于自习、活动等安排")
                
                with tab3:
                    # 导出为Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pd.DataFrame(sol['class_details']).to_excel(writer, sheet_name='开班详情', index=False)
                        pd.DataFrame(sol['slot_schedule']).to_excel(writer, sheet_name='时段总表', index=False)
                    
                    st.download_button(
                        label="📥 下载Excel文件",
                        data=output.getvalue(),
                        file_name=f"{sol['name'].replace('：', '_')}_排课结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
