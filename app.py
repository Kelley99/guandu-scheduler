#!/usr/bin/env python3
"""
官渡自动分配工具 - Flask后端
"""

import os
import re
import json
import random
from datetime import datetime
from flask import send_file, Flask, render_template, request, jsonify, send_file, session, Response, make_response
from werkzeug.utils import secure_filename
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

app = Flask(__name__)
app.secret_key = 'guandu-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# CORS 头
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    return response


# 数据目录
DATA_DIR = os.path.join(os.path.dirname(__file__), 'knowledge-base')
STATS_FILE = os.path.join(DATA_DIR, '凌霄数据统计表26.3.30.md')
GUANDU_FILE = os.path.join(DATA_DIR, '凌霄官渡26.md')

# 上传目录
UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx', 'csv', 'md'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ============ 解析函数 ============

def parse_stats_table(content: str) -> dict:
    """解析统计表，返回 {成员名: {hp, total, power, ...}}
    
    格式1: | 序号 | 成员名称 | 集结加成 | 步兵防御 | 步兵生命值 | ... | 六维属性总和 | ...
    格式2: | 序号 | 成员名称 | 战力 | ...
    支持动态定位 步兵生命值、六维属性总和、战力 列
    """
    stats = {}
    hp_col_idx = None
    total_col_idx = None
    power_col_idx = None
    
    for line in content.split('\n'):
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|')]
        parts = [p for p in parts if p]
        
        # 找表头确定各列索引
        if hp_col_idx is None and '步兵生命值' in parts:
            hp_col_idx = parts.index('步兵生命值')
        if total_col_idx is None and '六维属性总和' in parts:
            total_col_idx = parts.index('六维属性总和')
        if power_col_idx is None and '战力' in parts:
            power_col_idx = parts.index('战力')
        
        # 数据行
        if len(parts) >= 5 and parts[0].isdigit():
            name = parts[1]
            if name in stats:
                continue  # 避免重复
            
            hp = 0
            total = 0
            power = 0
            
            # 获取步兵生命值
            if hp_col_idx is not None and len(parts) > hp_col_idx:
                try:
                    hp = float(parts[hp_col_idx])
                except:
                    pass
            
            # 获取六维属性总和
            if total_col_idx is not None and len(parts) > total_col_idx:
                try:
                    total = float(parts[total_col_idx])
                except:
                    pass
            
            # 获取战力
            if power_col_idx is not None and len(parts) > power_col_idx:
                try:
                    power = float(parts[power_col_idx])
                except:
                    pass
            
            # 如果HP值异常（>100000），尝试智能检测
            if hp > 100000:
                for i, p in enumerate(parts):
                    try:
                        v = float(p)
                        if 100 <= v <= 5000:
                            hp = v
                            break
                    except:
                        pass
            
            # 如果还是异常，设为0
            if hp > 100000:
                hp = 0
            
            stats[name] = {
                'hp': hp,
                'total': total,
                'power': power,
            }
    
    return stats


def parse_guandu_table(content: str, section: str) -> dict:
    """解析官渡表，返回各队数据和候补名单
    
    表格结构:
    | 队伍 | 队长 | 队员 | 分组 | 队员1 | 战术1 | 战术2 |
    parts: ['', '队伍', '队长', '队员', '分组', '队员1', '战术1', '战术2']
    
    A/B行结构:
    | | | A | 队员名 |
    parts: ['', '', 'A', '队员名']
    """
    pattern = rf'## {re.escape(section)}[\s\S]*?(?=## 团|$)'
    match = re.search(pattern, content)
    if not match:
        return {'teams': {}, 'bench': []}
    
    section_content = match.group(0)
    lines = section_content.split('\n')
    
    tactic_kw = ['0-10分钟', '10-20分钟', '20分钟以后', '拿下', '驻守', '集结', 
                 '粮仓', '乌巢', '官渡', '霹雳', '锱重', '驻防', '采集', '远程',
                 '不动', '应变', '必要时', '跟随', '工匠坊', '兵器坊', '首占']
    
    current_team = None
    teams_data = {}
    bench_members = []
    
    for line in lines:
        # 保留原始列位置，不过滤空字符串
        raw_parts = [p.strip() for p in line.split('|')]
        # 同时生成过滤版用于简单判断
        parts = [p for p in raw_parts if p]
        
        if not parts:
            continue
        
        # 跳过分隔符行
        if all(re.match(r'^-+$', p) for p in parts):
            continue
        
        # 队号行: parts[0] = '1队' 等
        # 原始列位置: col1=队伍, col2=队长, col3=分组(A/B/空), col4=队员, col5=0-10分钟, col6=10-20分钟, col7=20分钟以后
        if re.match(r'^\d+队$', parts[0]):
            current_team = parts[0]
            captain = raw_parts[2].strip() if len(raw_parts) > 2 else parts[1]
            teams_data[current_team] = {
                'captain': captain,
                'A_members': [],
                'B_members': [],
                'A_tasks': {'0-10': '', '10-20': '', '20+': ''},
                'B_tasks': {'0-10': '', '10-20': '', '20+': ''},
            }
            group = raw_parts[3].strip() if len(raw_parts) > 3 else ''
            member = raw_parts[4].strip() if len(raw_parts) > 4 else ''
            task_0_10 = raw_parts[5].strip() if len(raw_parts) > 5 else ''
            task_10_20 = raw_parts[6].strip() if len(raw_parts) > 6 else ''
            task_20_plus = raw_parts[7].strip() if len(raw_parts) > 7 else ''
            
            if group in ['A', 'B']:
                # 1队2队格式: 有AB分组
                if member and not any(kw in member for kw in tactic_kw):
                    teams_data[current_team][f'{group}_members'].append(member)
                teams_data[current_team][f'{group}_tasks']['0-10'] = task_0_10
                teams_data[current_team][f'{group}_tasks']['10-20'] = task_10_20
                teams_data[current_team][f'{group}_tasks']['20+'] = task_20_plus
            else:
                # 3-6队格式: 无AB分组，parts[2]=成员，parts[3]=时段1(10-20), parts[4]=时段2(20+)
                if member and not any(kw in member for kw in tactic_kw) and not member.isdigit():
                    teams_data[current_team]['A_members'].append(member)
                # 根据队类型取时段索引：1-2队用[4/5/6]，3-6队用[3/4/5]
                if len(parts) > 3 and parts[2] in ['A', 'B']:
                    # 1队2队: [4]=0-10, [5]=10-20, [6]=20+
                    task_0_10 = parts[4].strip() if len(parts) > 4 else ''
                    task_10_20 = parts[5].strip() if len(parts) > 5 else ''
                    task_20_plus = parts[6].strip() if len(parts) > 6 else ''
                else:
                    # 3-6队: [3]=10-20(col5,含"大粮仓"), [4]=20+(col6), [5]=col7
                    task_0_10 = ''
                    task_10_20 = parts[3].strip() if len(parts) > 3 else ''
                    task_20_plus = parts[4].strip() if len(parts) > 4 else ''
                teams_data[current_team]['A_tasks']['0-10'] = task_0_10
                teams_data[current_team]['A_tasks']['10-20'] = task_10_20
                teams_data[current_team]['A_tasks']['20+'] = task_20_plus
        
        # A/B分组行: raw_parts[3] = 'A' 或 'B'
        elif len(raw_parts) > 4 and raw_parts[3].strip() in ['A', 'B'] and current_team:
            group = raw_parts[3].strip()
            member = raw_parts[4].strip() if len(raw_parts) > 4 else ''
            task_10_20 = raw_parts[6].strip() if len(raw_parts) > 6 else ''
            if member and not any(kw in member for kw in tactic_kw):
                teams_data[current_team][f'{group}_members'].append(member)
            # B组第一行可能带任务(如 col6='B队驻守')
            if group == 'B' and task_10_20 and any(kw in task_10_20 for kw in tactic_kw):
                teams_data[current_team]['B_tasks']['10-20'] = task_10_20
        
        # 无队号/分组标记的行，可能是队员行(3-6队后续队员)
        elif current_team and parts[0] not in ['队伍', '队长', '替补']:
            # 3-6队后续队员在 col4
            member = raw_parts[4].strip() if len(raw_parts) > 4 else parts[0]
            if member and not any(kw in member for kw in tactic_kw) and not member.isdigit():
                if not re.match(r'^[\d\s]+$', member):
                    teams_data[current_team]['A_members'].append(member)
        
        # 替补行
        if parts[0] == '替补' and len(parts) > 1:
            bench_text = parts[1]
            bench_members = [m.strip() for m in bench_text.split('、') if m.strip()]
            # 替补任务在 raw_parts[6] (10-20分钟列)
            bench_task = raw_parts[6].strip() if len(raw_parts) > 6 else ''
    
    return {'teams': teams_data, 'bench': bench_members, 'bench_task': bench_task if 'bench_task' in dir() else ''}


def extract_j_members(teams_data: dict) -> list:
    """提取J列成员（队长+队员，去重）"""
    members = []
    seen = set()
    
    for team in ['1队', '2队', '3队', '4队', '5队', '6队']:
        if team in teams_data:
            data = teams_data[team]
            # 队长
            if data['captain'] and data['captain'] not in seen:
                seen.add(data['captain'])
                members.append(data['captain'])
            # A组成员
            for m in data['A_members']:
                for name in m.split('、'):
                    name = name.strip()
                    if name and name not in seen:
                        seen.add(name)
                        members.append(name)
            # B组成员
            for m in data['B_members']:
                for name in m.split('、'):
                    name = name.strip()
                    if name and name not in seen:
                        seen.add(name)
                        members.append(name)
    
    return members


def expand_members(members: list) -> list:
    """展开顿号分隔的成员名"""
    result = []
    seen = set()
    for m in members:
        if '、' in m:
            for name in m.split('、'):
                name = name.strip()
                if name and name not in seen:
                    seen.add(name)
                    result.append(name)
        else:
            if m and m not in seen:
                seen.add(m)
                result.append(m)
    return result


@app.route('/api/sections')
def get_sections():
    """从默认官渡表提取可用分组列表"""
    try:
        with open(GUANDU_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        sections = []
        for m in re.finditer(r'^## (.+)$', content, re.MULTILINE):
            title = m.group(1).strip()
            if '排名' not in title:
                sections.append(title)
        return jsonify({'success': True, 'sections': sections})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/init')
def init_api():
    """一次性初始化：返回sections、stats、demo数据和匹配结果"""
    section = request.args.get('section', '团一')
    print(f'[INIT] section={repr(section)}')
    result = {'success': True}
    # 1. Sections
    try:
        with open(GUANDU_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        sections = []
        for m in re.finditer(r'^## (.+)$', content, re.MULTILINE):
            title = m.group(1).strip()
            if '排名' not in title:
                sections.append(title)
        result['sections'] = sections
    except Exception as e:
        app.logger.error(f'init sections error: {e}')
        result['sections'] = []

    # 2. Stats
    try:
        with open(STATS_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        stats = parse_stats_table(content)
        sort_fields = detect_sort_fields(stats)
        result['stats'] = stats
        result['statsCount'] = len(stats)
        result['sortFields'] = sort_fields
    except Exception as e:
        app.logger.error(f'init stats error: {e}')
        result['stats'] = {}
        result['statsCount'] = 0
        result['sortFields'] = []

    # 3. Demo data (members from guandu)
    try:
        with open(GUANDU_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        data = parse_guandu_table(content, section)
        print(f'[INIT] parse_guandu_table返回teams={list(data["teams"].keys())}')
        members = extract_j_members(data['teams'])
        bench = data['bench']
        bench_task = data.get('bench_task', '')
        result['members'] = members
        result['bench'] = bench
        result['bench_task'] = bench_task
        result['teams'] = data['teams']
    except Exception as e:
        app.logger.error(f'init demo error: {e}')
        result['members'] = []
        result['bench'] = []
        result['teams'] = {}

    # 4. Auto match
    if result['members'] and result['stats']:
        match_results = []
        for m in result['members']:
            match_results.append(match_member(m, result['stats']))
        name_map = {}
        unmatched_list = []
        for r in match_results:
            if r['status'] != 'not_found':
                name_map[r['original']] = r['matched']
            else:
                unmatched_list.append(r)
        result['nameMap'] = name_map
        result['unmatched'] = unmatched_list
        result['matchDone'] = True
    else:
        result['nameMap'] = {}
        result['unmatched'] = []
        result['matchDone'] = False

    # Force Content-Length so axios doesn't hang on chunked transfer
    resp = make_response(json.dumps(result, ensure_ascii=False))
    resp.headers['Content-Type'] = 'application/json; charset=utf-8'
    resp.headers['Content-Length'] = len(resp.data)
    return resp


@app.route('/api/demo_data')
def get_demo_data():
    """从默认官渡表提取全部成员作为测试数据"""
    section = request.args.get('section', '团一')
    app.logger.info(f'demo_data: section={section}')
    try:
        with open(GUANDU_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        data = parse_guandu_table(content, section)
        members = extract_j_members(data['teams'])
        bench = data['bench']
        app.logger.info(f'demo_data: {len(members)} members, {len(bench)} bench')
        return jsonify({
            'success': True,
            'members': members,
            'bench': bench,
            'teams': data['teams'],
            'section': section
        })
    except Exception as e:
        app.logger.error(f'demo_data error: {e}')
        return jsonify({'success': False, 'error': str(e)})


def match_member(name: str, stats: dict) -> dict:
    """匹配成员名到统计表"""
    # 直接匹配
    if name in stats:
        return {'original': name, 'matched': name, 'hp': stats[name]['hp'], 'status': 'exact'}
    
    # 顿号/逗号转换匹配
    for stat_name in stats.keys():
        if name.replace('、', '丶') == stat_name or name.replace('、', ',') == stat_name:
            return {'original': name, 'matched': stat_name, 'hp': stats[stat_name]['hp'], 'status': 'exact'}
    
    # 部分匹配
    for stat_name in stats.keys():
        if name in stat_name or stat_name in name:
            if name.replace('、', '丶') in stat_name or stat_name.replace('丶', '、') in name:
                return {'original': name, 'matched': stat_name, 'hp': stats[stat_name]['hp'], 'status': 'partial'}
    
    # 未找到
    return {'original': name, 'matched': None, 'hp': 0, 'status': 'not_found'}


def assign_members(members: list, stats: dict, name_map: dict, threshold: float = 900, seed: int = 42, manual_captains: dict = None, sort_by: str = 'hp') -> dict:
    """蛇形分配成员到B列和D列
    
    官渡表格式：
    - 1队2队: 各8队员(A组4+B组4)
    - 3-6队: 各2队员(无分组)
    
    蛇形分配策略：
    - B1=第1名, B2=第2名, B3-B6按序分配
    - 高值成员蛇形分配到1队2队的A/B组
    - 低值成员蛇形分配到3-6队
    """
    random.seed(seed)
    manual_captains = manual_captains or {}
    sort_key = sort_by if sort_by in ('hp', 'total', 'power') else 'hp'
    
    # 匹配成员并收集所有属性
    matched_members = []
    for m in members:
        mapped_name = name_map.get(m, m)
        if mapped_name and mapped_name in stats:
            hp = stats[mapped_name].get('hp', 0)
            total = stats[mapped_name].get('total', 0)
            power = stats[mapped_name].get('power', 0)
        else:
            hp = total = power = 0
        matched_members.append({
            'original': m, 'mapped': mapped_name if mapped_name else m,
            'hp': hp, 'total': total, 'power': power,
        })
    
    sorted_members = sorted(matched_members, key=lambda x: x[sort_key], reverse=True)
    b_assign = {}
    d_assign = {}
    
    # 1. 手动指定队长
    manual_used = set()
    for pos in ['B1', 'B2', 'B3', 'B4', 'B5', 'B6']:
        if manual_captains.get(pos):
            manual_name = manual_captains[pos]
            for m in members:
                if m == manual_name:
                    mapped = name_map.get(m, m)
                    s = stats.get(mapped, {})
                    found = {'original': m, 'mapped': mapped or m,
                             'hp': s.get('hp', 0), 'total': s.get('total', 0), 'power': s.get('power', 0)}
                    b_assign[pos] = found
                    manual_used.add(manual_name)
                    break
    
    auto_pool = [m for m in sorted_members if m['original'] not in manual_used]
    
    # D列位置映射（与官渡表格式一致）
    # 1队: A组D1-D4, B组D9-D12 (共8队员)
    # 2队: A组D5-D8, B组D13-D16 (共8队员)
    # 3队: D17,D18 / 4队: D19,D20 / 5队: D21,D22 / 6队: D23,D24
    
    # 2. 自动分配队长
    # B1,B2拿最强队长，B3-B6按序分配
    pool_idx = 0
    for pos in ['B1', 'B2']:
        if pos not in b_assign:
            if pool_idx < len(auto_pool):
                b_assign[pos] = auto_pool[pool_idx]
                pool_idx += 1
    for pos in ['B3', 'B4', 'B5', 'B6']:
        if pos not in b_assign:
            if pool_idx < len(auto_pool):
                b_assign[pos] = auto_pool[pool_idx]
                pool_idx += 1
    
    # 3. D列蛇形分配（所有非队长成员）
    assigned_set = set()
    for pos_data in list(b_assign.values()):
        assigned_set.add(id(pos_data))
    remaining = [m for m in auto_pool if id(m) not in assigned_set]
    
    # 按sort_key降序排，同分值随机打散
    random.shuffle(remaining)
    remaining.sort(key=lambda x: x[sort_key], reverse=True)
    
    # 蛇形分配到所有D位置
    # 策略：高值→1队2队(A/B组) → 3-6队，蛇形保证各队实力均衡
    # 1队A→2队A→1队A→2队A→...→1队B→2队B→...→3队→4队→5队→6队→6队→5队→4队→3队
    all_d_positions = [
        # 1队2队A/B组(16个位置)
        'D1','D5','D2','D6','D3','D7','D4','D8',       # A组蛇形
        'D9','D13','D10','D14','D11','D15','D12','D16', # B组蛇形
        # 3-6队(8个位置)
        'D17','D19','D21','D23','D24','D22','D20','D18'  # 低队蛇形
    ]
    for i, pos in enumerate(all_d_positions):
        if i < len(remaining):
            d_assign[pos] = remaining[i]
    
    return {
        'b_assign': b_assign,
        'd_assign': d_assign,
        'sorted': sorted_members,
        'unmatched': [m for m in matched_members if m[sort_key] == 0],
    }


# ============ 路由 ============

@app.route('/')
def index():
    """主页"""
    with open(os.path.join(os.path.dirname(__file__), 'templates', 'index.html'), 'r', encoding='utf-8') as f:
        html = f.read()
    return html



def detect_sort_fields(stats: dict) -> list:
    """检测属性表中哪些排序字段有数据，返回可用字段列表
    
    检查逻辑：如果某字段在任一成员中有非零值，则认为该字段可用
    """
    fields = []
    has_hp = False
    has_total = False
    has_power = False
    for v in stats.values():
        if v.get('hp', 0) > 0:
            has_hp = True
        if v.get('total', 0) > 0:
            has_total = True
        if v.get('power', 0) > 0:
            has_power = True
    if has_hp:
        fields.append({'key': 'hp', 'label': '步兵生命值'})
    if has_total:
        fields.append({'key': 'total', 'label': '六维属性总和'})
    if has_power:
        fields.append({'key': 'power', 'label': '战力'})
    return fields


@app.route('/api/load_stats')
def load_stats():
    """加载统计表数据（从默认文件）"""
    try:
        with open(STATS_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        stats = parse_stats_table(content)
        # 动态检测可用的排序字段
        fields = detect_sort_fields(stats)
        return jsonify({'success': True, 'stats': stats, 'count': len(stats), 'sort_fields': fields})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/upload_stats', methods=['POST'])
def upload_stats():
    """上传属性表文件"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有上传文件'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': '不支持的文件格式，请上传 .xlsx, .csv 或 .md 文件'})
    
    try:
        filename = secure_filename(file.filename)
        ext = filename.rsplit('.', 1)[1].lower()
        
        # 根据文件类型解析
        stats = {}
        
        if ext == 'md':
            content = file.read().decode('utf-8')
            stats = parse_stats_table(content)
        
        elif ext == 'csv':
            content = file.read().decode('utf-8')
            stats = parse_stats_csv(content)
        
        elif ext == 'xlsx':
            stats = parse_stats_xlsx(file)
        
        if not stats:
            return jsonify({'success': False, 'error': '未能解析到有效数据，请检查文件格式'})
        
        fields = detect_sort_fields(stats)
        return jsonify({
            'success': True, 
            'stats': stats, 
            'count': len(stats),
            'filename': filename,
            'sort_fields': fields
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失败: {str(e)}'})


def parse_stats_csv(content: str) -> dict:
    """解析CSV格式的属性表"""
    import csv
    from io import StringIO
    
    stats = {}
    reader = csv.reader(StringIO(content))
    header = None
    hp_idx = None
    total_idx = None
    power_idx = None
    
    for row in reader:
        if not header:
            header = row
            for i, col in enumerate(row):
                if '步兵生命值' in col or '生命值' in col or 'HP' in col.upper():
                    hp_idx = i
                elif '六维属性总和' in col or '六维总和' in col:
                    total_idx = i
                elif '战力' in col:
                    power_idx = i
            continue
        
        if len(row) < 2:
            continue
        
        name = row[1].strip() if len(row) > 1 else row[0].strip()
        if not name or name.isdigit():
            continue
        
        hp = 0
        total = 0
        power = 0
        
        if hp_idx is not None and len(row) > hp_idx:
            try: hp = float(row[hp_idx])
            except: pass
        if total_idx is not None and len(row) > total_idx:
            try: total = float(row[total_idx])
            except: pass
        if power_idx is not None and len(row) > power_idx:
            try: power = float(row[power_idx])
            except: pass
        
        # 如果HP异常，尝试智能检测
        if hp > 100000 or (hp == 0 and total == 0 and power == 0):
            for val in row:
                try:
                    v = float(val)
                    if 100 <= v <= 5000:
                        hp = v
                        break
                except:
                    pass
        
        if hp > 100000:
            hp = 0
        
        stats[name] = {'hp': hp, 'total': total, 'power': power}
    
    return stats


def parse_stats_xlsx(file) -> dict:
    """解析Excel格式的属性表"""
    wb = openpyxl.load_workbook(file, read_only=True)
    ws = wb.active
    
    stats = {}
    header = None
    hp_idx = None
    total_idx = None
    power_idx = None
    
    for row in ws.iter_rows(values_only=True):
        if not header:
            header = row
            for i, col in enumerate(row):
                if col and ('步兵生命值' in str(col) or '生命值' in str(col) or 'HP' in str(col).upper()):
                    hp_idx = i
                elif col and ('六维属性总和' in str(col) or '六维总和' in str(col)):
                    total_idx = i
                elif col and '战力' in str(col):
                    power_idx = i
            continue
        
        if len(row) < 2:
            continue
        
        name = str(row[1]).strip() if len(row) > 1 and row[1] else str(row[0]).strip() if row[0] else ''
        if not name or name.isdigit():
            continue
        
        hp = 0
        total = 0
        power = 0
        
        if hp_idx is not None and len(row) > hp_idx and row[hp_idx]:
            try: hp = float(row[hp_idx])
            except: pass
        if total_idx is not None and len(row) > total_idx and row[total_idx]:
            try: total = float(row[total_idx])
            except: pass
        if power_idx is not None and len(row) > power_idx and row[power_idx]:
            try: power = float(row[power_idx])
            except: pass
        
        if hp > 100000 or (hp == 0 and total == 0 and power == 0):
            for val in row:
                try:
                    v = float(val) if val else 0
                    if 100 <= v <= 5000:
                        hp = v
                        break
                except:
                    pass
        
        if hp > 100000:
            hp = 0
        
        stats[name] = {'hp': hp, 'total': total, 'power': power}
    
    wb.close()
    return stats


@app.route('/api/load_guandu')
def load_guandu():
    """加载官渡表数据（从默认文件）"""
    section = request.args.get('section', '团一')
    try:
        with open(GUANDU_FILE, 'r', encoding='utf-8') as f:
            content = f.read()
        data = parse_guandu_table(content, section)
        members = extract_j_members(data['teams'])
        return jsonify({
            'success': True,
            'teams': data['teams'],
            'bench': data['bench'],
            'members': members,
            'bench_task': data.get('bench_task', ''),
            'section': section,
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/upload_guandu', methods=['POST'])
def upload_guandu():
    """上传官渡名单文件"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '没有上传文件'})
    
    file = request.files['file']
    section = request.form.get('section', '团一')
    
    if file.filename == '':
        return jsonify({'success': False, 'error': '没有选择文件'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': '不支持的文件格式，请上传 .xlsx, .csv 或 .md 文件'})
    
    try:
        filename = secure_filename(file.filename)
        ext = filename.rsplit('.', 1)[1].lower()
        
        teams_data = {}
        bench = []
        members = []
        
        if ext == 'md':
            content = file.read().decode('utf-8')
            data = parse_guandu_table(content, section)
            teams_data = data['teams']
            bench = data['bench']
            members = extract_j_members(teams_data)
        
        elif ext == 'csv':
            content = file.read().decode('utf-8')
            members, bench = parse_guandu_csv(content)
        
        elif ext == 'xlsx':
            members, bench = parse_guandu_xlsx(file)
            app.logger.info(f'xlsx解析: members={len(members)}, bench={len(bench)}, first5={members[:5]}')
        
        if not members and not teams_data:
            return jsonify({'success': False, 'error': '未能解析到有效数据，请检查文件格式'})
        
        return jsonify({
            'success': True,
            'teams': teams_data,
            'bench': bench,
            'members': members,
            'section': section,
            'filename': filename
        })
    
    except Exception as e:
        return jsonify({'success': False, 'error': f'解析失败: {str(e)}'})


def parse_guandu_csv(content: str) -> tuple:
    """解析CSV格式的官渡名单，返回 (members, bench)"""
    import csv
    from io import StringIO
    
    members = []
    bench = []
    
    reader = csv.reader(StringIO(content))
    for row in reader:
        for val in row:
            val = str(val).strip()
            if not val or val.isdigit():
                continue
            if '替补' in val:
                # 后面的值可能是替补名单
                continue
            if val in ['队伍', '队长', '队员', 'A', 'B', '分组']:
                continue
            # 展开顿号分隔的名字
            for name in val.replace(',', '、').split('、'):
                name = name.strip()
                if name and name not in members:
                    members.append(name)
    
    return members, bench


def parse_guandu_xlsx(file) -> tuple:
    """解析Excel格式的官渡名单，返回 (members, bench)"""
    wb = openpyxl.load_workbook(file, read_only=True)
    ws = wb.active
    
    members = []
    bench = []
    
    for row in ws.iter_rows(values_only=True):
        for val in row:
            val = str(val).strip() if val else ''
            if not val or val.isdigit():
                continue
            if val in ['队伍', '队长', '队员', 'A', 'B', '分组']:
                continue
            for name in val.replace(',', '、').split('、'):
                name = name.strip()
                if name and name not in members:
                    members.append(name)
    
    wb.close()
    return members, bench


@app.route('/api/match_members', methods=['POST'])
def match_members_api():
    """匹配成员"""
    app.logger.info('match_members 被调用')
    data = request.json
    members = data.get('members', [])
    stats = data.get('stats', {})
    app.logger.info(f'match_members: {len(members)} members, {len(stats)} stats')
    
    results = []
    for m in members:
        result = match_member(m, stats)
        results.append(result)
    
    unmatched = [r for r in results if r['status'] == 'not_found']
    matched = [r for r in results if r['status'] != 'not_found']
    
    return jsonify({
        'success': True,
        'results': results,
        'matched': matched,
        'unmatched': unmatched,
    })


@app.route('/api/assign', methods=['POST'])
def assign_api():
    """分配成员"""
    data = request.json
    members = data.get('members', [])
    stats = data.get('stats', {})
    name_map = data.get('name_map', {})
    threshold = data.get('threshold', 900)
    seed = data.get('seed', 42)
    manual_captains = data.get('manual_captains', {})
    sort_by = data.get('sort_by', 'hp')  # 'hp' 或 'total'
    
    result = assign_members(members, stats, name_map, threshold, seed, manual_captains, sort_by)
    return jsonify({
        'success': True,
        'b_assign': result['b_assign'],
        'd_assign': result['d_assign'],
        'sorted': result['sorted'],
        'unmatched': result['unmatched'],
    })


@app.route('/api/export', methods=['POST'])
def export_api():
    """导出Excel - 将分配结果填入官渡表格式"""
    data = request.json
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "官渡分配表"
    
    # 样式
    header_font = Font(bold=True, size=12)
    captain_font = Font(bold=True, size=11, color='1a1a2e')
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    header_fill = PatternFill(start_color='CCE5FF', end_color='CCE5FF', fill_type='solid')
    captain_fill = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
    
    b_assign = data.get('b_assign', {})
    d_assign = data.get('d_assign', {})
    section = data.get('section', '团一')
    sort_by = data.get('sort_by', 'hp')
    teams_data = data.get('teams_data', {})
    bench_task = data.get('bench_task', '')
    
    # 队伍到D位置的映射
    # 1队2队: A组4人 + B组4人 = 8行队员
    # 3-6队: 无分组, 2行队员
    team_d_map = {
        1: {'a': ['D1','D2','D3','D4'], 'b': ['D9','D10','D11','D12']},
        2: {'a': ['D5','D6','D7','D8'], 'b': ['D13','D14','D15','D16']},
        3: {'members': ['D17','D18']},
        4: {'members': ['D19','D20']},
        5: {'members': ['D21','D22']},
        6: {'members': ['D23','D24']},
    }
    
    # 写入表头（不含HP/排序值列）
    headers = ['队伍', 'B列(队长)', '分组', 'D列(队员)', '0-10分钟', '10-20分钟', '20分钟以后', '备注']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.alignment = center_align
        cell.fill = header_fill
        cell.border = thin_border
    
    def set_cell(ws, row, col, value='', font=None, fill=None, align=None, border=None):
        """写入单元格并设置样式"""
        cell = ws.cell(row=row, column=col, value=value)
        if font: cell.font = font
        if fill: cell.fill = fill
        if align: cell.alignment = align
        if border: cell.border = border
        return cell

    def apply_border_range(ws, r1, c1, r2, c2):
        """给范围内的所有单元格加边框"""
        for r in range(r1, r2+1):
            for c in range(c1, c2+1):
                ws.cell(row=r, column=c).border = thin_border

    row = 2
    # 3-5队跨队合并追踪
    group_35_start_row = None
    group_35_end_row = None
    group_35_gh_content = ''

    for team_num in range(1, 7):
        team_name = f'{team_num}队'
        b_key = f'B{team_num}'
        captain_data = b_assign.get(b_key, {})
        captain = captain_data.get('original', '') if captain_data else ''
        captain_score = captain_data.get(sort_by, 0) if captain_data else 0
        tmap = team_d_map.get(team_num, {})
        
        if team_num <= 2:
            # === 1队2队: 8行 — A组4队员 + B组4队员 ===
            a_keys = tmap.get('a', [])
            b_keys = tmap.get('b', [])
            tdata = teams_data.get(team_name, {})
            a_tasks = tdata.get('A_tasks', {})
            b_tasks = tdata.get('B_tasks', {})
            
            team_start_row = row
            a_start_row = row
            
            # --- 写入A组4行 ---
            for i, dk in enumerate(a_keys):
                d_data = d_assign.get(dk, {})
                m_name = d_data.get('original', '') if d_data else ''
                m_score = d_data.get(sort_by, 0) if d_data else 0
                set_cell(ws, row, 3, 'A', align=center_align, border=thin_border)
                set_cell(ws, row, 4, m_name, border=thin_border)
                set_cell(ws, row, 8, '', border=thin_border)
                row += 1
            a_end_row = row - 1
            
            b_start_row = row
            
            # --- 写入B组4行 ---
            for i, dk in enumerate(b_keys):
                d_data = d_assign.get(dk, {})
                m_name = d_data.get('original', '') if d_data else ''
                set_cell(ws, row, 3, 'B', align=center_align, border=thin_border)
                set_cell(ws, row, 4, m_name, border=thin_border)
                set_cell(ws, row, 8, '', border=thin_border)
                row += 1
            b_end_row = row - 1
            team_end_row = row - 1
            
            # --- 合并单元格：队伍名(全队8行) ---
            if team_end_row > team_start_row:
                ws.merge_cells(start_row=team_start_row, start_column=1, end_row=team_end_row, end_column=1)
            set_cell(ws, team_start_row, 1, team_name, font=Font(bold=True, size=12), align=center_align, border=thin_border)
            apply_border_range(ws, team_start_row, 1, team_end_row, 1)
            
            # --- 合并单元格：队长(全队8行) ---
            if team_end_row > team_start_row:
                ws.merge_cells(start_row=team_start_row, start_column=2, end_row=team_end_row, end_column=2)
            set_cell(ws, team_start_row, 2, captain, font=captain_font, fill=captain_fill, align=center_align, border=thin_border)
            apply_border_range(ws, team_start_row, 2, team_end_row, 2)
            
            # --- 合并单元格：0-10分钟任务(全队8行) ---
            task_010 = a_tasks.get('0-10', '')
            if team_end_row > team_start_row:
                ws.merge_cells(start_row=team_start_row, start_column=5, end_row=team_end_row, end_column=5)
            set_cell(ws, team_start_row, 5, task_010, align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
            apply_border_range(ws, team_start_row, 5, team_end_row, 5)
            
            # --- 合并单元格：A组 10-20分钟(4行) ---
            if a_end_row > a_start_row:
                ws.merge_cells(start_row=a_start_row, start_column=6, end_row=a_end_row, end_column=6)
            set_cell(ws, a_start_row, 6, a_tasks.get('10-20', ''), align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
            apply_border_range(ws, a_start_row, 6, a_end_row, 6)
            
            # --- 合并单元格：A组 20+分钟(4行) ---
            if a_end_row > a_start_row:
                ws.merge_cells(start_row=a_start_row, start_column=7, end_row=a_end_row, end_column=7)
            set_cell(ws, a_start_row, 7, a_tasks.get('20+', ''), align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
            apply_border_range(ws, a_start_row, 7, a_end_row, 7)
            
            # --- 合并单元格：B组 10-20分钟(4行) ---
            if b_end_row > b_start_row:
                ws.merge_cells(start_row=b_start_row, start_column=6, end_row=b_end_row, end_column=6)
            set_cell(ws, b_start_row, 6, b_tasks.get('10-20', ''), align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
            apply_border_range(ws, b_start_row, 6, b_end_row, 6)
            
            # --- 合并单元格：B组 20+分钟(4行) ---
            if b_end_row > b_start_row:
                ws.merge_cells(start_row=b_start_row, start_column=7, end_row=b_end_row, end_column=7)
            set_cell(ws, b_start_row, 7, b_tasks.get('20+', ''), align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
            apply_border_range(ws, b_start_row, 7, b_end_row, 7)
        
        else:
            # === 3-6队: 队长+队员 ===
            member_keys = tmap.get('members', [])
            tdata = teams_data.get(team_name, {})
            a_tasks = tdata.get('A_tasks', {})
            
            # 收集有数据的队员
            filled_members = []
            for mk in member_keys:
                d_data = d_assign.get(mk, {})
                m_name = d_data.get('original', '') if d_data else ''
                if m_name:
                    filled_members.append((m_name, d_data.get(sort_by, 0) if d_data else 0))
            
            team_start_row = row
            # 写入队员行
            for i, (m_name, m_score) in enumerate(filled_members):
                set_cell(ws, row, 4, m_name, border=thin_border)
                set_cell(ws, row, 8, '', border=thin_border)
                row += 1
            team_end_row = row - 1
            
            # 如果没有队员，至少写一行
            if not filled_members:
                set_cell(ws, row, 4, '', border=thin_border)
                set_cell(ws, row, 8, '', border=thin_border)
                team_end_row = row
                row += 1
            
            # --- 合并单元格：队伍名 ---
            if team_end_row > team_start_row:
                ws.merge_cells(start_row=team_start_row, start_column=1, end_row=team_end_row, end_column=1)
            set_cell(ws, team_start_row, 1, team_name, font=Font(bold=True, size=12), align=center_align, border=thin_border)
            apply_border_range(ws, team_start_row, 1, team_end_row, 1)
            
            # --- 合并单元格：队长 ---
            if team_end_row > team_start_row:
                ws.merge_cells(start_row=team_start_row, start_column=2, end_row=team_end_row, end_column=2)
            set_cell(ws, team_start_row, 2, captain, font=captain_font, fill=captain_fill, align=center_align, border=thin_border)
            apply_border_range(ws, team_start_row, 2, team_end_row, 2)
            
            # --- 任务列 ---
            task_010 = a_tasks.get('0-10', '')
            task_1020 = a_tasks.get('10-20', '')
            task_20plus = a_tasks.get('20+', '')
            
            if team_num <= 5:
                # === 3-5队 ===
                # E列(0-10分钟)：每队的任务相同，队内垂直合并
                f_content = task_1020  # 实际内容来自10-20分钟字段
                if f_content:
                    if team_end_row > team_start_row:
                        ws.merge_cells(start_row=team_start_row, start_column=5, end_row=team_end_row, end_column=5)
                    set_cell(ws, team_start_row, 5, f_content, align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
                    apply_border_range(ws, team_start_row, 5, team_end_row, 5)
                
                # F:G列(10-20分钟+20分钟以后)：3-5队任务一样，记录起止行，循环结束后统一跨队合并
                if team_num == 3:
                    group_35_start_row = team_start_row
                    group_35_gh_content = task_20plus  # 使用3队的20+任务内容
                if team_num == 5:
                    group_35_end_row = team_end_row
            
            else:
                # === 6队 ===
                # E:F合并(0-10+10-20分钟)，内容=task_1020
                ws.merge_cells(start_row=team_start_row, start_column=5, end_row=team_end_row, end_column=6)
                set_cell(ws, team_start_row, 5, task_1020, align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
                apply_border_range(ws, team_start_row, 5, team_end_row, 6)
                
                # G列(20分钟以后)合并，内容=task_20plus
                if team_end_row > team_start_row:
                    ws.merge_cells(start_row=team_start_row, start_column=7, end_row=team_end_row, end_column=7)
                set_cell(ws, team_start_row, 7, task_20plus, align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
                apply_border_range(ws, team_start_row, 7, team_end_row, 7)
    
    # --- 3-5队 G:H跨队合并（10-20分钟+20分钟以后任务相同） ---
    if group_35_start_row and group_35_end_row:
        ws.merge_cells(start_row=group_35_start_row, start_column=6, end_row=group_35_end_row, end_column=7)
        set_cell(ws, group_35_start_row, 6, group_35_gh_content, align=Alignment(wrap_text=True, vertical='center'), border=thin_border)
        apply_border_range(ws, group_35_start_row, 6, group_35_end_row, 7)

    # 调整列宽
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 25
    ws.column_dimensions['H'].width = 15
    
    # 候补人员
    bench_list = data.get('bench_list', [])
    if bench_list:
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=1, value='候补人员')
        cell.font = Font(bold=True, size=12, color='FFFFFF')
        cell.fill = PatternFill(start_color='6c757d', end_color='6c757d', fill_type='solid')
        cell.alignment = center_align
        for col in range(1, 9):
            ws.cell(row=row, column=col).border = thin_border
        row += 1
        # 候补人员名单合并到 B:D
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        bench_names = '、'.join(bench_list)
        cell_names = ws.cell(row=row, column=2, value=bench_names)
        cell_names.alignment = Alignment(wrap_text=True)
        cell_names.border = thin_border
        ws.cell(row=row, column=1).border = thin_border
        for col in range(5, 9):
            ws.cell(row=row, column=col).border = thin_border
        # 替补任务描述移到同一行 E:G
        if bench_task:
            ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
            cell_task = ws.cell(row=row, column=5, value=bench_task)
            cell_task.alignment = Alignment(wrap_text=True)
            cell_task.border = thin_border
        row += 1
    
    # 底部备注
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    note = ws.cell(row=row, column=1, value='*各自队伍拿下首占后，可根据对手的动向进行调整，随机应变。*')
    note.alignment = Alignment(horizontal='center', wrap_text=True)
    note.font = Font(italic=True, size=10, color='666666')
    
    # 保存
    filename = f'官渡分配表_{section}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    filepath = os.path.join(os.path.dirname(__file__), 'downloads', filename)
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    wb.save(filepath)
    
    return jsonify({'success': True, 'filename': filename, 'path': filepath})


@app.route('/download/<filename>')
def download(filename):
    """下载文件"""
    filepath = os.path.join(os.path.dirname(__file__), 'downloads', filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return 'File not found', 404


# ============ 主程序 ============

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5001, threaded=True)
