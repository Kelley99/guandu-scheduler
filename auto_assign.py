#!/usr/bin/env python3
"""
官渡自动分配成员工具
根据步兵生命值自动分配成员到B列和D列
"""

import re
import random
import json
from typing import Dict, List, Tuple, Optional

# ============ 配置 ============
INPUT_GUANDU = "knowledge-base/凌霄官渡26.md"
INPUT_STATS = "knowledge-base/凌霄数据统计表26.3.30.md"
RANDOM_SEED = 42  # 固定种子以便复现

# ============ 名字映射 ============
NAME_MAPPING = {
    '破晓、小妹': '破晓丶小妹',
    '破晓、轩辕兮瞳': '破晓丶轩辕兮瞳',
    '私人专属、七': '私人专属丶七',
    '破晓、偏恋你的香': '破晓丶偏恋你的香',
    '轮回': '轮丶回',
    '叫霸霸': '霸霸',
    '小妹': '破晓丶小妹',
    '轩辕兮瞳': '破晓丶轩辕兮瞳',
    '七': '私人专属丶七',
    '偏恋你的香': '破晓丶偏恋你的香',
}

# ============ 解析函数 ============

def parse_stats_table(content: str) -> Dict[str, float]:
    """解析统计表，返回 {成员名: 步兵生命值}"""
    stats = {}
    for line in content.split('\n'):
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|')]
        parts = [p for p in parts if p]
        # 序号 | 成员名称 | 集结加成 | 步兵防御 | 步兵生命值 | ...
        if len(parts) >= 5 and parts[0].isdigit():
            name = parts[1]
            try:
                stats[name] = float(parts[4])
            except:
                pass
    return stats

def map_name(name: str, stats: Dict[str, float]) -> Optional[str]:
    """映射成员名到统计表中的名字"""
    if name in stats:
        return name
    if name in NAME_MAPPING:
        mapped = NAME_MAPPING[name]
        if mapped in stats:
            return mapped
    # 模糊匹配
    for stat_name in stats.keys():
        if name in stat_name:
            return stat_name
    return None

def parse_guandu_table(content: str, section: str) -> Tuple[Dict, List[str]]:
    """解析官渡表，返回 (各队数据, 候补名单)"""
    # 提取指定团的数据
    pattern = rf'## {section}[\s\S]*?(?=## 团|$)'
    match = re.search(pattern, content)
    if not match:
        return {}, []
    
    section_content = match.group(0)
    lines = section_content.split('\n')
    
    tactic_kw = ['0-10分钟', '10-20分钟', '20分钟以后', '拿下', '驻守', '集结', 
                 '粮仓', '乌巢', '官渡', '霹雳', '锱重', '驻防', '采集', '远程',
                 '不动', '应变', '必要时', '跟随', '工匠坊', '兵器坊', '首占']
    
    current_team = None
    teams_data = {}
    bench_members = []
    
    for line in lines:
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|')]
        parts = [p for p in parts if p]
        
        if not parts or all(re.match(r'^-+$', p) for p in parts):
            continue
        
        # 队号行
        if re.match(r'^\d+队$', parts[0]):
            current_team = parts[0]
            teams_data[current_team] = {'captain': None, 'A_members': [], 'B_members': []}
            if len(parts) > 1 and parts[1]:
                teams_data[current_team]['captain'] = parts[1]
            # D列(parts[3])是队员
            if len(parts) > 3 and parts[3]:
                member = parts[3]
                if not any(kw in member for kw in tactic_kw):
                    if len(parts) > 2 and parts[2] == 'A':
                        teams_data[current_team]['A_members'].append(member)
                    else:
                        teams_data[current_team]['B_members'].append(member)
        
        # A/B分组行
        elif parts[0] in ['A', 'B'] and current_team:
            group = parts[0]
            if len(parts) > 1 and parts[1]:
                member = parts[1]
                if not any(kw in member for kw in tactic_kw):
                    if group == 'A':
                        teams_data[current_team]['A_members'].append(member)
                    else:
                        teams_data[current_team]['B_members'].append(member)
        
        # 替补行
        elif parts[0] == '替补' and len(parts) > 1:
            bench_text = parts[1]
            bench_members = [m.strip() for m in bench_text.split('、') if m.strip()]
    
    return teams_data, bench_members

def extract_j_members(teams_data: Dict) -> List[str]:
    """提取J列成员（队长+队员，去重）"""
    members = []
    seen = set()
    
    for team in ['1队', '2队', '3队', '4队', '5队', '6队']:
        if team in teams_data:
            data = teams_data[team]
            # 队长
            if data['captain']:
                if data['captain'] not in seen:
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

# ============ 分配函数 ============

def assign_members(j_members: List[str], stats: Dict[str, float], 
                   threshold: float = 900, seed: int = 42) -> Tuple[Dict, Dict, List]:
    """
    分配成员到B列和D列
    
    规则：
    - B列6个位置：排序1-2填B1-B2，排序5-8填B3-B6
    - D列24个位置：排序3-4填D1-D9，排序9-30填D2-D8、D10-D16、D17-D24
    - D2-D8和D10-D16用随机分配
    - D17-D24：步兵生命值低于threshold的随机分配
    
    返回：(B列分配, D列分配, 排序后的成员列表)
    """
    random.seed(seed)
    
    # 获取每个成员的步兵生命值
    members_with_hp = []
    for m in j_members:
        mapped = map_name(m, stats)
        hp = stats.get(mapped, 0) if mapped else 0
        members_with_hp.append({
            'original': m,
            'mapped': mapped if mapped else m,
            'hp': hp,
        })
    
    # 按步兵生命值降序排序
    sorted_members = sorted(members_with_hp, key=lambda x: x['hp'], reverse=True)
    
    b_assign = {}  # B1-B6
    d_assign = {}  # D1-D24
    
    # B1 = 排序1, B2 = 排序2
    if len(sorted_members) >= 1:
        b_assign['B1'] = sorted_members[0]
    if len(sorted_members) >= 2:
        b_assign['B2'] = sorted_members[1]
    
    # D1 = 排序3, D9 = 排序4
    if len(sorted_members) >= 3:
        d_assign['D1'] = sorted_members[2]
    if len(sorted_members) >= 4:
        d_assign['D9'] = sorted_members[3]
    
    # B3-B6 = 排序5-8
    for i, pos in enumerate(['B3', 'B4', 'B5', 'B6']):
        if i + 4 < len(sorted_members):
            b_assign[pos] = sorted_members[i + 4]
    
    # 剩余成员随机分配
    remaining = [m for i, m in enumerate(sorted_members) if i >= 8]
    random.shuffle(remaining)
    
    # D2-D8 = 随机分配
    for i, pos in enumerate(['D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8']):
        if i < len(remaining):
            d_assign[pos] = remaining[i]
    
    # D10-D16 = 随机分配
    rem_idx = 7
    for i, pos in enumerate(['D10', 'D11', 'D12', 'D13', 'D14', 'D15', 'D16']):
        if rem_idx < len(remaining):
            d_assign[pos] = remaining[rem_idx]
            rem_idx += 1
    
    # D17-D24 = 阈值以下随机
    low_hp = [m for m in remaining[rem_idx:] if m['hp'] < threshold]
    random.shuffle(low_hp)
    for i, pos in enumerate(['D17', 'D18', 'D19', 'D20', 'D21', 'D22', 'D23', 'D24']):
        if i < len(low_hp):
            d_assign[pos] = low_hp[i]
    
    return b_assign, d_assign, sorted_members

# ============ 主程序 ============

def main(section: str = "团一", threshold: float = 900):
    # 读取文件
    with open(INPUT_GUANDU, 'r', encoding='utf-8') as f:
        guandu_content = f.read()
    with open(INPUT_STATS, 'r', encoding='utf-8') as f:
        stats_content = f.read()
    
    # 解析
    stats = parse_stats_table(stats_content)
    teams_data, bench_members = parse_guandu_table(guandu_content, section)
    j_members = extract_j_members(teams_data)
    
    print(f"=== {section} 自动分配结果 ===\n")
    print(f"J列成员数: {len(j_members)}")
    print(f"候补数: {len(bench_members)}")
    
    # 分配
    b_assign, d_assign, sorted_members = assign_members(j_members, stats, threshold, RANDOM_SEED)
    
    # 输出排序
    print(f"\n{'='*50}")
    print("按步兵生命值排序（从高到低）")
    print('='*50)
    for i, m in enumerate(sorted_members):
        hp_status = "" if m['hp'] > 0 else " ⚠️未找到"
        print(f"  {i+1:2d}. {m['original']:<15} HP={m['hp']:>10.2f}{hp_status}")
    
    # 输出B列
    print(f"\n{'='*50}")
    print("B列分配（队长）")
    print('='*50)
    for pos in ['B1', 'B2', 'B3', 'B4', 'B5', 'B6']:
        if pos in b_assign:
            m = b_assign[pos]
            print(f"  {pos}: {m['original']} (HP={m['hp']:.2f})")
        else:
            print(f"  {pos}: (空)")
    
    # 输出D列
    print(f"\n{'='*50}")
    print("D列分配（队员）")
    print('='*50)
    
    print("\nD1-D8:")
    for pos in ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8']:
        if pos in d_assign:
            m = d_assign[pos]
            print(f"  {pos}: {m['original']} (HP={m['hp']:.2f})")
        else:
            print(f"  {pos}: (空)")
    
    print("\nD9-D16:")
    for pos in ['D9', 'D10', 'D11', 'D12', 'D13', 'D14', 'D15', 'D16']:
        if pos in d_assign:
            m = d_assign[pos]
            print(f"  {pos}: {m['original']} (HP={m['hp']:.2f})")
        else:
            print(f"  {pos}: (空)")
    
    print("\nD17-D24:")
    for pos in ['D17', 'D18', 'D19', 'D20', 'D21', 'D22', 'D23', 'D24']:
        if pos in d_assign:
            m = d_assign[pos]
            print(f"  {pos}: {m['original']} (HP={m['hp']:.2f})")
        else:
            print(f"  {pos}: (空)")
    
    # 输出候补
    print(f"\n{'='*50}")
    print("候补名单")
    print('='*50)
    for i, m in enumerate(bench_members):
        print(f"  {i+1}. {m}")
    
    # 返回结果供后续使用
    return {
        'b_assign': b_assign,
        'd_assign': d_assign,
        'sorted_members': sorted_members,
        'j_members': j_members,
        'bench_members': bench_members,
        'teams_data': teams_data,
    }

if __name__ == '__main__':
    main()
