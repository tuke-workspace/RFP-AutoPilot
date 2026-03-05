import re

def extract_numbered_paragraphs(doc):
    """提取编号段落 - 使用完整的技术参数排除逻辑"""
    
    # 技术单位列表
    tech_units = [
        # 基本电气单位
        r'[vV]',                  # V (伏特)
        r'[aA]',                  # A (安培)
        r'[wW]',                  # W (瓦特)
        r'[hH][zZ]',              # Hz (赫兹)
        r'[Ω]',                   # Ω (欧姆)
        r'[fF]',                  # F (法拉)
        r'[hH]',                  # H (亨利)
        
        # 基本物理单位
        r'[mM]',                  # m (米)
        r'[kK][gG]',              # kg (千克)
        r'[sS]',                  # s (秒)
        r'°[cC]',                 # °C (摄氏度)
        r'[pP][aA]',              # Pa (帕斯卡)
        r'[bB][aA][rR]',          # bar (巴)
        r'%',                     # % (百分比)
        
        # 带前缀的单位
        r'[kK][vV]',              # kV (千伏)
        r'[mM][vV]',              # mV (毫伏)
        r'[μµ][vV]',              # μV (微伏)
        r'[gG][vV]',              # GV (吉伏)
        r'[tT][vV]',              # TV (太伏)
        r'[kK][aA]',              # kA (千安)
        r'[mM][aA]',              # mA (毫安)
        r'[μµ][aA]',              # μA (微安)
        r'[kK][wW]',              # kW (千瓦)
        r'[mM][wW]',              # MW (兆瓦)
        r'[gG][wW]',              # GW (吉瓦)
        r'[kK][hH][zZ]',          # kHz (千赫)
        r'[mM][hH][zZ]',          # MHz (兆赫)
        r'[gG][hH][zZ]',          # GHz (吉赫)
        r'[kK][mM]',              # km (千米)
        r'[cC][mM]',              # cm (厘米)
        r'[mM][mM]',              # mm (毫米)
        r'[μµ][mM]',              # μm (微米)
        
        # 复合单位
        r'[kK]?[vV][aA]',         # VA, kVA (视在功率)
        r'[kK]?[vV][aA][rR]',     # VAR, kVAR (无功功率)
        r'[kK]?[wW][hH]',         # Wh, kWh (瓦时)
        r'[mM]\/[sS]',            # m/s (米每秒)
    ]
    
    # 组合所有技术单位模式
    tech_pattern = '|'.join(tech_units)
    
    clause_pattern = re.compile(rf'''
        ^\s*                  # 行首可选空白
        [*★※＊]?                 # 可选 * 或 ★ 或 ※ 或 ＊
        \s*
        (
        第[一二三四五六七八九十]+[章节]         |  # 第N章 或 第N节
        [一二三四五六七八九十]+、               |  # 中文"一、二、…" 
        [\(\（][一二三四五六七八九十]+[\)\）]、?  |  # （一）或（五）、
        [\(\（]\d+[\)\）]、?                    |  # (1) 或 (2)、
        \d+[）\)]                               |  # 1) 或 1）
        \d+[、，]\s*                                |  # 1、 或 1、 （新增）
        \d+(?:[.．]\d+)*[.．]?(?!\s*(?:{tech_pattern}))(?=[\s\u4e00-\u9fffa-zA-Z*★※＊（(【\[])  # 数字序号，支持半角点和全角句号，支持后接括号
        )
    ''', re.VERBOSE | re.IGNORECASE)
    
    num_paragraphs = []
    for idx, para in enumerate(doc.paragraphs):
        text = (para.text or "").strip()
        prefix, lvl0, is_paren, kind, match_end = get_prefix_and_level(text, clause_pattern)
        if not text or not prefix:
            continue
        
        # 后面如果加了这些单位，则不作为单独条款回应
        # quantity_units = ['台', '个', '套', '件', '只', '根', '条', '张', '块', '片','年','月','日','小时','分','分钟','秒']
        # if prefix and any(text[match_end:].strip().startswith(unit) for unit in quantity_units):
        #     continue
        num_paragraphs.append((idx, prefix, lvl0, is_paren, kind, para))
    
    return num_paragraphs

def get_prefix_and_level(text, clause_pattern):
    """获取前缀和层级信息 - 按照V1.0逻辑"""
    m = clause_pattern.match(text or "")
    if not m:
        return None, None, False, None, 0
    
    prefix = m.group(1)
    is_paren_enclosed = bool(re.match(r'^[\(\（]\d+[\)\）]$', prefix))
    is_paren_half = bool(re.match(r'^\d+[）\)]$', prefix))
    is_paren = is_paren_enclosed or is_paren_half
    
    # 分类 - 完全按照V1.0逻辑
    if re.match(r'^第[一二三四五六七八九十]+[章节]$', prefix):
        kind = 'chapter'
    elif re.match(r'^[一二三四五六七八九十]+、$', prefix):
        kind = 'chinese-section'
    elif re.match(r'^[\(\（][一二三四五六七八九十]+[\)\）]、?$', prefix):
        kind = 'paren-chinese'
    elif is_paren_enclosed:
        kind = "is_paren_enclosed"
    elif is_paren_half:
        kind = "is_paren_half"
    elif re.match(r'^\d+[、，]\s*$', prefix):  # 新增：处理"1、"格式
        kind = "dot-number-1"  # 与"1."同层级
    elif re.match(r'^\d+(?:[.．]\d+)*[.．]?$', prefix):
        depth = len(prefix.rstrip('.．').split('.')) if '.' in prefix else len(prefix.rstrip('.．').split('．'))
        kind = f"dot-number-{depth}"
    else:
        kind = None
    
    # 初步层级 - 完全按照V1.0逻辑
    if kind == 'chapter':
        level0 = (0,)
    elif kind == 'chinese-section':
        level0 = (1,)
    elif kind == 'paren-chinese':
        level0 = (2,)
    elif kind and kind.startswith('dot-number-'):
        depth = int(kind.rsplit('-',1)[1])
        level0 = (depth+2,)
    else:
        level0 = None
    
    return prefix, level0, is_paren, kind, m.end()
    


def analyze_hierarchy(num_paragraphs):
    """分析层级关系"""
    kind_order = {
        'chapter': 1,
        'chinese-section': 2,
        'paren-chinese': 3,
        'dot-number-1': 4,
        'dot-number-2': 5,
        'dot-number-3': 6,
        'dot-number-4': 7,
        'dot-number-5': 8,
        'is_paren_enclosed': 9,
        'is_paren_half': 10,
    }
    
    layered = []
    headings12 = []
    MIN_KEYWORDS = []
    
    # 第一遍：处理非括号编号
    for idx, prefix, lvl0, is_paren, kind, para in num_paragraphs:
        if kind in ('is_paren_enclosed', 'is_paren_half'):
            layered.append((idx, prefix, None, is_paren, kind, para))
            continue
        
        real_level = None
        
        if kind == 'chapter':
            real_level = 1
        elif not layered:
            real_level = 1
        else:
            cur_ord = kind_order.get(kind, 0)

            # 获取前一个有效条款
            prev_item = next((item for item in reversed(layered) if item[2] is not None), None)

            # 特殊处理：dot-number-1 紧跟在 dot-number-2+ 之后
            # 需要区分两种情况：
            # 1. "1.1" 后面跟 "1、2、3、" - 这是子条款格式，应该视为子条款
            # 2. "1.2" 后面跟 "2." - 这是新的一级条款，不是子条款
            if kind == 'dot-number-1' and prev_item:
                # 判断当前序号是否是"N."格式（新的一级条款）还是"N、"格式（子条款）
                is_dot_format = re.match(r'^\d+[.．]$', prefix)  # 如 "2."
                is_comma_format = re.match(r'^\d+[、，]\s*$', prefix)  # 如 "1、"
                
                if prev_item[4] in ('dot-number-2', 'dot-number-3', 'dot-number-4', 'dot-number-5'):
                    # 只有 "N、" 格式才视为子条款，"N." 格式是新的一级条款
                    if is_comma_format:
                        real_level = prev_item[2] + 1
                    # "N." 格式不做特殊处理，使用后面的通用逻辑
                elif prev_item[4] == 'dot-number-1' and prev_item[2] and prev_item[2] >= 4:
                    # 前一个是子条款级别的 dot-number-1
                    if is_comma_format:
                        # "N、" 格式保持相同层级
                        real_level = prev_item[2]
                    # "N." 格式不做特殊处理

            # 如果 real_level 还未设置，使用原有逻辑
            if real_level is None:
                prev = next(
                    (item for item in reversed(layered)
                        if kind_order.get(item[4], 0) < cur_ord),
                    None
                )
                if prev is None:
                    real_level = 1
                else:
                    _, _, prev_lvl, _, prev_kind, _ = prev
                    prev_ord = kind_order.get(prev_kind, 0)
                    if cur_ord > prev_ord:
                        real_level = prev_lvl + 1
                    else:
                        real_level = prev_lvl
        
        if real_level in (1, 2):
            para_text = para.text or ""
            text = para_text.strip()[len(prefix):].lstrip().splitlines()[0] if para_text else ""
            if len(text) <= 20 and any(kw in text for kw in MIN_KEYWORDS):
                headings12.append((idx, real_level, text))
        
        layered.append((idx, prefix, real_level, is_paren, kind, para))
    
    # 第二遍：处理括号编号
    for i, (idx, prefix, real_level, is_paren, kind, para) in enumerate(layered):
        if kind not in ('is_paren_enclosed', 'is_paren_half'):
            continue
        
        prev_h = next((h for h in reversed(headings12) if h[0] < idx), None)
        if not prev_h:
            layered[i] = (idx, prefix, None, is_paren, kind, para)
            continue
        
        prev_item = next(
            (item for item in reversed(layered[:i]) if item[2] is not None),
            None
        )
        if not prev_item:
            layered[i] = (idx, prefix, None, is_paren, kind, para)
            continue
        
        _, _, prev_lvl, _, prev_kind, _ = prev_item
        cur_ord = kind_order.get(kind, 0)
        prev_ord = kind_order.get(prev_kind, 0)
        
        if cur_ord > prev_ord:
            new_lvl = prev_lvl + 1
        elif cur_ord == prev_ord:
            new_lvl = prev_lvl
        else:
            same = next(
                (item for item in reversed(layered[:i]) if item[4] == kind and item[2] is not None),
                None
            )
            new_lvl = same[2] if same else prev_lvl
        
        layered[i] = (idx, prefix, new_lvl, is_paren, kind, para)
    
    return layered, headings12

def find_minimal_clauses(layered):
    """
    找出最小层级条款（需要插入回应的条款）
    
    规则说明：
    - 序号大小规则：第一章/第一节 > 一、 > （一）> 1. > 1.1 > 1.1.1 > 1.1.1.1...
    - 比大小规则：假定当前序号为B，下一个序号为C，如果 B ≤ C，则B为相对最小序号
    - 特殊情况：全文最后一个序号一定是相对最小序号
    - 全文第一个序号只和下一个序号相比
    
    注：代码中 lvl 值越大表示层级越深（序号越小），所以 B ≤ C 等价于 lvl >= next_lvl
    """
    filtered_layered = [
        item
        for item in layered
        if item[2] is not None
    ]

    minimal_indices = []
    n = len(filtered_layered)
    
    for i, (_, prefix, lvl, is_paren, kind, para) in enumerate(filtered_layered):
        if lvl is None:
            continue

        # 特殊情况：全文最后一个序号一定是相对最小序号
        if i == n - 1:
            minimal_indices.append(i)
            break

        _, _, next_lvl, _, next_kind, _ = filtered_layered[i+1]

        # 核心规则：B ≤ C 则B是最小条款
        # B ≤ C 等价于 lvl >= next_lvl（因为lvl越大表示层级越深，序号越小）
        if lvl >= next_lvl:
            minimal_indices.append(i)

    minimal_clauses = [filtered_layered[i] for i in minimal_indices]
    return minimal_clauses 