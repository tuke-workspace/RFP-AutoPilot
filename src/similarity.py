# -*- coding: utf-8 -*-
"""
相似度匹配模块
使用 SequenceMatcher 算法进行文本相似度计算
"""

import re
from difflib import SequenceMatcher


# 相似度阈值配置
# 总体相似度阈值：提高到0.60，避免不相关内容的误匹配
SIMILARITY_THRESHOLD = 0.60

# V3.2 修改：移除首行权重逻辑，采用全文匹配
# 原逻辑：FIRST_LINE_WEIGHT = 0.8, FULL_TEXT_WEIGHT = 0.2
# 现逻辑：全文匹配，不区分标题和正文

def preprocess_text(text):
    """
    预处理文本：去除标点符号、空白字符、数字等干扰项

    参数:
        text: 原始文本
    返回:
        处理后的文本
    """
    if not text:
        return ""
    # 移除常见标点、空白、数字和特殊符号
    # 保留关键的中文字符和英文单词，去除纯数字和符号
    # 注意：根据用户需求，可能需要保留数字以区分不同参数，但为了模糊匹配，通常去除数字
    text = re.sub(r'[，。、；：""''（）【】\[\]\s≤≥<>＜＞·•\-—_/\\|]+', '', text)
    # 统一大小写
    return text.lower()


def calculate_similarity(text1, text2):
    """
    计算两段文本的相似度

    使用 Python 内置的 SequenceMatcher，基于最长公共子序列算法
    对技术文档的相似文本检测效果好，区分度高

    参数:
        text1: 文本1
        text2: 文本2
    返回:
        0.0 - 1.0 的相似度值
    """
    if not text1 or not text2:
        return 0.0

    # 预处理
    t1 = preprocess_text(text1)
    t2 = preprocess_text(text2)

    if not t1 or not t2:
        return 0.0

    # 使用 SequenceMatcher 计算相似度
    return SequenceMatcher(None, t1, t2).ratio()


def find_best_match(query_text, template_clauses, threshold=None):
    """
    在模板条款列表中找到与查询文本最匹配的条款

    匹配算法改进（V3.2）：
    1. 移除标题权重逻辑，改为全文匹配，确保语义完整性
    2. 遍历所有模板条款，计算全文相似度
    3. 选择相似度最高且超过阈值的条款作为最佳匹配

    参数:
        query_text: 查询文本（招标文档中的条款）
        template_clauses: 模板条款列表，支持两种格式：
            - 旧格式: [(条款内容, 应答内容), ...]
            - 新格式: [{'clause_text': 条款内容, 'response_text': 应答内容, ...}, ...]
        threshold: 相似度阈值，默认使用 SIMILARITY_THRESHOLD
    返回:
        (matched, response_data, similarity, index)
        - matched: 是否匹配成功
        - response_data: 匹配的应答数据（新格式返回整个字典，旧格式返回应答文本，未匹配时为 None）
        - similarity: 最高相似度
        - index: 匹配的条款索引（未匹配时为 -1）
    """
    if threshold is None:
        threshold = SIMILARITY_THRESHOLD

    if not query_text or not template_clauses:
        return False, None, 0.0, -1

    # 预处理查询文本（全文）
    query_processed = preprocess_text(query_text)
    
    if not query_processed:
        return False, None, 0.0, -1

    best_similarity = 0.0
    best_response = None
    best_index = -1

    # 检测数据格式
    is_new_format = template_clauses and isinstance(template_clauses[0], dict)

    for i, item in enumerate(template_clauses):
        # 根据格式获取条款内容
        if is_new_format:
            clause_content = item.get('clause_text', '')
        else:
            clause_content = item[0] if isinstance(item, (tuple, list)) else ''

        # 预处理模板条款（全文）
        template_processed = preprocess_text(clause_content)
        
        if not template_processed:
            continue

        # 计算全文相似度
        similarity = SequenceMatcher(None, query_processed, template_processed).ratio()

        # 更新最佳匹配（如果当前相似度更高）
        if similarity > best_similarity:
            best_similarity = similarity
            # 新格式返回整个字典，旧格式返回应答文本
            best_response = item if is_new_format else item[1]
            best_index = i

    # 判断是否达到阈值
    if best_similarity >= threshold:
        return True, best_response, best_similarity, best_index
    else:
        return False, None, best_similarity, -1


def batch_preprocess(template_clauses):
    """
    批量预处理模板条款（用于优化性能）

    参数:
        template_clauses: 模板条款列表，格式为 [(条款内容, 应答内容), ...]
    返回:
        预处理后的列表 [(预处理后的条款, 应答内容, 原始条款), ...]
    """
    result = []
    for clause_content, response_content in template_clauses:
        processed = preprocess_text(clause_content)
        if processed:
            result.append((processed, response_content, clause_content))
    return result


def find_best_match_optimized(query_text, preprocessed_templates, threshold=None):
    """
    优化版本：在预处理后的模板中查找最佳匹配

    参数:
        query_text: 查询文本
        preprocessed_templates: 预处理后的模板列表（由 batch_preprocess 生成）
        threshold: 相似度阈值
    返回:
        (matched, response, similarity, index)
    """
    if threshold is None:
        threshold = SIMILARITY_THRESHOLD

    if not query_text or not preprocessed_templates:
        return False, None, 0.0, -1

    query_processed = preprocess_text(query_text)
    if not query_processed:
        return False, None, 0.0, -1

    best_similarity = 0.0
    best_response = None
    best_index = -1

    for i, (template_processed, response_content, _) in enumerate(preprocessed_templates):
        similarity = SequenceMatcher(None, query_processed, template_processed).ratio()

        if similarity > best_similarity:
            best_similarity = similarity
            best_response = response_content
            best_index = i

    if best_similarity >= threshold:
        return True, best_response, best_similarity, best_index
    else:
        return False, None, best_similarity, -1

