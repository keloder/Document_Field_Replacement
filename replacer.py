"""
智能替换引擎 - 核心替换逻辑
支持正向替换（原文本 → 替换文本）和反向还原（替换文本 → 原文本）
使用占位符机制避免短规则误替换长规则的问题

核心算法原理：
1. 占位符机制：先将所有匹配项替换为唯一标识符，再替换为目标文本
2. 长度优先：按原文本长度降序处理，避免短规则误替换长规则
3. 唯一标识：使用 UUID 生成唯一占位符，确保无冲突

技术特性：
- 支持正向替换和反向还原
- 保持替换顺序的稳定性
- 返回实际执行的替换操作列表
- 处理空文本和空规则的安全机制

作者：keloder
版本：v0.1
"""

import re
import uuid
from typing import List, Tuple


class SmartReplacer:
    def __init__(self):
        """
        初始化替换器
        
        属性：
        - rules: 存储替换规则的列表，格式为 [(原文本, 替换文本), ...]
        """
        self.rules = []

    def add_rule(self, original: str, replacement: str):
        """
        添加替换规则
        
        参数：
        - original: 原文本（需要被替换的文本）
        - replacement: 替换文本（替换后的文本）
        
        说明：
        - 如果原文本为空，则忽略该规则
        - 规则按添加顺序存储，但替换时会按长度重新排序
        """
        if original:
            self.rules.append((original, replacement))

    def clear_rules(self):
        """清空所有替换规则"""
        self.rules = []

    def replace(self, text: str) -> Tuple[str, List[Tuple[str, str]]]:
        """
        正向替换：将原文本替换为目标文本
        
        算法步骤：
        1. 验证输入：检查规则和文本是否有效
        2. 规则排序：按原文本长度降序排列（避免短规则误替换长规则）
        3. 占位符替换：将匹配的原文本替换为唯一占位符
        4. 目标替换：将占位符替换为对应的替换文本
        5. 记录结果：返回实际执行的替换操作列表
        
        参数：
        - text: 需要处理的文本
        
        返回值：
        - (处理后的文本, 实际执行的替换操作列表)
        """
        # 安全检查：如果无规则或空文本，直接返回
        if not self.rules or not text:
            return text, []

        # 按原文本长度降序排序，确保长规则优先处理
        sorted_rules = sorted(self.rules, key=lambda x: len(x[0]), reverse=True)

        # 创建占位符映射表：占位符 -> (原文本, 替换文本)
        placeholder_map = {}
        
        # 第一步：将所有匹配的原文本替换为唯一占位符
        for i, (original, replacement) in enumerate(sorted_rules):
            # 生成唯一占位符，格式：\x00R{12位UUID}\x00
            ph = f'\x00R{uuid.uuid4().hex[:12]}\x00'
            placeholder_map[ph] = (original, replacement)
            # 执行占位符替换
            text = text.replace(original, ph)

        # 第二步：将占位符替换为对应的替换文本
        replacements_made = []
        for ph, (original, replacement) in placeholder_map.items():
            # 检查占位符是否存在于文本中
            if ph in text:
                # 执行最终替换
                text = text.replace(ph, replacement)
                # 记录实际执行的替换操作
                replacements_made.append((original, replacement))

        return text, replacements_made

    def reverse_replace(self, text: str) -> Tuple[str, List[Tuple[str, str]]]:
        """
        反向还原：将替换文本还原为原文本
        
        算法步骤：
        1. 验证输入：检查规则和文本是否有效
        2. 规则排序：按替换文本长度降序排列
        3. 占位符替换：将匹配的替换文本替换为唯一占位符
        4. 目标替换：将占位符替换为对应的原文本
        5. 记录结果：返回实际执行的还原操作列表
        
        参数：
        - text: 需要还原的文本
        
        返回值：
        - (还原后的文本, 实际执行的还原操作列表)
        """
        # 安全检查：如果无规则或空文本，直接返回
        if not self.rules or not text:
            return text, []

        # 按替换文本长度降序排序，确保长规则优先处理
        sorted_rules = sorted(self.rules, key=lambda x: len(x[1]), reverse=True)

        # 创建占位符映射表：占位符 -> (替换文本, 原文本)
        placeholder_map = {}
        
        # 第一步：将所有匹配的替换文本替换为唯一占位符
        for i, (original, replacement) in enumerate(sorted_rules):
            # 跳过空替换文本的规则
            if not replacement:
                continue
            # 生成唯一占位符
            ph = f'\x00R{uuid.uuid4().hex[:12]}\x00'
            placeholder_map[ph] = (replacement, original)
            # 执行占位符替换
            text = text.replace(replacement, ph)

        # 第二步：将占位符替换为对应的原文本
        replacements_made = []
        for ph, (replacement, original) in placeholder_map.items():
            # 检查占位符是否存在于文本中
            if ph in text:
                # 执行最终还原
                text = text.replace(ph, original)
                # 记录实际执行的还原操作
                replacements_made.append((replacement, original))

        return text, replacements_made
