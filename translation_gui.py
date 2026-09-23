# -*- coding: utf-8 -*-
"""
自动化翻译处理工具 - GUI版本
功能描述:
    1. 支持文件夹和文件两种输入模式
    2. 支持超过100种语言的翻译
    3. 提供翻译确认机制和进度显示
    4. 自动生成各语言版本的XML文件
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
import xml.etree.ElementTree as ET
import base64
import copy
from dataclasses import dataclass
from datetime import datetime, timezone
from email.utils import formatdate
import hmac
import json
import os
import queue
import random
import requests
from urllib.parse import quote, urlparse
import time
import threading
import hashlib
import uuid
from openpyxl import Workbook, load_workbook
import argparse
import tempfile
import logging
import re
import subprocess
import shutil
import ctypes
import platform
import zipfile
import zlib
from PIL import Image, ImageTk

# 语言映射字典 - 包含语言代码、名称和中文说明
# 格式: {'语言缩写': {'code': 'Google翻译代码', 'name': '语言名称 (中文语言:中文国家)'}}
LANG_MAP = {
    'CAT': {'code': 'ca', 'name': 'Catalan: Spain (加泰罗尼亚语:西班牙)'},
    'CHT': {'code': 'zh-TW', 'name': 'Chinese: Traditional (Taiwan) (中文繁体:中国台湾)'},
    'CHS': {'code': 'zh-CN', 'name': 'Chinese: Simplified (PRC) (中文简体:中国)'},
    'ZHH': {'code': 'zh-HK', 'name': 'Chinese: Hong Kong S.A.R. (中文:中国香港)'},
    'ZHI': {'code': 'zh-SG', 'name': 'Chinese: Singapore (中文:新加坡)'},
    'ZHM': {'code': 'zh-MO', 'name': 'Chinese: Macau SAR (中文:中国澳门)'},
    'ARA': {'code': 'ar-SA', 'name': 'Arabic: Saudi Arabia (阿拉伯语:沙特阿拉伯)'},
    'ARI': {'code': 'ar-IQ', 'name': 'Arabic: Iraq (阿拉伯语:伊拉克)'},
    'ARE': {'code': 'ar-EG', 'name': 'Arabic: Egypt (阿拉伯语:埃及)'},
    'ARL': {'code': 'ar-LY', 'name': 'Arabic: Libya (阿拉伯语:利比亚)'},
    'ARG': {'code': 'ar-DZ', 'name': 'Arabic: Algeria (阿拉伯语:阿尔及利亚)'},
    'ARM': {'code': 'ar-MA', 'name': 'Arabic: Morocco (阿拉伯语:摩洛哥)'},
    'ART': {'code': 'ar-TN', 'name': 'Arabic: Tunisia (阿拉伯语:突尼斯)'},
    'ARO': {'code': 'ar-OM', 'name': 'Arabic: Oman (阿拉伯语:阿曼)'},
    'ARY': {'code': 'ar-YE', 'name': 'Arabic: Yemen (阿拉伯语:也门)'},
    'ARS': {'code': 'ar-SY', 'name': 'Arabic: Syria (阿拉伯语:叙利亚)'},
    'ARJ': {'code': 'ar-JO', 'name': 'Arabic: Jordan (阿拉伯语:约旦)'},
    'ARB': {'code': 'ar-LB', 'name': 'Arabic: Lebanon (阿拉伯语:黎巴嫩)'},
    'ARK': {'code': 'ar-KW', 'name': 'Arabic: Kuwait (阿拉伯语:科威特)'},
    'ARU': {'code': 'ar-AE', 'name': 'Arabic: U.A.E. (阿拉伯语:阿联酋)'},
    'ARH': {'code': 'ar-BH', 'name': 'Arabic: Bahrain (阿拉伯语:巴林)'},
    'ARQ': {'code': 'ar-QA', 'name': 'Arabic: Qatar (阿拉伯语:卡塔尔)'},
    'BGR': {'code': 'bg', 'name': 'Bulgarian: Bulgaria (保加利亚语:保加利亚)'},
    'CSY': {'code': 'cs', 'name': 'Czech: Czech Republic (捷克语:捷克)'},
    'DAN': {'code': 'da', 'name': 'Danish: Denmark (丹麦语:丹麦)'},
    'GER': {'code': 'de-DE', 'name': 'German: Germany (德语:德国)'},
    'DES': {'code': 'de-CH', 'name': 'German: Switzerland (德语:瑞士)'},
    'DEA': {'code': 'de-AT', 'name': 'German: Austria (德语:奥地利)'},
    'DEL': {'code': 'de-LU', 'name': 'German: Luxembourg (德语:卢森堡)'},
    'DEC': {'code': 'de-LI', 'name': 'German: Liechtenstein (德语:列支敦士登)'},
    'ELL': {'code': 'el', 'name': 'Greek: Greece (希腊语:希腊)'},
    'USA': {'code': 'en-US', 'name': 'English: United States (英语:美国)'},
    'ENG': {'code': 'en-GB', 'name': 'English: United Kingdom (英语:英国)'},
    'ENA': {'code': 'en-AU', 'name': 'English: Australia (英语:澳大利亚)'},
    'ENC': {'code': 'en-CA', 'name': 'English: Canada (英语:加拿大)'},
    'ENZ': {'code': 'en-NZ', 'name': 'English: New Zealand (英语:新西兰)'},
    'ENI': {'code': 'en-IE', 'name': 'English: Ireland (英语:爱尔兰)'},
    'ENS': {'code': 'en-ZA', 'name': 'English: South Africa (英语:南非)'},
    'ESP': {'code': 'es-ES', 'name': 'Spanish: Spain (西班牙语:西班牙)'},
    'ESM': {'code': 'es-MX', 'name': 'Spanish: Mexico (西班牙语:墨西哥)'},
    'ESG': {'code': 'es-GT', 'name': 'Spanish: Guatemala (西班牙语:危地马拉)'},
    'ESC': {'code': 'es-CR', 'name': 'Spanish: Costa Rica (西班牙语:哥斯达黎加)'},
    'ESA': {'code': 'es-PA', 'name': 'Spanish: Panama (西班牙语:巴拿马)'},
    'ESD': {'code': 'es-DO', 'name': 'Spanish: Dominican Republic (西班牙语:多米尼加)'},
    'ESV': {'code': 'es-VE', 'name': 'Spanish: Venezuela (西班牙语:委内瑞拉)'},
    'ESO': {'code': 'es-CO', 'name': 'Spanish: Colombia (西班牙语:哥伦比亚)'},
    'ESR': {'code': 'es-PE', 'name': 'Spanish: Peru (西班牙语:秘鲁)'},
    'ESS': {'code': 'es-AR', 'name': 'Spanish: Argentina (西班牙语:阿根廷)'},
    'ESF': {'code': 'es-EC', 'name': 'Spanish: Ecuador (西班牙语:厄瓜多尔)'},
    'ESL': {'code': 'es-CL', 'name': 'Spanish: Chile (西班牙语:智利)'},
    'ESY': {'code': 'es-UY', 'name': 'Spanish: Uruguay (西班牙语:乌拉圭)'},
    'ESZ': {'code': 'es-PY', 'name': 'Spanish: Paraguay (西班牙语:巴拉圭)'},
    'ESB': {'code': 'es-BO', 'name': 'Spanish: Bolivia (西班牙语:玻利维亚)'},
    'ESE': {'code': 'es-SV', 'name': 'Spanish: El Salvador (西班牙语:萨尔瓦多)'},
    'ESH': {'code': 'es-HN', 'name': 'Spanish: Honduras (西班牙语:洪都拉斯)'},
    'ESI': {'code': 'es-NI', 'name': 'Spanish: Nicaragua (西班牙语:尼加拉瓜)'},
    'ESU': {'code': 'es-PR', 'name': 'Spanish: Puerto Rico (西班牙语:波多黎各)'},
    'FIN': {'code': 'fi', 'name': 'Finnish: Finland (芬兰语:芬兰)'},
    'FRA': {'code': 'fr-FR', 'name': 'French: France (法语:法国)'},
    'FRB': {'code': 'fr-BE', 'name': 'French: Belgium (法语:比利时)'},
    'FRC': {'code': 'fr-CA', 'name': 'French: Canada (法语:加拿大)'},
    'FRS': {'code': 'fr-CH', 'name': 'French: Switzerland (法语:瑞士)'},
    'FRL': {'code': 'fr-LU', 'name': 'French: Luxembourg (法语:卢森堡)'},
    'FRM': {'code': 'fr-MC', 'name': 'French: Monaco (法语:摩纳哥)'},
    'HEB': {'code': 'he', 'name': 'Hebrew: Israel (希伯来语:以色列)'},
    'HUN': {'code': 'hu', 'name': 'Hungarian: Hungary (匈牙利语:匈牙利)'},
    'ISL': {'code': 'is', 'name': 'Icelandic: Iceland (冰岛语:冰岛)'},
    'ITA': {'code': 'it-IT', 'name': 'Italian: Italy (意大利语:意大利)'},
    'ITS': {'code': 'it-CH', 'name': 'Italian: Switzerland (意大利语:瑞士)'},
    'JPN': {'code': 'ja', 'name': 'Japanese: Japan (日语:日本)'},
    'KOR': {'code': 'ko', 'name': 'Korean: Korea (韩语:韩国)'},
    'NLD': {'code': 'nl-NL', 'name': 'Dutch: Netherlands (荷兰语:荷兰)'},
    'NLB': {'code': 'nl-BE', 'name': 'Dutch: Belgium (荷兰语:比利时)'},
    'NOR': {'code': 'nb', 'name': 'Norwegian: Norway (Bokmål) (挪威语:挪威)'},
    'NON': {'code': 'nn', 'name': 'Norwegian: Norway (Nynorsk) (挪威语:挪威)'},
    'PLK': {'code': 'pl', 'name': 'Polish: Poland (波兰语:波兰)'},
    'PTB': {'code': 'pt-BR', 'name': 'Portuguese: Brazil (葡萄牙语:巴西)'},
    'PTG': {'code': 'pt-PT', 'name': 'Portuguese: Portugal (葡萄牙语:葡萄牙)'},
    'ROM': {'code': 'ro', 'name': 'Romanian: Romania (罗马尼亚语:罗马尼亚)'},
    'RUS': {'code': 'ru', 'name': 'Russian: Russia (俄语:俄罗斯)'},
    'HRV': {'code': 'hr', 'name': 'Croatian: Croatia (克罗地亚语:克罗地亚)'},
    'SRL': {'code': 'sr-Latn', 'name': 'Serbian: Serbia (Latin) (塞尔维亚语:塞尔维亚)'},
    'SRB': {'code': 'sr-Cyrl', 'name': 'Serbian: Serbia (Cyrillic) (塞尔维亚语:塞尔维亚)'},
    'SKY': {'code': 'sk', 'name': 'Slovak: Slovakia (斯洛伐克语:斯洛伐克)'},
    'SQI': {'code': 'sq', 'name': 'Albanian: Albania (阿尔巴尼亚语:阿尔巴尼亚)'},
    'SVE': {'code': 'sv-SE', 'name': 'Swedish: Sweden (瑞典语:瑞典)'},
    'SVF': {'code': 'sv-FI', 'name': 'Swedish: Finland (瑞典语:芬兰)'},
    'THA': {'code': 'th', 'name': 'Thai: Thailand (泰语:泰国)'},
    'TRK': {'code': 'tr', 'name': 'Turkish: Turkey (土耳其语:土耳其)'},
    'URP': {'code': 'ur-PK', 'name': 'Urdu: Pakistan (乌尔都语:巴基斯坦)'},
    'IND': {'code': 'id', 'name': 'Indonesian: Indonesia (印尼语:印尼)'},
    'UKR': {'code': 'uk', 'name': 'Ukrainian: Ukraine (乌克兰语:乌克兰)'},
    'BEL': {'code': 'be', 'name': 'Belarusian: Belarus (白俄罗斯语:白俄罗斯)'},
    'SLV': {'code': 'sl', 'name': 'Slovene: Slovenia (斯洛文尼亚语:斯洛文尼亚)'},
    'ETI': {'code': 'et', 'name': 'Estonian: Estonia (爱沙尼亚语:爱沙尼亚)'},
    'LVI': {'code': 'lv', 'name': 'Latvian: Latvia (拉脱维亚语:拉脱维亚)'},
    'LTH': {'code': 'lt', 'name': 'Lithuanian: Lithuania (立陶宛语:立陶宛)'},
    'LTC': {'code': 'lt', 'name': 'Classic Lithuanian: Lithuania (立陶宛语:立陶宛)'},
    'FAR': {'code': 'fa', 'name': 'Farsi: Iran (波斯语:伊朗)'},
    'VIT': {'code': 'vi', 'name': 'Vietnamese: Vietnam (越南语:越南)'},
    'HYE': {'code': 'hy', 'name': 'Armenian: Armenia (亚美尼亚语:亚美尼亚)'},
    'AZE': {'code': 'az-Latn', 'name': 'Azeri: Azerbaijan (Latin) (阿塞拜疆语:阿塞拜疆)'},
    'EUQ': {'code': 'eu', 'name': 'Basque: Spain (巴斯克语:西班牙)'},
    'MKI': {'code': 'mk', 'name': 'FYRO Macedonian (马其顿语:马其顿)'},
    'AFK': {'code': 'af', 'name': 'Afrikaans: South Africa (南非语:南非)'},
    'KAT': {'code': 'ka', 'name': 'Georgian: Georgia (格鲁吉亚语:格鲁吉亚)'},
    'FOS': {'code': 'fo', 'name': 'Faeroese: Faeroe Islands (法罗语:法罗群岛)'},
    'HIN': {'code': 'hi', 'name': 'Hindi: India (印地语:印度)'},
    'MSL': {'code': 'ms-MY', 'name': 'Malay: Malaysia (马来语:马来西亚)'},
    'MSB': {'code': 'ms-BN', 'name': 'Malay: Brunei Darussalam (马来语:文莱)'},
    'KAZ': {'code': 'kk', 'name': 'Kazak: Kazakhstan (哈萨克语:哈萨克斯坦)'},
    'SWK': {'code': 'sw', 'name': 'Swahili: Kenya (斯瓦希里语:肯尼亚)'},
    'UZB': {'code': 'uz-Latn', 'name': 'Uzbek: Uzbekistan (Latin) (乌兹别克语:乌兹别克斯坦)'},
    'TAT': {'code': 'tt', 'name': 'Tatar: Tatarstan (鞑靼语:鞑靼斯坦)'},
    'BEN': {'code': 'bn', 'name': 'Bengali: India (孟加拉语:印度)'},
    'PAN': {'code': 'pa', 'name': 'Punjabi: India (旁遮普语:印度)'},
    'GUJ': {'code': 'gu', 'name': 'Gujarati: India (古吉拉特语:印度)'},
    'ORI': {'code': 'or', 'name': 'Oriya: India (奥里亚语:印度)'},
    'TAM': {'code': 'ta', 'name': 'Tamil: India (泰米尔语:印度)'},
    'TEL': {'code': 'te', 'name': 'Telugu: India (泰卢固语:印度)'},
    'KAN': {'code': 'kn', 'name': 'Kannada: India (卡纳达语:印度)'},
    'MAL': {'code': 'ml', 'name': 'Malayalam: India (马拉雅拉姆语:印度)'},
    'ASM': {'code': 'as', 'name': 'Assamese: India (阿萨姆语:印度)'},
    'MAR': {'code': 'mr', 'name': 'Marathi: India (马拉地语:印度)'},
    'SAN': {'code': 'sa', 'name': 'Sanskrit: India (梵语:印度)'},
    'KOK': {'code': 'kok', 'name': 'Konkani: India (孔卡尼语:印度)'}
}

# RES输出时追加的语言标识: 语言代码 → 本机语言名+中文国家名
LANG_NATIVE_NAME = {
    'CHS': '中文(简体) 中国', 'CHT': '中文(繁體) 台湾', 'ZHH': '中文(香港) 香港',
    'ZHI': '中文(台湾) 台湾', 'ZHM': '中文(澳门) 澳门',
    'ENG': 'English 英国', 'USA': 'English 美国', 'ENA': 'English 美国', 'ENC': 'English 英国',
    'ENZ': 'English 新西兰', 'ENI': 'English 印度', 'ENS': 'English 新加坡',
    'GER': 'Deutsch 德国', 'DES': 'Deutsch 瑞士', 'DEA': 'Deutsch 奥地利', 'DEL': 'Deutsch 列支敦士登', 'DEC': 'Deutsch 卢森堡',
    'JPN': '日本語 日本', 'KOR': '한국어 韩国', 'FRA': 'Français 法国', 'FRB': 'Français 比利时',
    'FRC': 'Français 加拿大', 'FRS': 'Français 瑞士', 'FRL': 'Français 卢森堡', 'FRM': 'Français 摩纳哥',
    'ITA': 'Italiano 意大利', 'ITS': 'Italiano 瑞士',
    'ESP': 'Español 西班牙', 'ESM': 'Español 墨西哥', 'ESG': 'Español 危地马拉', 'ESC': 'Español 哥斯达黎加',
    'ESA': 'Español 巴拿马', 'ESD': 'Español 多米尼加共和国', 'ESV': 'Español 委内瑞拉', 'ESO': 'Español 哥伦比亚',
    'ESR': 'Español 秘鲁', 'ESS': 'Español 阿根廷', 'ESF': 'Español 厄瓜多尔', 'ESL': 'Español 智利',
    'ESY': 'Español 乌拉圭', 'ESB': 'Español 玻利维亚', 'ESE': 'Español 巴拉圭', 'ESH': 'Esp萨尔瓦多',
    'ESN': 'Español 尼加拉瓜', 'ESU': 'Español 波多黎各', 'ESP2': 'Español 美国',
    'RUS': 'Русский 俄罗斯', 'PTG': 'Português 葡萄牙', 'PTB': 'Português 巴西',
    'TRK': 'Türkçe 土耳其', 'PLK': 'Polski 波兰', 'VIT': 'Tiếng Việt 越南',
    'CAT': 'Català 西班牙', 'THA': 'ไทย 泰国', 'DAN': 'Dansk 丹麦', 'NON': 'Norsk 挪威',
    'SVE': 'Svenska 瑞典', 'FIN': 'Suomi 芬兰', 'NLD': 'Nederlands 荷兰', 'NLB': 'Nederlands 比利时',
    'CSY': 'Čeština 捷克', 'SKY': 'Slovenčina 斯洛伐克', 'HUN': 'Magyar 匈牙利', 'ELL': 'Ελληνικά 希腊',
    'ARA': 'العربية 沙特阿拉伯', 'ARL': 'العربية 利比亚', 'ARG': 'العربية 阿尔及利亚', 'ARM': 'العربية 摩洛哥',
    'ART': 'الع العربية 伊拉克', 'ARO': 'العربية 阿曼', 'ARY': 'العربية 埃及', 'ARS': 'العربية 叙利亚',
    'ARJ': 'العربية 约旦', 'ARB': 'العربية 黎巴嫩', 'ARK': 'العربية 科威特', 'ARU': 'العربية 阿联酋',
    'ARH': 'العربية 巴林', 'ARQ': 'العربية 卡塔尔',
    'BGR': 'Български 保加利亚', 'UKR': 'Українська 乌克兰', 'BEL': 'Беларуская 白俄罗斯',
    'KAZ': 'Қазақша 哈萨克斯坦', 'SRB': 'Српски 塞尔维亚', 'HRV': 'Hrvatski 克罗地亚', 'SLV': 'Slovenščina 斯洛文尼亚',
    'EST': 'Eesti 爱沙尼亚', 'LVI': 'Latviešu 拉脱维亚', 'LTH': 'Lietuvių 立陶宛', 'LTC': 'Lietuvių 立陶宛',
    'ROM': 'Română 罗马尼亚', 'ROM2': 'Română 摩尔多瓦', 'IND': 'Bahasa Indonesia 印尼',
    'HEB': 'עברית 以色列', 'FAR': 'فارسی 伊朗', 'URD': 'اردو 巴基斯坦',
    'HYE': 'Հայերեն 亚美尼亚', 'KAT': 'ქართული 格鲁吉亚', 'AZE': 'Azərbaycan 阿塞拜疆',
    'HIN': 'हिन्दी 印度', 'BEN': 'বাংলা 孟加拉', 'PAN': 'ਪੰਜਾਬੀ 印度', 'GUJ': 'ગુજરાતી 印度',
    'ORI': 'ଓଡ଼ିଆ 印度', 'TAM': 'தமிழ் 印度', 'TEL': 'తెలుగు 印度', 'KAN': 'ಕನ್ನಡ 印度',
    'MAL': 'മലയാളം 印度', 'ASM': 'অসমীয়া 印度', 'MAR': 'मराठी 印度', 'SAN': 'संस्कृतम् 印度',
    'KOK': 'कोंकणी 印度', 'MSL': 'Bahasa Melayu 马来西亚', 'MSB': 'Bahasa Melayu 文莱',
    'SWK': 'Kiswahili 坦桑尼亚', 'UZB': 'O\'zbek 乌兹别克斯坦', 'TAT': 'Татарча 俄罗斯',
    'MKI': 'Македонски 北马其顿', 'AFK': 'Afrikaans 南非', 'FOS': 'Føroyskt 法罗群岛',
}

# 默认显示的语言列表（11种常用语言）
DEFAULT_LANGS = ['CAT', 'CHT', 'ESP', 'GER', 'ITA', 'KOR', 'PLK', 'PTG', 'RUS', 'TRK', 'VIT']

# 品牌名称集合 - 这些名称不翻译
BRAND_NAMES = {'禾川', '高创', '松下', '汇川', '迪维迅', '台达', '雷赛', '信捷', 'H系列'}

# 配置文件名
CONFIG_FILE = 'translation_config.json'

# 右键 DiskC 按钮时可设置的待实现工作流。选择仅保存本地偏好，不改变当前 DiskC 的左键行为。
WORKFLOW_SELECTION_OPTIONS = (
    ('marking_cam', '打标Cam工作流'),
    ('mb_backup', 'MB备份工作流'),
    ('simulator', '模拟器工作流'),
)
WORKFLOW_SELECTION_LABELS = dict(WORKFLOW_SELECTION_OPTIONS)

MB_BACKUP_FILE_SPECS = (
    (
        'alarm_macro',
        'AlarmMacro/AlarmMacro_CHS.xml',
        'AlarmMacro/AlarmMacro_{lang}.xml',
    ),
    (
        'alarm_plc',
        'Ladder/AlarmPLC_CHS.xml',
        'Ladder/AlarmPLC_{lang}.xml',
    ),
    (
        'param_ext',
        'ParameterExt/ParamExt_CHS.xml',
        'ParameterExt/ParamExt_{lang}.xml',
    ),
    (
        'param_ext_rbit',
        'ParameterExt/ParamExt_RBit_CHS.xml',
        'ParameterExt/ParamExt_RBit_{lang}.xml',
    ),
)
MB_BACKUP_OCRES_SOURCE_PREFIX = 'OCRes/CHS/String/'
MB_BACKUP_OCRES_TARGET_TEMPLATE = 'OCRes/{lang}/String/{relative_path}'
MB_BACKUP_INPUT_OPTIONS = (
    ('zip', 'MB备份 ZIP 压缩包'),
    ('folder', '已解压的 MB备份文件夹'),
)
MB_BACKUP_INPUT_LABELS = dict(MB_BACKUP_INPUT_OPTIONS)

VOLCENGINE_SERVICE = 'translate'
VOLCENGINE_DEFAULT_REGION = 'cn-beijing'
VOLCENGINE_API_VERSION = '2020-06-01'

# API 按官方免费额度/付费服务优先的顺序展示。凭据始终只保存在本机配置文件中，不会写入日志。
API_PROVIDERS = {
    'deepl_free': {
        'label': 'DeepL API Free（官方免费额度）',
        'description': '使用 DeepL Free Auth Key；免费和 Pro 使用不同服务地址。',
        'fields': [
            {'key': 'auth_key', 'label': 'Auth Key', 'default': '', 'required': True, 'secret': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://api-free.deepl.com/v2', 'required': True},
        ],
    },
    'baidu': {
        'label': '百度翻译（官方免费/付费）',
        'description': '使用百度翻译开放平台 App ID 和密钥。',
        'fields': [
            {'key': 'app_id', 'label': 'App ID', 'default': '', 'required': True},
            {'key': 'secret_key', 'label': 'Secret Key', 'default': '', 'required': True, 'secret': True},
        ],
    },
    'youdao': {
        'label': '有道智云（官方试用/付费）',
        'description': '使用有道智云文本翻译 v3 的 App Key 和 App Secret。',
        'fields': [
            {'key': 'app_key', 'label': 'App Key', 'default': '', 'required': True},
            {'key': 'app_secret', 'label': 'App Secret', 'default': '', 'required': True, 'secret': True},
        ],
    },
    'niutrans': {
        'label': '小牛翻译（官方免费额度/付费）',
        'description': '使用小牛翻译 API Key，可按账户文档调整服务地址。',
        'fields': [
            {'key': 'api_key', 'label': 'API Key', 'default': '', 'required': True, 'secret': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://api.niutrans.com/NiuTransServer/translation', 'required': True},
        ],
    },
    'tencent': {
        'label': '腾讯云 TMT（官方免费额度/付费）',
        'description': '使用腾讯云 SecretId、SecretKey，并按 TC3-HMAC-SHA256 签名。',
        'fields': [
            {'key': 'secret_id', 'label': 'SecretId', 'default': '', 'required': True},
            {'key': 'secret_key', 'label': 'SecretKey', 'default': '', 'required': True, 'secret': True},
            {'key': 'region', 'label': '地域', 'default': 'ap-beijing', 'required': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://tmt.tencentcloudapi.com', 'required': True},
        ],
    },
    'volcengine': {
        'label': '火山引擎机器翻译（官方免费额度/付费）',
        'description': '使用火山引擎 AccessKey 和 Secret AccessKey；机器翻译固定使用 cn-beijing/translate。',
        'fields': [
            {'key': 'access_key', 'label': 'AccessKey ID', 'default': '', 'required': True},
             {'key': 'secret_key', 'label': 'Secret AccessKey', 'default': '', 'required': True, 'secret': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://translate.volcengineapi.com', 'required': True},
        ],
    },
    'aliyun': {
        'label': '阿里云机器翻译（官方免费额度/付费）',
        'description': '使用阿里云 AccessKey，默认调用通用版机器翻译接口。',
        'fields': [
            {'key': 'access_key_id', 'label': 'AccessKey ID', 'default': '', 'required': True},
            {'key': 'access_key_secret', 'label': 'AccessKey Secret', 'default': '', 'required': True, 'secret': True},
            {'key': 'scene', 'label': '场景', 'default': 'general', 'required': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://mt.cn-hangzhou.aliyuncs.com/api/translate/web/general', 'required': True},
            {'key': 'batch_endpoint', 'label': '批量服务地址', 'default': 'https://mt.cn-hangzhou.aliyuncs.com/', 'required': True},
            {'key': 'region_id', 'label': '地域 ID', 'default': 'cn-hangzhou', 'required': True},
            {'key': 'batch_api_type', 'label': '批量 API 类型', 'default': 'translate_standard', 'required': True},
        ],
    },
    'deepl_pro': {
        'label': 'DeepL API Pro（官方付费）',
        'description': '使用 DeepL Pro Auth Key；请勿与 Free Key 混用。',
        'fields': [
            {'key': 'auth_key', 'label': 'Auth Key', 'default': '', 'required': True, 'secret': True},
            {'key': 'endpoint', 'label': '服务地址', 'default': 'https://api.deepl.com/v2', 'required': True},
        ],
    },
    'google': {
        'label': 'Google GTX（仅测试，不建议生产）',
        'description': '非官方公共接口，可能被限流或中断；不适合正式批量翻译。',
        'fields': [],
    },
}

# The public GTX endpoint is particularly sensitive to concurrent requests.
API_PROVIDER_WORKERS = {
    'google': 2,
    'deepl_free': 4,
    'baidu': 4,
    'youdao': 4,
    'niutrans': 4,
    'tencent': 4,
    'volcengine': 4,
    'aliyun': 4,
    'deepl_pro': 4,
}

# Tencent TMT reports a 5 requests/second account limit. Keep one request of headroom
# because the limit is enforced across all worker threads and uses a rolling window.
API_PROVIDER_REQUESTS_PER_SECOND = {
    'tencent': 4.0,
}

TMT_BATCH_MAX_TEXTS = 200
TMT_BATCH_MAX_TEXT_LENGTH = 2000

# Only providers with an ordered, documented array response use the shared batch workflow.
# DeepL keeps an 8 KiB margin below its documented 128 KiB request-body limit.
API_BATCH_LIMITS = {
    'tencent': {
        'max_texts': TMT_BATCH_MAX_TEXTS,
        'max_text_length': TMT_BATCH_MAX_TEXT_LENGTH,
    },
    'deepl_free': {
        'max_texts': 50,
        'max_total_bytes': 120 * 1024,
    },
    'deepl_pro': {
        'max_texts': 50,
        'max_total_bytes': 120 * 1024,
    },
    'volcengine': {
        'max_texts': 16,
        'max_total_characters': 5000,
    },
    'aliyun': {
        'max_texts': 50,
        'max_text_length': 1000,
        'max_total_characters': 8000,
    },
}

REMOVED_API_PROVIDERS = {
    'libretranslate': 'google',
}


def default_provider_configs():
    """Create a fresh local configuration dictionary for every provider."""
    return {
        provider: {field['key']: field['default'] for field in spec['fields']}
        for provider, spec in API_PROVIDERS.items()
    }


@dataclass(frozen=True)
class TranslationResult:
    """The result of one provider request, including an actionable failure state."""
    success: bool
    text: str = ''
    error: str = ''


@dataclass(frozen=True)
class MbBackupSource:
    """One staged MB backup localization resource and its ZIP output mapping."""
    category: str
    archive_entry: str
    staged_path: str
    target_template: str
    relative_path: str = ''

    def target_entry(self, language):
        return self.target_template.format(
            lang=language,
            relative_path=self.relative_path
        )


class TranslationProviderError(Exception):
    """An expected provider or transport error."""

    def __init__(self, message, retryable=False):
        super().__init__(message)
        self.retryable = retryable


class TranslationApp:
    """
    自动化翻译处理工具主类
    提供可视化用户界面，支持XML文件翻译和多语言输出
    """
    
    def __init__(self, root):
        """
        初始化方法
        :param root: Tkinter主窗口对象
        """
        self.root = root
        self.root.title("自动化翻译处理工具")
        self.root.geometry("1000x800")
        
        # 初始化实例变量
        self.source_files = []           # 源文件列表
        self.source_folder = ""         # 源文件夹路径（CLI兼容）
        self.source_folders = []        # 已添加的文件夹路径列表（统一UI）
        self.selected_langs = []        # 选中的语言列表
        self.translation_table = {}     # 翻译表字典
        self.output_dir = os.path.dirname(os.path.abspath(__file__))  # 输出目录
        self.table_dir = os.path.dirname(os.path.abspath(__file__))   # 翻译表保存目录
        self.default_langs = DEFAULT_LANGS.copy()  # 默认语言列表
        self.show_all_langs = False     # 是否显示所有语言
        self.default_maximized = False  # 启动时是否默认窗口最大化
        self.translation_complete = False  # 翻译是否完成
        self.translation_failures = []
        self.headless = False  # 是否为无界面模式（命令行）
        self.logger = None
        
        # DiskC工作流模式
        self.diskc_mode = False       # 是否为DiskC工作流模式
        self.diskc_root = ""          # DiskC根目录路径
        self.diskc_res_source = ""    # CHS.res解压后的prepared路径
        self.diskc_xml_map = {}       # prepared_path -> 相对于String目录的子路径
        self.diskc_plugin_source = "" # Plugin/Config/CHS.xml的prepared路径

        # 模拟器工作流模式
        self.simulator_mode = False
        self.simulator_root = ""
        self.simulator_xml_map = {}   # prepared_path -> 相对于CHS/String目录的子路径

        # MB备份工作流模式
        self.mb_backup_mode = False
        self.mb_backup_input_type = ""
        self.mb_backup_root_path = ""
        self.mb_backup_archive_path = ""
        self.mb_backup_output_base_path = ""
        self.mb_backup_work_dir = ""
        self.mb_backup_sources = {}
        self.mb_backup_output_path = ""
        self.mb_backup_source_type = "zip"
        self.pack_res_default = False
        self.selected_workflow = ""   # 右键工作流菜单的本地保存选择

        # API配置变量。保留百度旧字段，兼容已有 translation_config.json。
        self.api_type = 'google'
        self.provider_configs = default_provider_configs()
        self.baidu_app_id = ''
        self.baidu_secret_key = ''
        
        # 翻译缓存 - 避免相同文本重复翻译
        self.translation_cache = {}
        
        # Each worker owns its Session; requests.Session is not safe to share across workers.
        self._http_local = threading.local()
        self._log_queue = queue.Queue()
        self._ui_queue = queue.Queue()
        self._provider_rate_lock = threading.Lock()
        self._provider_next_request_at = {}
        self._main_thread_id = threading.get_ident()
        
        # 加载配置、创建界面、加载翻译表
        self.load_config()
        self.create_widgets()
        self._flush_pending_logs()
        self._flush_pending_ui_tasks()
        self.load_translation_table()

    def _sync_legacy_baidu_credentials(self):
        """Keep legacy Baidu configuration keys working after the provider migration."""
        baidu_config = self.provider_configs['baidu']
        self.baidu_app_id = baidu_config.get('app_id', '')
        self.baidu_secret_key = baidu_config.get('secret_key', '')
    
    def load_config(self):
        """
        加载配置文件
        从translation_config.json读取用户自定义的默认语言、窗口模式和API配置。
        旧版百度顶级配置会自动迁移到 provider_configs。
        """
        if not hasattr(self, 'default_maximized'):
            self.default_maximized = False
        if not hasattr(self, 'selected_workflow'):
            self.selected_workflow = ''
        if not hasattr(self, 'mb_backup_source_type'):
            self.mb_backup_source_type = 'zip'
        if not hasattr(self, 'pack_res_default'):
            self.pack_res_default = False
        if os.path.exists(CONFIG_FILE):
            config_migrated = False
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'default_langs' in config:
                        self.default_langs = config['default_langs']
                    saved_maximized = config.get('default_maximized')
                    legacy_fullscreen = config.get('default_fullscreen')
                    if isinstance(saved_maximized, bool):
                        self.default_maximized = saved_maximized
                    elif isinstance(legacy_fullscreen, bool):
                        # Migrate the previous full-screen setting to window maximization.
                        self.default_maximized = legacy_fullscreen
                    if 'default_fullscreen' in config:
                        config_migrated = True
                    saved_workflow = config.get('selected_workflow', '')
                    if isinstance(saved_workflow, str) and saved_workflow in WORKFLOW_SELECTION_LABELS:
                        self.selected_workflow = saved_workflow
                    elif saved_workflow not in (None, ''):
                        self.selected_workflow = ''
                        config_migrated = True
                        self.log(f"忽略不支持的工作流设置: {saved_workflow!r}")
                    saved_mb_source_type = config.get('mb_backup_source_type', 'zip')
                    if saved_mb_source_type in MB_BACKUP_INPUT_LABELS:
                        self.mb_backup_source_type = saved_mb_source_type
                    else:
                        self.mb_backup_source_type = 'zip'
                        if saved_mb_source_type not in (None, ''):
                            config_migrated = True
                            self.log(f"忽略不支持的 MB备份输入类型设置: {saved_mb_source_type!r}")
                    saved_pack_res_default = config.get('pack_res_default')
                    if isinstance(saved_pack_res_default, bool):
                        self.pack_res_default = saved_pack_res_default
                    elif self.selected_workflow == 'marking_cam':
                        # Preserve the previous marking-CAM default for old configs.
                        self.pack_res_default = True
                    elif saved_pack_res_default not in (None, ''):
                        config_migrated = True
                        self.pack_res_default = False
                        self.log(f"忽略不支持的打包 .res 设置: {saved_pack_res_default!r}")
                    saved_api_type = config.get('api_type')
                    if saved_api_type in API_PROVIDERS:
                        self.api_type = saved_api_type
                    elif saved_api_type in REMOVED_API_PROVIDERS:
                        self.api_type = REMOVED_API_PROVIDERS[saved_api_type]
                        config_migrated = True
                        self.log(
                            "LibreTranslate 接口已移除，已切换为 Google GTX（仅测试）。"
                            "请在 API 配置中选择其他服务。"
                        )
                    saved_providers = config.get('provider_configs', {})
                    if isinstance(saved_providers, dict):
                        for provider, values in saved_providers.items():
                            if provider in REMOVED_API_PROVIDERS:
                                config_migrated = True
                                continue
                            if provider not in self.provider_configs or not isinstance(values, dict):
                                continue
                            if provider == 'volcengine' and 'region' in values:
                                config_migrated = True
                            for field in self.provider_configs[provider]:
                                value = values.get(field)
                                if isinstance(value, str):
                                    self.provider_configs[provider][field] = value

                    # v1.2 and earlier stored only the Baidu credentials at the top level.
                    legacy_app_id = config.get('baidu_app_id', '')
                    legacy_secret_key = config.get('baidu_secret_key', '')
                    if legacy_app_id and not self.provider_configs['baidu']['app_id']:
                        self.provider_configs['baidu']['app_id'] = legacy_app_id
                    if legacy_secret_key and not self.provider_configs['baidu']['secret_key']:
                        self.provider_configs['baidu']['secret_key'] = legacy_secret_key
                    self._sync_legacy_baidu_credentials()
                if config_migrated:
                    self.save_config()
            except (OSError, json.JSONDecodeError) as e:
                self.log(f"加载配置失败: {e}")
                self._ensure_logger()
                if self.logger:
                    self.logger.exception(f"加载配置失败: {e}")

    def _ensure_logger(self, log_file=None):
        """
        初始化文件日志（可选）
        """
        if self.logger:
            return
        try:
            self.logger = logging.getLogger('translation_app')
            self.logger.setLevel(logging.DEBUG)
            fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            # 控制台处理器
            ch = logging.StreamHandler()
            ch.setFormatter(fmt)
            self.logger.addHandler(ch)
            # 可选文件处理器
            if log_file:
                fh = logging.FileHandler(log_file, encoding='utf-8')
                fh.setFormatter(fmt)
                self.logger.addHandler(fh)
        except Exception:
            self.logger = None
    
    def save_config(self):
        """
        保存配置文件
        将当前默认语言设置和各翻译服务的本地配置保存到translation_config.json
        """
        self._sync_legacy_baidu_credentials()
        selected_workflow = getattr(self, 'selected_workflow', '')
        if selected_workflow not in WORKFLOW_SELECTION_LABELS:
            selected_workflow = ''
            self.selected_workflow = ''
        mb_backup_source_type = getattr(self, 'mb_backup_source_type', 'zip')
        if mb_backup_source_type not in MB_BACKUP_INPUT_LABELS:
            mb_backup_source_type = 'zip'
            self.mb_backup_source_type = mb_backup_source_type
        pack_res_default = getattr(self, 'pack_res_default', False)
        if not isinstance(pack_res_default, bool):
            pack_res_default = False
            self.pack_res_default = pack_res_default
        config = {
            'default_langs': self.default_langs,
            'default_maximized': bool(getattr(self, 'default_maximized', False)),
            'selected_workflow': selected_workflow,
            'mb_backup_source_type': mb_backup_source_type,
            'pack_res_default': pack_res_default,
            'api_type': self.api_type,
            'provider_configs': self.provider_configs,
            # Retain these keys so configuration files stay compatible with older releases.
            'baidu_app_id': self.baidu_app_id,
            'baidu_secret_key': self.baidu_secret_key
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    
    def create_widgets(self):
        self.root.title("Translate Pro")
        self.root.geometry("1280x850")
        self.root.configure(bg='#F5F6FA')
        
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TFrame', background='#F5F6FA')
        style.configure('Card.TFrame', background='white')
        style.configure('Sidebar.TFrame', background='#2D2B4E')
        style.configure('Stats.TFrame', background='white')
        style.configure('Header.TLabel', background='#2D2B4E', foreground='white', font=('Segoe UI', 11, 'bold'))
        style.configure('Nav.TLabel', background='#2D2B4E', foreground='#A5A0C0', font=('Segoe UI', 10))
        style.configure('NavActive.TLabel', background='#2D2B4E', foreground='white', font=('Segoe UI', 10, 'bold'))
        style.configure('Title.TLabel', background='#F5F6FA', foreground='#2D2B4E', font=('Segoe UI', 13, 'bold'))
        style.configure('Subtitle.TLabel', background='#F5F6FA', foreground='#6B6B8D', font=('Segoe UI', 9))
        style.configure('StatValue.TLabel', background='white', foreground='#2D2B4E', font=('Segoe UI', 24, 'bold'))
        style.configure('StatLabel.TLabel', background='white', foreground='#8E8EA0', font=('Segoe UI', 9))
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('TButton', font=('Segoe UI', 9), padding=6)
        style.configure('Tab.TButton', font=('Segoe UI', 9), padding=(12, 6), background='#E8E8F0', foreground='#6E6E73')
        style.map('Tab.TButton', background=[('active', '#D0D0E0')])
        style.configure('TabActive.TButton', font=('Segoe UI', 9, 'bold'), padding=(12, 6), background='#FFFFFF', foreground='#6C63FF')
        style.map('TabActive.TButton', background=[('active', '#F5F5F7')])
        style.configure('TCheckbutton', background='#F5F6FA', font=('Segoe UI', 9))
        style.configure('TProgressbar', thickness=8, troughcolor='#E8E8F0', background='#6C63FF')
        style.configure('Card.TLabelframe', background='white')
        style.configure('Card.TLabelframe.Label', background='white', foreground='#2D2B4E', font=('Segoe UI', 11, 'bold'))
        
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        sidebar = ttk.Frame(self.root, style='Sidebar.TFrame', width=200)
        sidebar.grid(row=0, column=0, sticky='ns')
        sidebar.grid_propagate(False)
        
        ttk.Label(sidebar, text="Translate", style='Header.TLabel').pack(pady=(30, 8), anchor='w', padx=24)
        ttk.Label(sidebar, text="Pro", style='Header.TLabel', foreground='#6C63FF').pack(anchor='w', padx=24)
        
        ttk.Separator(sidebar, orient='horizontal').pack(fill='x', pady=20, padx=16)
        
        nav_items = [
            ("📁", "文件管理", "files"),
            ("🌐", "语言选择", "langs"),
            ("🖼️", "Logo设置", "logo"),
            ("📊", "处理进度", "progress"),
            ("📝", "处理日志", "logs"),
            ("⚙️", "设置选项", "settings")
        ]
        
        self.nav_buttons = {}
        for icon, text, name in nav_items:
            btn_frame = ttk.Frame(sidebar, style='Sidebar.TFrame')
            btn_frame.pack(fill='x', padx=12, pady=2)
            
            nav_btn_frame = ttk.Frame(btn_frame, style='Sidebar.TFrame')
            nav_btn_frame.pack(fill='x', pady=8, padx=12)
            
            icon_lbl = ttk.Label(nav_btn_frame, text=icon, style='Nav.TLabel', 
                                cursor='hand2', width=2)
            icon_lbl.pack(side='left', padx=(0, 8))
            
            text_lbl = ttk.Label(nav_btn_frame, text=text, style='Nav.TLabel',
                                cursor='hand2')
            text_lbl.pack(side='left')
            
            def bind_events(widget, label):
                widget.bind('<Enter>', lambda e, l=label: l.configure(style='NavActive.TLabel'))
                widget.bind('<Leave>', lambda e, l=label: l.configure(style='Nav.TLabel'))
            
            bind_events(icon_lbl, icon_lbl)
            bind_events(text_lbl, text_lbl)
            
            for widget in (btn_frame, nav_btn_frame, icon_lbl, text_lbl):
                widget.bind(
                    '<Button-1>',
                    lambda e, n=name: self._handle_nav_click(e, n)
                )
            self.nav_buttons[name] = btn_frame
        
        self.current_content = None
        
        ttk.Frame(sidebar, style='Sidebar.TFrame').pack(side='bottom', fill='x', pady=20)
        
        main_content = ttk.Frame(self.root, style='TFrame')
        main_content.grid(row=0, column=1, sticky='nsew', padx=16, pady=16)
        main_content.grid_columnconfigure(0, weight=1)
        main_content.grid_rowconfigure(3, weight=1)
        
        header_frame = ttk.Frame(main_content, style='TFrame')
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 16))
        header_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(header_frame, text="翻译工作台", style='Title.TLabel').grid(row=0, column=0, sticky='w')
        ttk.Label(header_frame, text="支持 XML/RES 文件 · 100+ 语言 · 智能缩写", style='Subtitle.TLabel').grid(row=1, column=0, sticky='w', pady=(2, 0))
        
        action_frame = ttk.Frame(header_frame, style='TFrame')
        action_frame.grid(row=0, column=2, rowspan=2, sticky='e')
        
        self.start_btn = ttk.Button(action_frame, text="▶ 开始翻译", command=self.start_translation, style='Accent.TButton')
        self.start_btn.pack(side='left', padx=4)
        
        self.confirm_btn = ttk.Button(action_frame, text="✓ 确认生成", command=self.confirm_and_generate, state='disabled')
        self.confirm_btn.pack(side='left', padx=4)
        
        ttk.Button(action_frame, text="修改LOGO并创建快捷方式", command=lambda: self.switch_tab('logo')).pack(side='left', padx=4)
        
        content_area = ttk.Frame(main_content, style='TFrame')
        content_area.grid(row=1, column=0, sticky='nsew')
        content_area.grid_columnconfigure(0, weight=3)
        content_area.grid_columnconfigure(1, weight=1)
        content_area.grid_rowconfigure(2, weight=1)
        
        files_card = ttk.LabelFrame(content_area, text=" 📂 输入源文件 ", style='Card.TLabelframe', padding=12)
        files_card.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 12))
        files_card.grid_columnconfigure(1, weight=1)
        
        toolbar = ttk.Frame(files_card, style='Card.TFrame')
        toolbar.grid(row=0, column=0, columnspan=2, sticky='ew', pady=(0, 8))
        
        ttk.Button(toolbar, text="+ 添加文件夹", command=self.add_folder).pack(side='left', padx=2)
        ttk.Button(toolbar, text="+ 添加文件", command=self.add_files).pack(side='left', padx=2)
        self.diskc_workflow_button = ttk.Button(
            toolbar,
            text="DiskC 工作流",
            command=self._start_selected_workflow
        )
        self.diskc_workflow_button.pack(side='left', padx=2)
        self.workflow_choice_var = tk.StringVar(value=self.selected_workflow)
        self.workflow_menu = tk.Menu(self.root, tearoff=False)
        for workflow_id, workflow_label in WORKFLOW_SELECTION_OPTIONS:
            self.workflow_menu.add_radiobutton(
                label=workflow_label,
                value=workflow_id,
                variable=self.workflow_choice_var,
                command=self._save_selected_workflow
            )
        self.diskc_workflow_button.bind('<Button-3>', self._show_workflow_menu)
        ttk.Button(toolbar, text="移除选中", command=self.remove_files).pack(side='left', padx=2)
        ttk.Button(toolbar, text="清空列表", command=self.clear_files).pack(side='left', padx=2)
        
        ttk.Separator(toolbar, orient='vertical').pack(side='left', fill='y', padx=8, pady=2)
        
        self.pack_var = tk.BooleanVar(value=self.pack_res_default)
        self._update_workflow_button_label()
        self._apply_workflow_defaults()
        
        ttk.Button(toolbar, text="翻译表", command=self.view_translation_table).pack(side='left', padx=2)
        ttk.Button(toolbar, text="导出Excel", command=self.export_to_excel).pack(side='left', padx=2)
        ttk.Button(toolbar, text="导入Excel", command=self.import_from_excel).pack(side='left', padx=2)
        ttk.Button(toolbar, text="导出日志", command=self.export_log).pack(side='left', padx=2)
        
        list_container = ttk.Frame(files_card, style='Card.TFrame')
        list_container.grid(row=1, column=0, columnspan=2, sticky='nsew')
        files_card.grid_rowconfigure(1, weight=1)
        
        self.file_listbox = tk.Listbox(list_container, height=4, bg='#FAFBFE', fg='#2D2B4E',
                                       font=('Segoe UI', 9), selectbackground='#6C63FF', selectforeground='white',
                                       borderwidth=0, highlightthickness=1, highlightbackground='#E8E8F0')
        self.file_listbox.pack(side='left', fill='both', expand=True)
        
        file_scrollbar = ttk.Scrollbar(list_container, orient='vertical', command=self.file_listbox.yview)
        file_scrollbar.pack(side='right', fill='y')
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        self.tab_frame = ttk.Frame(content_area, style='Card.TFrame')
        self.tab_frame.grid(row=1, column=0, sticky='nsew', padx=(0, 6))
        self.tab_frame.grid_rowconfigure(1, weight=1)
        self.tab_frame.grid_columnconfigure(0, weight=1)
        
        self.tab_buttons = ttk.Frame(self.tab_frame, style='Card.TFrame')
        self.tab_buttons.grid(row=0, column=0, sticky='ew')
        
        self.lang_tab_btn = ttk.Button(self.tab_buttons, text="🌍 目标语言", 
                                       command=lambda: self.switch_tab('langs'),
                                       style='Tab.TButton')
        self.lang_tab_btn.pack(side='left', padx=2)
        
        self.logo_tab_btn = ttk.Button(self.tab_buttons, text="🖼️ Logo设置", 
                                       command=lambda: self.switch_tab('logo'),
                                       style='Tab.TButton')
        self.logo_tab_btn.pack(side='left', padx=2)
        
        self.lang_card = ttk.LabelFrame(self.tab_frame, text=" 目标语言 ", style='Card.TLabelframe', padding=12)
        self.lang_card.grid(row=1, column=0, sticky='nsew')
        self.lang_card.grid_rowconfigure(1, weight=1)
        
        self.logo_card = ttk.LabelFrame(self.tab_frame, text=" Logo和快捷方式设置 ", style='Card.TLabelframe', padding=12)
        self.logo_card.grid(row=1, column=0, sticky='nsew')
        self.logo_card.grid_rowconfigure(1, weight=1)
        self.logo_card.grid_remove()

        self.settings_card = ttk.LabelFrame(
            self.tab_frame,
            text=" 设置选项 ",
            style='Card.TLabelframe',
            padding=12
        )
        self.settings_card.grid(row=1, column=0, sticky='nsew')
        self.settings_card.grid_remove()
        
        self.current_tab = 'langs'
        
        lang_toolbar = ttk.Frame(self.lang_card, style='Card.TFrame')
        lang_toolbar.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        
        self.lang_show_var = tk.BooleanVar(value=self.show_all_langs)
        ttk.Checkbutton(lang_toolbar, text="显示全部语言", variable=self.lang_show_var,
                       command=self.toggle_lang_display).pack(side='left', padx=4)
        
        self.main_lang_search_var = tk.StringVar()
        search_entry = ttk.Entry(lang_toolbar, textvariable=self.main_lang_search_var, width=14,
                                  font=('Segoe UI', 9))
        search_entry.pack(side='left', padx=(8, 4))
        search_entry.insert(0, '🔍 搜索...')
        search_entry.configure(foreground='#A5A0C0')
        
        def _on_focus_in(e):
            if self.main_lang_search_var.get() == '🔍 搜索...':
                search_entry.delete(0, tk.END)
                search_entry.configure(foreground='#2D2B4E')
        
        def _on_focus_out(e):
            if not self.main_lang_search_var.get().strip():
                search_entry.insert(0, '🔍 搜索...')
                search_entry.configure(foreground='#A5A0C0')
        
        search_entry.bind('<FocusIn>', _on_focus_in)
        search_entry.bind('<FocusOut>', _on_focus_out)
        self.main_lang_search_var.trace('w', lambda *args: self._filter_main_langs())
        
        ttk.Separator(lang_toolbar, orient='vertical').pack(side='left', fill='y', padx=6, pady=2)
        
        api_frame = ttk.Frame(lang_toolbar, style='Card.TFrame')
        api_frame.pack(side='left', padx=(0, 4))
        
        ttk.Label(api_frame, text="API:", font=('Segoe UI', 9)).pack(side='left')
        self.api_type_var = tk.StringVar(value=self._provider_label(self.api_type))
        api_combo = ttk.Combobox(
            api_frame,
            textvariable=self.api_type_var,
            values=[spec['label'] for spec in API_PROVIDERS.values()],
            state='readonly',
            width=30,
            font=('Segoe UI', 9)
        )
        api_combo.pack(side='left', padx=3)
        api_combo.bind('<<ComboboxSelected>>', lambda e: self.on_api_type_changed())
        ttk.Button(api_frame, text="配置", command=self.configure_api).pack(side='left')
        
        ttk.Button(lang_toolbar, text="设置默认", command=self.set_default_langs).pack(side='left', padx=4)
        
        btn_frame = ttk.Frame(lang_toolbar, style='Card.TFrame')
        btn_frame.pack(side='right')
        ttk.Button(btn_frame, text="全选", command=self.select_all_langs).pack(side='left', padx=2)
        ttk.Button(btn_frame, text="清空", command=self.deselect_all_langs).pack(side='left', padx=2)
        
        lang_canvas_frame = ttk.Frame(self.lang_card, style='Card.TFrame')
        lang_canvas_frame.grid(row=1, column=0, sticky='nsew')
        self.lang_card.grid_rowconfigure(1, weight=1)
        
        canvas = tk.Canvas(lang_canvas_frame, bg='white', highlightthickness=0, width=340, height=260)
        scrollbar_y = ttk.Scrollbar(lang_canvas_frame, orient='vertical', command=canvas.yview)
        scrollable_lang = ttk.Frame(canvas, style='Card.TFrame')
        
        scrollable_lang.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable_lang, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar_y.set)
        
        def on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind_all('<MouseWheel>', on_mousewheel, add='+')
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        
        self.lang_frame = scrollable_lang
        
        right_panel = ttk.Frame(content_area, style='TFrame', width=220)
        right_panel.grid(row=1, column=1, sticky='nsew', padx=(12, 0))
        right_panel.grid_propagate(False)
        right_panel.grid_rowconfigure(1, weight=1)
        
        stats_card = ttk.LabelFrame(right_panel, text=" 📊 统计信息 ", style='Card.TLabelframe', padding=16)
        stats_card.grid(row=0, column=0, sticky='ew', pady=(0, 12))
        
        self.stat_file_count = ttk.Label(stats_card, text="--", style='StatValue.TLabel')
        self.stat_file_count.pack(anchor='center')
        ttk.Label(stats_card, text="已添加文件", style='StatLabel.TLabel').pack(anchor='center', pady=(0, 12))
        
        stat_mid = ttk.Frame(stats_card, style='Stats.TFrame')
        stat_mid.pack(fill='x', pady=8)
        
        self.stat_lang_count = ttk.Label(stat_mid, text="--", style='StatValue.TLabel', font=('Segoe UI', 16, 'bold'))
        self.stat_lang_count.pack(side='left')
        ttk.Label(stat_mid, text="\n已选语言", style='StatLabel.TLabel').pack(side='left', padx=(8, 0))
        
        progress_card = ttk.LabelFrame(right_panel, text=" ⏳ 处理进度 ", style='Card.TLabelframe', padding=16)
        progress_card.grid(row=1, column=0, sticky='nsew')
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_card, variable=self.progress_var, maximum=100,
                                           mode='determinate', length=200)
        self.progress_bar.pack(fill='x', pady=(0, 8))
        
        pct_frame = ttk.Frame(progress_card, style='Stats.TFrame')
        pct_frame.pack()
        self.progress_label = ttk.Label(pct_frame, text="0%", font=('Segoe UI', 14, 'bold'),
                                        foreground='#6C63FF')
        self.progress_label.pack(side='left')
        ttk.Label(pct_frame, text="  就绪", style='Subtitle.TLabel').pack(side='left')
        
        log_card = ttk.LabelFrame(main_content, text=" 📋 处理日志 ", style='Card.TLabelframe', padding=12)
        log_card.grid(row=2, column=0, sticky='nsew', pady=(12, 0))
        main_content.grid_rowconfigure(2, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_card, height=10, wrap=tk.WORD,
                                                   font=('Consolas', 9), bg='#FAFBFE', fg='#4A4A68',
                                                   borderwidth=0, highlightthickness=0)
        self.log_text.pack(fill='both', expand=True)
        self.log_text.insert(tk.END, "欢迎使用 Translate Pro！\n")
        self.log_text.insert(tk.END, "请添加输入源文件并选择目标语言，然后点击「开始翻译」。\n")
        
        stats_panel = ttk.Frame(self.root, style='Stats.TFrame', width=220)
        stats_panel.grid(row=0, column=2, sticky='ns', padx=(0, 16), pady=16)
        stats_panel.grid_propagate(False)
        
        outer_stats = ttk.Frame(stats_panel, style='Stats.TFrame', padding=20)
        outer_stats.pack(fill='both', expand=True)
        
        ttk.Label(outer_stats, text="Statistic", font=('Segoe UI', 14, 'bold'),
                  foreground='#2D2B4E', background='white').pack(anchor='w', pady=(0, 20))
        
        stat_items = [
            ("downloads", "本周处理", "0", "次"),
            ("space", "可用空间", "--", ""),
            ("shared", "翻译条目", "0", "条")
        ]
        
        self.stats_refs = {}
        for icon_id, label, default_val, unit in stat_items:
            item_frame = ttk.Frame(outer_stats, style='Stats.TFrame')
            item_frame.pack(fill='x', pady=10)
            
            left_part = ttk.Frame(item_frame, style='Stats.TFrame')
            left_part.pack(side='left')
            
            val_lbl = ttk.Label(left_part, text=default_val, style='StatValue.TLabel',
                                font=('Segoe UI', 18, 'bold'))
            val_lbl.pack(anchor='w')
            ttk.Label(left_part, text=label, style='StatLabel.TLabel').pack(anchor='w')
            
            self.stats_refs[icon_id] = val_lbl
            
            if unit:
                ttk.Label(item_frame, text=f"\n{unit}", style='StatLabel.TLabel').pack(side='left', padx=(8, 0))
        
        ttk.Separator(outer_stats, orient='horizontal').pack(fill='x', pady=16)
        
        tip_frame = ttk.Frame(outer_stats, style='Stats.TFrame')
        tip_frame.pack(fill='x')
        ttk.Label(tip_frame, text="💡 提示", font=('Segoe UI', 10, 'bold'),
                  foreground='#2D2B4E', background='white').pack(anchor='w')
        ttk.Label(tip_frame, text="智能缩写将自动优化\n超长翻译文本的长度",
                  font=('Segoe UI', 9), foreground='#8E8EA0', background='white',
                  justify='left').pack(anchor='w', pady=(4, 0))
        
        self.init_logo_settings()
        self.init_window_settings()
        
        self.update_lang_display()
        self._apply_window_mode(self.default_maximized)
    
    def show_content(self, name):
        if name == 'logo':
            self.switch_tab('logo')
        elif name == 'langs':
            self.switch_tab('langs')
        elif name == 'settings':
            self.switch_tab('settings')

    def _handle_nav_click(self, _event, name):
        self.show_content(name)
        return 'break'
    
    def switch_tab(self, tab_name):
        if self.current_tab == tab_name:
            return
        
        if self.current_tab == 'langs':
            self.lang_card.grid_remove()
        elif self.current_tab == 'logo':
            self.logo_card.grid_remove()
        elif self.current_tab == 'settings':
            self.settings_card.grid_remove()
        
        if tab_name == 'langs':
            self.lang_card.grid(row=1, column=0, sticky='nsew')
            self.lang_tab_btn.configure(style='TabActive.TButton')
            self.logo_tab_btn.configure(style='Tab.TButton')
        elif tab_name == 'logo':
            self.logo_card.grid(row=1, column=0, sticky='nsew')
            self.lang_tab_btn.configure(style='Tab.TButton')
            self.logo_tab_btn.configure(style='TabActive.TButton')
        elif tab_name == 'settings':
            self.settings_card.grid(row=1, column=0, sticky='nsew')
            self.lang_tab_btn.configure(style='Tab.TButton')
            self.logo_tab_btn.configure(style='Tab.TButton')
        
        self.current_tab = tab_name

    def init_window_settings(self):
        settings_main = ttk.Frame(self.settings_card, style='Card.TFrame')
        settings_main.pack(fill='both', expand=True, padx=8, pady=8)

        ttk.Label(
            settings_main,
            text="窗口启动模式",
            font=('Segoe UI', 11, 'bold'),
            foreground='#2D2B4E'
        ).pack(anchor='w')
        ttk.Label(
            settings_main,
            text="选择程序启动时使用窗口化还是窗口最大化。保存后立即应用，并在下次启动时保持。",
            font=('Segoe UI', 9),
            foreground='#6E6E73',
            wraplength=620,
            justify='left'
        ).pack(anchor='w', pady=(4, 14))

        self.window_mode_var = tk.StringVar(
            value='maximized' if self.default_maximized else 'windowed'
        )
        ttk.Radiobutton(
            settings_main,
            text="默认窗口化",
            variable=self.window_mode_var,
            value='windowed'
        ).pack(anchor='w', pady=4)
        ttk.Radiobutton(
            settings_main,
            text="默认窗口最大化",
            variable=self.window_mode_var,
            value='maximized'
        ).pack(anchor='w', pady=4)

        ttk.Separator(settings_main, orient='horizontal').pack(
            fill=tk.X,
            pady=(18, 14)
        )
        ttk.Label(
            settings_main,
            text="MB备份输入类型",
            font=('Segoe UI', 11, 'bold'),
            foreground='#2D2B4E'
        ).pack(anchor='w')
        ttk.Label(
            settings_main,
            text="保存后，点击“MB备份工作流”会直接打开对应的选择器，不再询问输入类型。",
            font=('Segoe UI', 9),
            foreground='#6E6E73',
            wraplength=620,
            justify='left'
        ).pack(anchor='w', pady=(4, 10))

        self.mb_backup_source_type_var = tk.StringVar(
            value=self.mb_backup_source_type
        )
        for source_type, label in MB_BACKUP_INPUT_OPTIONS:
            ttk.Radiobutton(
                settings_main,
                text=label,
                variable=self.mb_backup_source_type_var,
                value=source_type
            ).pack(anchor='w', pady=3)
        ttk.Checkbutton(
            settings_main,
            text="默认打包 .res",
            variable=self.pack_var
        ).pack(anchor='w', pady=(10, 0))

        ttk.Button(
            settings_main,
            text="保存并应用",
            command=self.save_window_settings
        ).pack(anchor='w', pady=(16, 0))

    def _apply_window_mode(self, maximized):
        if self.root.state() == 'withdrawn':
            return True
        try:
            self.root.state('zoomed' if maximized else 'normal')
            if not maximized:
                self.root.geometry("1280x850")
            return True
        except tk.TclError as exc:
            self.log(f"窗口模式设置失败: {exc}")
            return False

    def save_window_settings(self):
        self.default_maximized = self.window_mode_var.get() == 'maximized'
        selected_source_type = self.mb_backup_source_type_var.get()
        if selected_source_type in MB_BACKUP_INPUT_LABELS:
            self.mb_backup_source_type = selected_source_type
        else:
            self.mb_backup_source_type = 'zip'
        self.pack_res_default = bool(self.pack_var.get())
        applied = self._apply_window_mode(self.default_maximized)
        self.save_config()
        if applied:
            messagebox.showinfo("设置", "窗口启动模式已保存并应用。")
        else:
            messagebox.showwarning(
                "设置",
                "窗口启动模式已保存，但当前系统应用失败，请重启程序后重试。"
            )
    
    def init_logo_settings(self):
        self.logo_image_path = None
        
        logo_main = ttk.Frame(self.logo_card, style='Card.TFrame')
        logo_main.pack(fill='both', expand=True, padx=8, pady=8)
        
        title_label = ttk.Label(logo_main, text="轻松转换图片格式并创建桌面快捷方式", 
                               font=('Segoe UI', 10), foreground='#6e6e73')
        title_label.pack(anchor='w', pady=(0, 12))
        
        select_frame = ttk.Frame(logo_main, style='Card.TFrame')
        select_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Label(select_frame, text="1. 选择图片:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 4))
        
        self.logo_file_label = ttk.Label(select_frame, text="未选择文件", font=('Segoe UI', 11), foreground='#6e6e73')
        self.logo_file_label.pack(side='left', padx=(0, 8))
        
        ttk.Button(select_frame, text="选择图片", command=self.select_logo_image).pack(side='left')
        
        settings_frame = ttk.Frame(logo_main, style='Card.TFrame')
        settings_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Label(settings_frame, text="2. 转换设置:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 4))
        
        ttk.Label(settings_frame, text="公司名称:", font=('Segoe UI', 10)).pack(anchor='w', pady=(0, 2))
        self.ico_filename_var = tk.StringVar(value="LOGO")
        ico_entry = ttk.Entry(settings_frame, textvariable=self.ico_filename_var, width=20, font=('Segoe UI', 10))
        ico_entry.pack(anchor='w')
        
        action_frame = ttk.Frame(logo_main, style='Card.TFrame')
        action_frame.pack(fill='x', pady=(0, 12))
        
        ttk.Label(action_frame, text="3. 执行操作:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 4))
        
        button_frame = ttk.Frame(action_frame)
        button_frame.pack(fill='x')
        
        self.logo_action_btn = ttk.Button(button_frame, text="修改LOGO并创建快捷方式", 
                                          command=self.modify_logo_and_create_shortcut,
                                          state='disabled')
        self.logo_action_btn.pack(side='left', padx=(0, 8))
        
        ttk.Button(button_frame, text="仅转换图片", command=self.convert_logo_images).pack(side='left', padx=(0, 8))
        ttk.Button(button_frame, text="仅创建快捷方式", command=self.create_shortcut_only).pack(side='left')
        
        status_frame = ttk.Frame(logo_main, style='Card.TFrame')
        status_frame.pack(fill='x')
        
        self.logo_status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(status_frame, textvariable=self.logo_status_var, 
                                font=('Segoe UI', 11), foreground='#007AFF')
        status_label.pack(anchor='w')
        
        self.logo_progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.logo_progress.pack(fill='x', pady=(8, 0))
    
    def add_folder(self):
        folder = filedialog.askdirectory(title="选择输入文件夹")
        if not folder:
            return
        self.source_folders.append(folder)
        if not self.source_files:
            self.output_dir = os.path.dirname(folder)
            self.log(f"输出目录已设置为: {self.output_dir}")
        count = 0
        for root_dir, dirs, files in os.walk(folder):
            for file in files:
                file_path = os.path.join(root_dir, file)
                prepared = self.prepare_file(file_path)
                if prepared and prepared not in self.source_files:
                    self.source_files.append(prepared)
                    tag = self._file_tag(prepared)
                    self.file_listbox.insert(tk.END, f"{tag} {prepared}")
                    count += 1
        self.log(f"从文件夹添加 {count} 个文件: {folder}")

    def refresh_file_list(self):
        self.source_files = []
        self.file_listbox.delete(0, tk.END)
        all_folders = list(self.source_folders)
        if self.source_folder and self.source_folder not in all_folders:
            all_folders.append(self.source_folder)
        if not all_folders:
            self.log("无已添加文件夹，请先添加文件夹或文件")
            return
        for folder in all_folders:
            if os.path.isdir(folder):
                count = 0
                for root_dir, dirs, files in os.walk(folder):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        prepared = self.prepare_file(file_path)
                        if prepared and prepared not in self.source_files:
                            self.source_files.append(prepared)
                            tag = self._file_tag(prepared)
                            self.file_listbox.insert(tk.END, f"{tag} {prepared}")
                            count += 1
                self.log(f"刷新文件夹: {os.path.basename(folder)} → {count} 个文件")
        self.log(f"总计扫描到 {len(self.source_files)} 个可处理文件")

    def _reset_diskc_state(self):
        self.diskc_mode = False
        self.diskc_root = ""
        self.diskc_res_source = ""
        self.diskc_xml_map = {}
        self.diskc_plugin_source = ""

    def _reset_simulator_state(self):
        self.simulator_mode = False
        self.simulator_root = ""
        self.simulator_xml_map = {}

    def _reset_mb_backup_state(self):
        work_dir = getattr(self, 'mb_backup_work_dir', '')
        if work_dir and os.path.isdir(work_dir):
            try:
                shutil.rmtree(work_dir)
                self.log(f"已删除 MB备份临时目录: {work_dir}")
            except OSError as error:
                self.log(f"删除 MB备份临时目录失败: {work_dir} -> {error}")
        self.mb_backup_mode = False
        self.mb_backup_input_type = ""
        self.mb_backup_root_path = ""
        self.mb_backup_archive_path = ""
        self.mb_backup_output_base_path = ""
        self.mb_backup_work_dir = ""
        self.mb_backup_sources = {}
        self.mb_backup_output_path = ""

    def _show_workflow_menu(self, event):
        try:
            self.workflow_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.workflow_menu.grab_release()
            except tk.TclError:
                pass
        return 'break'

    def _save_selected_workflow(self):
        workflow_id = self.workflow_choice_var.get()
        workflow_label = WORKFLOW_SELECTION_LABELS.get(workflow_id)
        if workflow_label is None:
            self.log(f"忽略不支持的工作流选择: {workflow_id!r}")
            return
        self.selected_workflow = workflow_id
        self._apply_workflow_defaults()
        self._update_workflow_button_label()
        self.save_config()
        if workflow_id == 'marking_cam':
            self.log(f"已保存工作流预设: {workflow_label}（适用于打标/玻切DiskC文件夹翻译）")
        elif workflow_id == 'mb_backup':
            self.log(f"已保存工作流预设: {workflow_label}（支持 ZIP 或已解压文件夹）")
        else:
            self.log(f"已保存工作流预设: {workflow_label}（后端尚未实现）")

    def _workflow_button_label(self):
        return WORKFLOW_SELECTION_LABELS.get(self.selected_workflow, 'DiskC 工作流')

    def _update_workflow_button_label(self):
        if hasattr(self, 'diskc_workflow_button'):
            self.diskc_workflow_button.config(text=self._workflow_button_label())

    def _apply_workflow_defaults(self):
        if self.selected_workflow == 'marking_cam':
            default_pack_res = True
        elif self.selected_workflow in ('mb_backup', 'simulator'):
            default_pack_res = False
        else:
            default_pack_res = getattr(self, 'pack_res_default', False)
        self.pack_res_default = default_pack_res
        if hasattr(self, 'pack_var'):
            self.pack_var.set(default_pack_res)

    def _start_selected_workflow(self):
        if self.selected_workflow == 'marking_cam':
            self.setup_marking_cam_sources()
            return
        if self.selected_workflow == 'mb_backup':
            self.setup_mb_backup_sources()
            return
        if self.selected_workflow == 'simulator':
            self.setup_simulator_sources()
            return
        if not self.selected_workflow:
            self.setup_diskc_sources()
            return

        workflow_label = WORKFLOW_SELECTION_LABELS[self.selected_workflow]
        message = f"{workflow_label}后端尚未实现。"
        self.log(message)
        messagebox.showinfo("工作流", message)

    def setup_marking_cam_sources(self, marking_cam_root=None):
        """暂时复用 DiskC 的目录发现、输出路由和 RES 打包后端。"""
        self.log("打标Cam工作流当前复用 DiskC 工作流后端。")
        self.setup_diskc_sources(marking_cam_root)

    def setup_simulator_sources(self, simulator_root=None):
        if simulator_root is None:
            simulator_root = filedialog.askdirectory(
                title="选择模拟器工作流根目录"
            )
            if not simulator_root:
                return

        simulator_root = os.path.abspath(simulator_root)
        nested_diskc_root = os.path.join(simulator_root, 'DiskC')
        if os.path.isdir(os.path.join(nested_diskc_root, 'OpenCnc Shared')):
            simulator_root = nested_diskc_root

        string_dir = os.path.join(
            simulator_root,
            'OpenCnc Shared',
            'OCRes',
            'CHS',
            'String'
        )
        if not os.path.isdir(string_dir):
            message = f"模拟器工作流未找到目录: {string_dir}"
            self.log(message)
            if not self.headless:
                messagebox.showerror("模拟器工作流", message)
            return

        self._reset_mb_backup_state()
        self._reset_diskc_state()
        self._reset_simulator_state()
        self.simulator_mode = True
        self.simulator_root = simulator_root
        self.source_files = []
        self.source_folders = []
        self.source_folder = simulator_root
        self.output_dir = simulator_root
        self.file_listbox.delete(0, tk.END)

        xml_count = 0
        for root_dir, dirs, files in os.walk(string_dir):
            dirs.sort()
            for file_name in sorted(files):
                if not file_name.lower().endswith('.xml'):
                    continue
                full_path = os.path.join(root_dir, file_name)
                prepared = self.prepare_file(full_path)
                if not prepared:
                    continue
                relative_path = os.path.relpath(prepared, string_dir)
                self.simulator_xml_map[prepared] = relative_path
                self.source_files.append(prepared)
                self.file_listbox.insert(
                    tk.END,
                    f"[SIM] String/{relative_path}"
                )
                xml_count += 1

        self.update_stats()
        message = f"模拟器工作流已就绪: {xml_count} 个 String XML 文件"
        self.log(message)
        if not self.headless:
            messagebox.showinfo("模拟器工作流", message)

    @staticmethod
    def _mb_backup_staging_path(work_dir, archive_entry):
        path_parts = archive_entry.split('/')
        if (
            not path_parts
            or any(
                not part or part in ('.', '..') or '\\' in part or ':' in part
                for part in path_parts
            )
        ):
            raise ValueError(f"MB备份包含不安全的文件路径: {archive_entry}")
        work_dir = os.path.abspath(work_dir)
        staged_path = os.path.abspath(os.path.join(work_dir, *path_parts))
        if os.path.normcase(os.path.commonpath((work_dir, staged_path))) != os.path.normcase(work_dir):
            raise ValueError(f"MB备份文件路径越界: {archive_entry}")
        return staged_path

    def _discover_mb_backup_sources(self, archive):
        entries_by_name = {}
        duplicate_entries = set()
        for entry in archive.infolist():
            if entry.is_dir():
                continue
            if entry.filename in entries_by_name:
                duplicate_entries.add(entry.filename)
            entries_by_name[entry.filename] = entry

        discovered = []
        for category, source_entry, target_template in MB_BACKUP_FILE_SPECS:
            if source_entry in duplicate_entries:
                raise ValueError(f"MB备份包含重复资源条目: {source_entry}")
            entry = entries_by_name.get(source_entry)
            if entry is None:
                self.log(f"MB备份工作流: 跳过缺失资源 {source_entry}")
                continue
            discovered.append((category, entry, target_template, ''))

        ocres_entries = []
        for entry_name, entry in entries_by_name.items():
            if (
                entry_name.startswith(MB_BACKUP_OCRES_SOURCE_PREFIX)
                and entry_name.lower().endswith('.xml')
            ):
                if entry_name in duplicate_entries:
                    raise ValueError(f"MB备份包含重复资源条目: {entry_name}")
                relative_path = entry_name[len(MB_BACKUP_OCRES_SOURCE_PREFIX):]
                self._mb_backup_staging_path(tempfile.gettempdir(), relative_path)
                ocres_entries.append((entry, relative_path))

        if not ocres_entries:
            self.log(
                f"MB备份工作流: 跳过缺失资源 {MB_BACKUP_OCRES_SOURCE_PREFIX} 下的 XML"
            )
        else:
            for entry, relative_path in sorted(ocres_entries, key=lambda item: item[0].filename):
                discovered.append((
                    'ocres_string',
                    entry,
                    MB_BACKUP_OCRES_TARGET_TEMPLATE,
                    relative_path,
                ))
        return discovered

    def _create_mb_backup_archive_from_folder(self, backup_root, work_dir):
        archive_path = os.path.join(work_dir, 'source.zip')
        with zipfile.ZipFile(
            archive_path,
            'w',
            compression=zipfile.ZIP_DEFLATED,
            allowZip64=True
        ) as archive:
            for root_dir, directories, files in os.walk(backup_root):
                directories.sort()
                for file_name in sorted(files):
                    file_path = os.path.join(root_dir, file_name)
                    relative_path = os.path.relpath(file_path, backup_root)
                    archive_entry = relative_path.replace(os.sep, '/')
                    self._mb_backup_staging_path(work_dir, archive_entry)
                    archive.write(
                        file_path,
                        arcname=archive_entry,
                        compress_type=zipfile.ZIP_DEFLATED
                    )
        return archive_path

    def _choose_mb_backup_input(self):
        if self.mb_backup_source_type == 'folder':
            return filedialog.askdirectory(
                title="选择已解压的MB备份文件夹",
                mustexist=True
            )
        return filedialog.askopenfilename(
            title="选择MB备份ZIP文件",
            filetypes=[("MB备份 ZIP文件", "*.zip"), ("ZIP文件", "*.zip")]
        )

    def setup_mb_backup_sources(self, backup_path=None):
        if backup_path is None:
            backup_path = self._choose_mb_backup_input()
            if not backup_path:
                return

        backup_path = os.path.abspath(backup_path)
        work_dir = ''
        input_type = ''
        archive_path = ''
        output_base_path = backup_path
        if os.path.isdir(backup_path):
            input_type = 'folder'
            try:
                work_dir = tempfile.mkdtemp(prefix='translate_mb_backup_')
                archive_path = self._create_mb_backup_archive_from_folder(
                    backup_path,
                    work_dir
                )
            except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
                if work_dir and os.path.isdir(work_dir):
                    try:
                        shutil.rmtree(work_dir)
                    except OSError as cleanup_error:
                        self.log(f"清理 MB备份临时目录失败: {work_dir} -> {cleanup_error}")
                message = f"准备已解压MB备份文件夹失败: {error}"
                self.log(message)
                if not self.headless:
                    messagebox.showerror("MB备份工作流", message)
                return
        elif os.path.isfile(backup_path):
            if os.path.splitext(backup_path)[1].lower() != '.zip':
                message = f"MB备份工作流仅支持 ZIP 文件或已解压文件夹: {backup_path}"
                self.log(message)
                if not self.headless:
                    messagebox.showerror("MB备份工作流", message)
                return
            input_type = 'zip'
            archive_path = backup_path
        else:
            message = f"MB备份输入路径不存在: {backup_path}"
            self.log(message)
            if not self.headless:
                messagebox.showerror("MB备份工作流", message)
            return

        try:
            with zipfile.ZipFile(archive_path, 'r') as archive:
                discovered = self._discover_mb_backup_sources(archive)
                if not discovered:
                    if work_dir and os.path.isdir(work_dir):
                        try:
                            shutil.rmtree(work_dir)
                        except OSError as cleanup_error:
                            self.log(f"清理 MB备份临时目录失败: {work_dir} -> {cleanup_error}")
                    message = "MB备份中未找到可翻译的 CHS 资源。"
                    self.log(message)
                    if not self.headless:
                        messagebox.showwarning("MB备份工作流", message)
                    return

                if not work_dir:
                    work_dir = tempfile.mkdtemp(prefix='translate_mb_backup_')
                staged_sources = {}
                for category, entry, target_template, relative_path in discovered:
                    staged_path = self._mb_backup_staging_path(work_dir, entry.filename)
                    os.makedirs(os.path.dirname(staged_path), exist_ok=True)
                    with archive.open(entry, 'r') as source_file, open(staged_path, 'wb') as target_file:
                        shutil.copyfileobj(source_file, target_file)
                    staged_sources[staged_path] = MbBackupSource(
                        category=category,
                        archive_entry=entry.filename,
                        staged_path=staged_path,
                        target_template=target_template,
                        relative_path=relative_path,
                    )
        except (OSError, RuntimeError, ValueError, zipfile.BadZipFile) as error:
            if work_dir and os.path.isdir(work_dir):
                try:
                    shutil.rmtree(work_dir)
                except OSError as cleanup_error:
                    self.log(f"清理 MB备份临时目录失败: {work_dir} -> {cleanup_error}")
            message = f"准备MB备份工作流失败: {error}"
            self.log(message)
            if not self.headless:
                messagebox.showerror("MB备份工作流", message)
            return

        self._reset_simulator_state()
        self._reset_diskc_state()
        self._reset_mb_backup_state()
        self.mb_backup_mode = True
        self.mb_backup_input_type = input_type
        self.mb_backup_root_path = backup_path if input_type == 'folder' else ''
        self.mb_backup_archive_path = archive_path
        self.mb_backup_output_base_path = output_base_path
        self.mb_backup_work_dir = work_dir
        self.mb_backup_sources = staged_sources
        self.source_files = list(staged_sources)
        self.source_folders = []
        self.source_folder = ""
        self.output_dir = os.path.dirname(output_base_path)
        self.file_listbox.delete(0, tk.END)
        for source in staged_sources.values():
            self.file_listbox.insert(
                tk.END,
                f"[MB:{source.category}] {source.archive_entry}"
            )
        self.update_stats()
        input_label = "已解压文件夹" if input_type == 'folder' else "ZIP文件"
        message = (
            f"MB备份工作流已就绪（{input_label}）: "
            f"{len(staged_sources)} 个翻译资源"
        )
        self.log(message)
        if not self.headless:
            messagebox.showinfo("MB备份工作流", message)

    def setup_diskc_sources(self, diskc_root=None):
        if diskc_root is None:
            folder = filedialog.askdirectory(title="选择DiskC根目录")
            if not folder:
                return
            diskc_root = folder

        self._reset_mb_backup_state()
        self._reset_simulator_state()
        self._reset_diskc_state()
        self.diskc_root = diskc_root
        self.diskc_mode = True
        self.source_files = []
        self.file_listbox.delete(0, tk.END)
        self.source_folders = []
        self.source_folder = diskc_root
        self.output_dir = diskc_root

        res_path = os.path.join(diskc_root, 'OpenCNC', 'Bin', 'Language', 'CHS.res')
        prepared_res = self.prepare_file(res_path)
        if prepared_res:
            self.diskc_res_source = prepared_res
            self.source_files.append(prepared_res)
            self.file_listbox.insert(tk.END, f"[RES] {os.path.basename(res_path)} -> {prepared_res}")
            self.log(f"DiskC工作流: 加载 {res_path}")
        else:
            self.log(f"DiskC工作流: 未找到 {res_path}")

        string_dir = os.path.join(diskc_root, 'OpenCnc Shared', 'OCRes', 'CHS', 'String')
        if os.path.isdir(string_dir):
            xml_count = 0
            for root_dir, dirs, files in os.walk(string_dir):
                for file in files:
                    if file.lower().endswith('.xml'):
                        full_path = os.path.join(root_dir, file)
                        prepared = self.prepare_file(full_path)
                        if prepared:
                            rel_path = os.path.relpath(prepared, string_dir)
                            self.diskc_xml_map[prepared] = rel_path
                            self.source_files.append(prepared)
                            self.file_listbox.insert(tk.END, f"[XML] String/{rel_path}")
                            xml_count += 1
            self.log(f"DiskC工作流: 从String目录加载 {xml_count} 个XML文件")
        else:
            self.log(f"DiskC工作流: 未找到目录 {string_dir}")

        plugin_config_path = os.path.join(diskc_root, 'OpenCNC', 'Bin', 'Plugin', 'Config', 'CHS.xml')
        if os.path.isfile(plugin_config_path):
            prepared_plugin = self.prepare_file(plugin_config_path)
            if prepared_plugin:
                self.diskc_plugin_source = prepared_plugin
                self.source_files.append(prepared_plugin)
                self.file_listbox.insert(tk.END, f"[XML] Plugin/Config/CHS.xml")
                self.log(f"DiskC工作流: 加载 {plugin_config_path}")
        else:
            self.log(f"DiskC工作流: 未找到 {plugin_config_path}")

        msg = f"DiskC工作流已就绪: {len(self.source_files)} 个文件"
        if not self.headless:
            messagebox.showinfo("DiskC工作流", msg)
        self.log(msg)

    def prepare_file(self, file_path):
        """
        准备单个文件：自动识别类型，RES文件解压为无后缀名的XML文件并返回新路径
        返回可供后续处理的文件路径（XML文本文件路径），失败返回None
        """
        try:
            if file_path.lower().endswith('.res'):
                # 将RES解压写入到输出目录下的临时文件（无扩展名）
                base = os.path.splitext(os.path.basename(file_path))[0]
                conv_dir = os.path.join(self.output_dir, '_res_converted')
                os.makedirs(conv_dir, exist_ok=True)
                target_path = os.path.join(conv_dir, base)  # 无后缀名
                if os.path.exists(target_path) and os.path.getmtime(target_path) >= os.path.getmtime(file_path):
                    # 验证缓存文件是否为有效XML内容
                    try:
                        with open(target_path, 'r', encoding='utf-8') as tf:
                            preview = tf.read(200).strip()
                        if preview and (preview.startswith('<?xml') or preview.startswith('<')):
                            return target_path
                        else:
                            self.log(f"缓存文件 {os.path.basename(target_path)} 内容无效，重新解压")
                    except Exception:
                        self.log(f"缓存文件 {os.path.basename(target_path)} 读取失败，重新解压")
                content = self.decompress_res(file_path)
                if content is None:
                    self.log(f"RES解压失败: {file_path}")
                    if self.logger:
                        self.logger.error(f"RES解压失败: {file_path}")
                    return None
                with open(target_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.log(f"已转换RES -> 无后缀XML: {os.path.basename(target_path)}")
                return target_path
            else:
                # 仅接受XML或无后缀（可能是已转换的RES）
                if file_path.lower().endswith('.xml') or not os.path.splitext(file_path)[1]:
                    return file_path
                # 其它类型不支持
                self.log(f"跳过不支持的文件类型: {file_path}")
                if self.logger:
                    self.logger.warning(f"跳过不支持的文件类型: {file_path}")
                return None
        except Exception as e:
            self.log(f"准备文件失败: {file_path} -> {e}")
            if self.logger:
                self.logger.exception(e)
            return None
    
    def add_files(self):
        files = filedialog.askopenfilenames(title="选择XML或RES文件",
            filetypes=[("支持的源文件", "*.xml;*.res"), ("XML文件", "*.xml"), ("RES文件", "*.res"), ("所有文件", "*.*")])
        if not files:
            return
        if not self.source_files:
            first_file_dir = os.path.dirname(files[0])
            self.output_dir = os.path.dirname(first_file_dir)
            self.log(f"输出目录已设置为: {self.output_dir}")
        count = 0
        for file in files:
            prepared = self.prepare_file(file)
            if prepared and prepared not in self.source_files:
                self.source_files.append(prepared)
                tag = self._file_tag(prepared)
                self.file_listbox.insert(tk.END, f"{tag} {prepared}")
                count += 1
            else:
                self.log(f"跳过文件: {file}")
        if count:
            self.log(f"添加 {count} 个文件")

    @staticmethod
    def _file_tag(file_path):
        lower = file_path.lower()
        if lower.endswith('.res'):
            return '[RES]'
        elif '_res_converted' in lower:
            return '[RES→XML]'
        elif lower.endswith('.xml'):
            return '[XML]'
        elif not os.path.splitext(file_path)[1]:
            return '[无后缀]'
        return '[?]'
    
    def remove_files(self):
        """
        从列表中移除选中的文件
        """
        selected = self.file_listbox.curselection()
        for idx in reversed(selected):
            self.file_listbox.delete(idx)
            del self.source_files[idx]
    
    def clear_files(self):
        self.file_listbox.delete(0, tk.END)
        self.source_files = []
        self.source_folders = []
        self.source_folder = ""
        self._reset_diskc_state()
        self._reset_simulator_state()
        self._reset_mb_backup_state()
        self.update_stats()
    
    def toggle_lang_display(self):
        """
        切换语言显示模式
        显示所有语言或仅显示默认语言
        """
        self.show_all_langs = self.lang_show_var.get()
        self.update_lang_display()
    
    def update_lang_display(self):
        """
        更新语言复选框显示
        根据当前模式显示相应的语言选项
        """
        # 清除现有组件
        for widget in self.lang_frame.winfo_children():
            widget.destroy()
        
        self.lang_vars = {}
        self.main_lang_frames = {}
        
        display_langs = LANG_MAP if self.show_all_langs else {k: v for k, v in LANG_MAP.items() if k in self.default_langs}
        
        col = 0
        row_lang = 0
        for lang, info in sorted(display_langs.items()):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(self.lang_frame, text=f"{lang} - {info['name']}", variable=var)
            chk.grid(row=row_lang, column=col, sticky=tk.W, padx=5, pady=2)
            self.lang_vars[lang] = var
            self.main_lang_frames[lang] = chk
            col += 1
            if col >= 2:
                col = 0
                row_lang += 1
        
        self.lang_frame.update_idletasks()
        self.update_stats()
    
    def update_stats(self):
        if hasattr(self, 'stat_file_count'):
            self.stat_file_count.configure(text=str(len(self.source_files)))
        if hasattr(self, 'stat_lang_count'):
            selected = len([l for l, v in self.lang_vars.items() if v.get()]) if hasattr(self, 'lang_vars') else 0
            self.stat_lang_count.configure(text=str(selected))
        if hasattr(self, 'stats_refs') and 'shared' in self.stats_refs:
            total = len(self.translation_table)
            self.stats_refs['shared'].configure(text=str(total))
    
    def _filter_main_langs(self):
        if not hasattr(self, 'main_lang_frames'):
            return
        txt = self.main_lang_search_var.get().strip()
        if txt == '🔍 搜索...':
            txt = ''
        txt = txt.lower()
        import re as _re
        
        for code, chk in self.main_lang_frames.items():
            info = LANG_MAP.get(code, {})
            name = info.get('name', '')
            cn = ' '.join(_re.findall(r'[\u4e00-\u9fff]+', name)).lower()
            
            if not txt or txt in code.lower() or txt in name.lower() or txt in cn:
                chk.grid()
            else:
                chk.grid_remove()
    
    def select_all_langs(self):
        """
        全选所有语言
        """
        for var in self.lang_vars.values():
            var.set(True)
    
    def deselect_all_langs(self):
        """
        取消全选所有语言
        """
        for var in self.lang_vars.values():
            var.set(False)
    
    def set_default_langs(self):
        top = tk.Toplevel(self.root)
        top.title("设置默认语言")
        top.geometry("680x650")
        top.configure(bg='#F5F6FA')
        top.transient(self.root)
        top.grab_set()
        
        dlg_style = ttk.Style()
        if 'Dialog.TFrame' not in dlg_style.theme_names():
            dlg_style.configure('Dialog.TFrame', background='#F5F6FA')
            dlg_style.configure('DialogCard.TFrame', background='white')
            dlg_style.configure('DialogTitle.TLabel', background='#F5F6FA', foreground='#2D2B4E',
                                font=('Segoe UI', 16, 'bold'))
            dlg_style.configure('DialogSub.TLabel', background='#F5F6FA', foreground='#8E8EA0',
                                font=('Segoe UI', 10))
            dlg_style.configure('Search.TEntry', font=('Segoe UI', 10), padding=8)
            dlg_style.configure('Status.TLabel', background='#F5F6FA', foreground='#6C63FF',
                                font=('Segoe UI', 10, 'bold'))
            dlg_style.configure('LangItem.TCheckbutton', background='white', font=('Segoe UI', 9),
                                padding=4)
            dlg_style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'), padding=10)
            dlg_style.configure('TButton', font=('Segoe UI', 9), padding=6)
            dlg_style.configure('Match.TFrame', background='#EDE7FF')
        
        main = ttk.Frame(top, style='Dialog.TFrame', padding=24)
        main.pack(fill=tk.BOTH, expand=True)
        
        title_frame = ttk.Frame(main, style='Dialog.TFrame')
        title_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(title_frame, text="设置默认语言", style='DialogTitle.TLabel').pack(anchor='w')
        ttk.Label(title_frame, text="选择最多 20 种常用语言，用于快速开始翻译任务", 
                  style='DialogSub.TLabel').pack(anchor='w', pady=(4, 0))
        
        search_card = ttk.Frame(main, style='DialogCard.TFrame', padding=12)
        search_card.pack(fill=tk.X, pady=(0, 12))
        
        search_row = ttk.Frame(search_card, style='DialogCard.TFrame')
        search_row.pack(fill=tk.X)
        
        ttk.Label(search_row, text="🔍", font=('Segoe UI', 14), background='white').pack(side='left')
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=search_var, width=40, style='Search.TEntry')
        search_entry.pack(side='left', padx=(8, 0), fill=tk.X, expand=True)
        search_entry.insert(0, '搜索语言代码、英文或中文国家名...')
        search_entry.configure(foreground='#A5A0C0')
        
        def on_search_focus_in(e):
            if search_entry.get() == '搜索语言代码、英文或中文国家名...':
                search_entry.delete(0, tk.END)
                search_entry.configure(foreground='#2D2B4E')
        
        def on_search_focus_out(e):
            if not search_entry.get().strip():
                search_entry.insert(0, '搜索语言代码、英文或中文国家名...')
                search_entry.configure(foreground='#A5A0C0')
        
        search_entry.bind('<FocusIn>', on_search_focus_in)
        search_entry.bind('<FocusOut>', on_search_focus_out)
        
        list_card = ttk.Frame(main, style='DialogCard.TFrame', padding=16)
        list_card.pack(fill=tk.BOTH, expand=True, pady=(0, 12))
        
        canvas_frame = ttk.Frame(list_card, style='DialogCard.TFrame')
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas, style='DialogCard.TFrame')
        
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw", width=600)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        all_langs = [(code, info) for code, info in sorted(LANG_MAP.items())]
        lang_vars = {}
        lang_frames = {}
        
        status_bar = ttk.Frame(main, style='Dialog.TFrame')
        status_bar.pack(fill=tk.X, pady=(0, 16))
        status_label = ttk.Label(status_bar, text=f"已选择 {len(self.default_langs)}/20 种语言",
                                 style='Status.TLabel')
        status_label.pack(side='left')
        
        def update_status():
            c = sum(1 for v in lang_vars.values() if v.get())
            status_label.config(text=f"已选择 {c}/20 种语言")
        
        def on_check(code):
            v = lang_vars[code]
            if v.get() and sum(1 for x in lang_vars.values() if x.get()) > 20:
                v.set(False)
                messagebox.showwarning("提示", "最多只能选择 20 种语言", parent=top)
            update_status()
        
        for i, (code, info) in enumerate(all_langs):
            var = tk.BooleanVar(value=code in self.default_langs)
            lang_vars[code] = var
            
            item = ttk.Frame(scrollable, style='DialogCard.TFrame')
            item.pack(fill=tk.X, pady=2, padx=4)
            
            hover_bg = '#F0EEFF'
            normal_bg = 'white'
            
            def on_enter(e, frm=item):
                try: frm.configure(style='Hover.TFrame')
                except: pass
            def on_leave(e, frm=item):
                try: frm.configure(style='DialogCard.TFrame')
                except: pass
            
            item.bind('<Enter>', on_enter)
            item.bind('<Leave>', on_leave)
            
            chk = ttk.Checkbutton(item, text=f"{code}  {info['name']}", variable=var,
                                   style='LangItem.TCheckbutton',
                                   command=lambda c=code: on_check(c))
            chk.pack(anchor='w', padx=8, pady=6)
            lang_frames[code] = item
        
        def extract_cn_text(name):
            import re as _re
            cn = _re.findall(r'[\u4e00-\u9fff]+', name)
            return ' '.join(cn).lower()
        
        match_count = 0
        
        def on_search(*args):
            nonlocal match_count
            txt = search_var.get().lower().strip()
            
            if txt == '搜索语言代码、英文或中文国家名...'.lower():
                txt = ''
            
            first_match = None
            match_count = 0
            
            for code, frm in lang_frames.items():
                info = LANG_MAP[code]
                full_text = f"{code} {info['name']}".lower()
                cn_text = extract_cn_text(info['name'])
                
                if txt:
                    matched = (txt in full_text or txt in code.lower() or 
                               txt in cn_text or txt in info['name'].lower())
                    
                    if matched:
                        frm.pack(fill=tk.X, pady=2, padx=4)
                        try: frm.configure(style='Match.TFrame')
                        except: pass
                        if first_match is None:
                            first_match = frm
                        match_count += 1
                    else:
                        frm.pack_forget()
                else:
                    frm.pack(fill=tk.X, pady=2, padx=4)
                    try: frm.configure(style='DialogCard.TFrame')
                    except: pass
            
            if txt:
                status_label.config(text=f"已选择 {sum(1 for v in lang_vars.values() if v.get())}/20 种语言  |  找到 {match_count} 个匹配")
            else:
                status_label.config(text=f"已选择 {sum(1 for v in lang_vars.values() if v.get())}/20 种语言")
            
            if first_match:
                canvas.update_idletasks()
                try:
                    y = canvas.coords(canvas.find_withtag('all')[0])[1] if canvas.find_withtag('all') else 0
                    canvas.yview_moveto((frm.winfo_rooty() - scrollable.winfo_rooty() + 20) / scrollable.winfo_height())
                except Exception:
                    pass
        
        search_var.trace('w', on_search)
        
        btn_bar = ttk.Frame(main, style='Dialog.TFrame')
        btn_bar.pack(fill=tk.X)
        
        def deselect_all():
            for v in lang_vars.values(): v.set(False)
            update_status()
        
        ttk.Button(btn_bar, text="取消全选", command=deselect_all).pack(side='left')
        ttk.Separator(btn_bar, orient='vertical').pack(side='left', fill='y', padx=12, pady=4)
        
        btn_frame_right = ttk.Frame(btn_bar, style='Dialog.TFrame')
        btn_frame_right.pack(side='right')
        
        ttk.Button(btn_frame_right, text="取消", command=top.destroy).pack(side='right', padx=4)
        ok_btn = ttk.Button(btn_frame_right, text="确定保存",
                              command=lambda: self._confirm_default_langs_from_vars(top, lang_vars),
                              style='Accent.TButton')
        ok_btn.pack(side='right', padx=4)
        
        def on_mousewheel(e): canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        def on_close():
            canvas.unbind_all("<MouseWheel>")
            top.destroy()
        top.protocol("WM_DELETE_WINDOW", on_close)
    
    def _confirm_default_langs_from_vars(self, top, lang_check_vars):
        """
        从勾选框变量确认默认语言设置
        """
        selected_langs = [code for code, var in lang_check_vars.items() if var.get()]
        
        if selected_langs:
            self.default_langs = selected_langs[:20]
            self.save_config()
            self.show_all_langs = False
            self.lang_show_var.set(False)
            self.update_lang_display()
            messagebox.showinfo("提示", f"已设置 {len(self.default_langs)} 种默认语言")
        else:
            messagebox.showwarning("警告", "请至少选择一种语言")
        
        top.destroy()
    
    def load_translation_table(self):
        """
        加载翻译表
        从输出目录读取translation_table.json
        """
        table_path = os.path.join(self.table_dir, 'translation_table.json')
        if os.path.exists(table_path):
            try:
                with open(table_path, 'r', encoding='utf-8') as f:
                    self.translation_table = json.load(f)
                self.log(f"已加载翻译表，共 {len(self.translation_table)} 条记录")
            except Exception as e:
                self.log(f"加载翻译表失败: {e}")
    
    def save_translation_table(self):
        """
        保存翻译表
        将翻译表写入translation_table.json
        """
        table_path = os.path.join(self.table_dir, 'translation_table.json')
        temp_path = None
        try:
            file_descriptor, temp_path = tempfile.mkstemp(
                prefix='.translation_table.',
                suffix='.tmp',
                dir=self.table_dir
            )
            os.close(file_descriptor)
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(self.translation_table, f, ensure_ascii=False, indent=2)
            os.replace(temp_path, table_path)
            temp_path = None
            self.log("翻译表已保存")
        except (OSError, TypeError, ValueError) as e:
            self.log(f"保存翻译表失败: {e}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError as cleanup_error:
                    self.log(f"清理临时翻译表失败: {cleanup_error}")
    
    def log(self, message):
        """
        添加日志消息
        :param message: 日志消息内容
        """
        line = f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n"
        if self.logger:
            self.logger.info(message)
        if self.headless:
            if not self.logger:
                print(line, end='')
            return
        if not hasattr(self, 'log_text') or threading.get_ident() != self._main_thread_id:
            self._log_queue.put(line)
            return
        self._append_log_line(line)
    
    def extract_chinese_texts(self):
        """
        从XML或RES文件中提取中文文本
        RES文件会自动使用GZIP格式解压
        :return: 中文文本集合
        """
        all_texts = set()
        
        for source_file in self.source_files:
            try:
                # 当前source_files应为XML文本文件路径（可能为无后缀的已转换RES）
                with open(source_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                parser = ET.XMLParser(encoding='utf-8')
                try:
                    root = ET.fromstring(content, parser=parser)
                    for elem in root.findall('.//Message'):
                        text = elem.get('Content', '')
                        if text and self.is_chinese(text):
                            all_texts.add(text)
                    self.log(f"从 {os.path.basename(source_file)} 提取到中文文本")
                except Exception as pe:
                    # 容错：XML 解析失败，尝试通过正则抽取 Message 的 Content 属性
                    self.log(f"解析 {os.path.basename(source_file)} 时 XML 解析失败: {pe}, 尝试正则容错抽取")
                    if self.logger:
                        self.logger.exception(pe)

                    try:
                        # 匹配 <Message ... Content="..." .../> 或 Content='...'
                        msgs = []
                        for m in re.finditer(r'<Message\b([^>]*)>', content, flags=re.IGNORECASE | re.DOTALL):
                            attrs = m.group(1)
                            # 尝试双引号
                            m1 = re.search(r'Content\s*=\s*"(.*?)"', attrs, flags=re.DOTALL)
                            if m1:
                                import html
                                msgs.append(html.unescape(m1.group(1)))
                                continue
                            m2 = re.search(r"Content\s*=\s*'(.*?)'", attrs, flags=re.DOTALL)
                            if m2:
                                import html
                                msgs.append(html.unescape(m2.group(1)))

                        for text in msgs:
                            if text and self.is_chinese(text):
                                all_texts.add(text)

                        self.log(f"正则方式从 {os.path.basename(source_file)} 提取到 {len(msgs)} 条 Message")
                        if len(msgs) == 0:
                            content_preview = content.strip()[:120]
                            self.log(f"警告: {os.path.basename(source_file)} 未匹配到任何Message标签，文件预览: {content_preview}")
                    except Exception as re_e:
                        self.log(f"正则抽取失败: {re_e}")
                        if self.logger:
                            self.logger.exception(re_e)
            except Exception as e:
                self.log(f"解析 {source_file} 失败: {e}")
                if self.logger:
                    self.logger.exception(f"解析失败: {source_file}")
        # 将新文本添加到翻译表（标准化处理）
        for text in all_texts:
            normalized = self._normalize_text(text)
            if normalized not in self.translation_table:
                self.translation_table[normalized] = {lang: '' for lang in LANG_MAP.keys()}
        
        self.save_translation_table()
        return all_texts
    

    
    def _normalize_text(self, text):
        """文本标准化：统一文本格式以确保准确匹配"""
        if not text:
            return text
        
        result = text.strip()
        result = result.replace('\r\n', '\n').replace('\r', '\n')
        result = result.replace('，', ',').replace('。', '.').replace('：', ':')
        result = result.replace('；', ';').replace('！', '!').replace('？', '?')
        result = result.replace('（', '(').replace('）', ')')
        result = re.sub(r'\s+', ' ', result)
        
        return result
    
    def decompress_res(self, res_file_path):
        """
        解压RES文件（GZIP格式）
        :param res_file_path: RES文件路径
        :return: 解压后的内容（字符串），失败返回None
        """
        try:
            with open(res_file_path, 'rb') as f:
                head = f.read(2)
                f.seek(0)
                if head != b'\x1f\x8b':
                    self.log(f"警告: {os.path.basename(res_file_path)} 不是有效的GZIP文件")
                    return None
                import gzip
                with gzip.GzipFile(fileobj=f) as gz:
                    data = gz.read()

            # 先尝试utf-8解码，失败则使用替代策略
            try:
                text = data.decode('utf-8')
            except UnicodeDecodeError:
                # 记录并尝试使用替代解码
                self.log(f"警告: {os.path.basename(res_file_path)} UTF-8 解码失败，使用替代解码")
                if self.logger:
                    self.logger.exception(f"UTF-8 解码失败: {res_file_path}")
                try:
                    text = data.decode('utf-8', errors='replace')
                except Exception:
                    text = data.decode('latin-1', errors='replace')

            stripped = text.strip()
            if not stripped:
                self.log(f"错误: {os.path.basename(res_file_path)} 解压后内容为空")
                return None
            if not (stripped.startswith('<?xml') or stripped.startswith('<')):
                self.log(f"错误: {os.path.basename(res_file_path)} 解压后内容不是有效XML（预览: {stripped[:80]}）")
                return None
            return text
        except Exception as e:
            self.log(f"解压错误: {e}")
            if self.logger:
                self.logger.exception(e)
            return None
    
    def is_chinese(self, text):
        """
        判断文本是否包含中文
        :param text: 待检查文本
        :return: True表示包含中文，False表示不包含
        """
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False

    def _provider_label(self, provider):
        return API_PROVIDERS.get(provider, API_PROVIDERS['google'])['label']

    def _provider_from_label(self, label):
        for provider, spec in API_PROVIDERS.items():
            if spec['label'] == label:
                return provider
        return 'google'

    def _provider_config(self, provider, overrides=None):
        config = dict(self.provider_configs.get(provider, {}))
        if overrides:
            config.update(overrides)
        return config

    def _missing_provider_fields(self, provider, config=None):
        values = self._provider_config(provider, config)
        return [
            field['label']
            for field in API_PROVIDERS[provider]['fields']
            if field.get('required') and not values.get(field['key'], '').strip()
        ]

    def _get_http_session(self, use_environment_proxy=True):
        session_name = 'session' if use_environment_proxy else 'direct_session'
        session = getattr(self._http_local, session_name, None)
        if session is None:
            session = requests.Session()
            # Keep the normal session compatible with HTTPS_PROXY/HTTP_PROXY.
            # A separate direct session is used when a proxy breaks TLS.
            session.trust_env = use_environment_proxy
            session.headers.update({
                'User-Agent': 'TranslatePro/1.3 (+https://github.com/)',
                'Accept': 'application/json',
            })
            setattr(self._http_local, session_name, session)
        return session

    def _provider_worker_count(self):
        return API_PROVIDER_WORKERS.get(self.api_type, 4)

    def _provider_requests_per_second(self, provider):
        return API_PROVIDER_REQUESTS_PER_SECOND.get(provider)

    def _wait_for_provider_slot(self, provider):
        requests_per_second = self._provider_requests_per_second(provider)
        if not requests_per_second:
            return

        interval = 1.0 / requests_per_second
        with self._provider_rate_lock:
            now = time.monotonic()
            next_request_at = self._provider_next_request_at.get(provider, now)
            scheduled_at = max(now, next_request_at)
            self._provider_next_request_at[provider] = scheduled_at + interval
        wait_seconds = scheduled_at - now
        if wait_seconds > 0:
            time.sleep(wait_seconds)

    def _append_log_line(self, line):
        self.log_text.insert(tk.END, line)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _flush_pending_logs(self):
        while True:
            try:
                line = self._log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log_line(line)
        self.root.after(100, self._flush_pending_logs)

    def _run_on_ui(self, callback):
        if self.headless:
            return
        if threading.get_ident() == self._main_thread_id:
            callback()
            return
        self._ui_queue.put(callback)

    def _flush_pending_ui_tasks(self):
        while True:
            try:
                callback = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            callback()
        self.root.after(50, self._flush_pending_ui_tasks)
    
    def on_api_type_changed(self):
        """
        API类型切换处理
        保存用户选择的API类型
        """
        self.api_type = self._provider_from_label(self.api_type_var.get())
        self.api_type_var.set(self._provider_label(self.api_type))
        self.save_config()
        provider = API_PROVIDERS[self.api_type]
        self.log(f"当前翻译API: {provider['label']}")
        missing = self._missing_provider_fields(self.api_type)
        if missing:
            detail = f"请先配置 {', '.join(missing)}"
            self.log(f"警告: {provider['label']} {detail}")
            messagebox.showwarning("API配置", detail)
        elif self.api_type == 'google':
            self.log("警告: Google GTX 是非官方公共接口，仅建议用于小批量连通性测试。")
    
    def configure_api(self):
        """
        配置翻译API
        各服务的密钥仅保存在本机 translation_config.json 中。
        """
        top = tk.Toplevel(self.root)
        top.title("API配置")
        top.geometry("680x470")
        top.minsize(620, 420)
        top.transient(self.root)
        top.grab_set()
        
        main_frame = ttk.Frame(top, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="翻译服务配置", font=('Arial', 12, 'bold')).pack(anchor=tk.W)
        ttk.Label(
            main_frame,
            text="免费/自建服务排在前面。请先填写配置后使用“测试连接”验证。",
            font=('Arial', 9),
            foreground='gray'
        ).pack(anchor=tk.W, pady=(4, 12))

        selector_frame = ttk.Frame(main_frame)
        selector_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(selector_frame, text="配置服务:").pack(side=tk.LEFT)
        selected_provider_var = tk.StringVar(value=self._provider_label(self.api_type))
        provider_combo = ttk.Combobox(
            selector_frame,
            textvariable=selected_provider_var,
            values=[spec['label'] for spec in API_PROVIDERS.values()],
            state='readonly',
            width=34
        )
        provider_combo.pack(side=tk.LEFT, padx=(8, 0))

        description_var = tk.StringVar()
        ttk.Label(
            main_frame,
            textvariable=description_var,
            font=('Arial', 9),
            foreground='gray',
            wraplength=620,
            justify=tk.LEFT
        ).pack(anchor=tk.W, pady=(0, 10))

        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True)
        status_var = tk.StringVar()
        ttk.Label(main_frame, textvariable=status_var, foreground='#6C63FF').pack(anchor=tk.W, pady=(8, 4))

        draft_configs = {provider: dict(values) for provider, values in self.provider_configs.items()}
        field_vars = {}
        current_provider = [self.api_type]

        def save_current_fields():
            provider = current_provider[0]
            if provider not in field_vars:
                return
            for key, value_var in field_vars[provider].items():
                draft_configs[provider][key] = value_var.get().strip()

        def render_fields(event=None):
            save_current_fields()
            provider = self._provider_from_label(selected_provider_var.get())
            current_provider[0] = provider
            selected_provider_var.set(self._provider_label(provider))
            description_var.set(API_PROVIDERS[provider]['description'])
            for child in fields_frame.winfo_children():
                child.destroy()

            variables = {}
            fields = API_PROVIDERS[provider]['fields']
            if not fields:
                ttk.Label(
                    fields_frame,
                    text="此模式不需要密钥。它是非官方公共接口，仅用于小批量测试。",
                    foreground='#A35A00'
                ).pack(anchor=tk.W, pady=8)
            for field in fields:
                row = ttk.Frame(fields_frame)
                row.pack(fill=tk.X, pady=5)
                required = " *" if field.get('required') else ""
                ttk.Label(row, text=f"{field['label']}{required}:", width=18).pack(side=tk.LEFT)
                value_var = tk.StringVar(value=draft_configs[provider].get(field['key'], ''))
                variables[field['key']] = value_var
                ttk.Entry(
                    row,
                    textvariable=value_var,
                    width=62,
                    show='*' if field.get('secret') else ''
                ).pack(side=tk.LEFT, fill=tk.X, expand=True)
            field_vars[provider] = variables
            status_var.set("")

        def test_connection():
            save_current_fields()
            provider = current_provider[0]
            config = dict(draft_configs[provider])
            missing = self._missing_provider_fields(provider, config)
            if missing:
                status_var.set(f"请先填写: {', '.join(missing)}")
                return

            test_button.config(state='disabled')
            status_var.set("正在测试连接...")

            def run_test():
                result = self._translate_with_provider(
                    provider,
                    "翻译连接测试",
                    "en",
                    config_override=config,
                    use_cache=False
                )

                def show_result():
                    test_button.config(state='normal')
                    if result.success:
                        status_var.set(f"连接成功，测试结果: {result.text}")
                    else:
                        status_var.set(f"连接失败: {result.error}")

                self._run_on_ui(show_result)

            threading.Thread(target=run_test, daemon=True).start()

        def save_and_close():
            save_current_fields()
            self.provider_configs = {provider: dict(values) for provider, values in draft_configs.items()}
            self._sync_legacy_baidu_credentials()
            self.save_config()
            messagebox.showinfo("提示", "API配置已保存！")
            top.destroy()

        provider_combo.bind('<<ComboboxSelected>>', render_fields)
        render_fields()

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(anchor=tk.E, pady=(8, 0))
        
        test_button = ttk.Button(btn_frame, text="测试连接", command=test_connection)
        test_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="保存", command=save_and_close).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=top.destroy).pack(side=tk.LEFT, padx=5)
    
    def baidu_translate(self, text, target_lang, config_override=None):
        """
        使用百度翻译API翻译文本
        :param text: 待翻译文本
        :param target_lang: 目标语言代码
        :return: 翻译后的文本
        """
        config = self._provider_config('baidu', config_override)
        app_id = config.get('app_id', '')
        secret_key = config.get('secret_key', '')
        if not app_id or not secret_key:
            raise TranslationProviderError("百度翻译缺少 App ID 或 Secret Key。")
        if len(text.encode('utf-8')) > 6000:
            raise TranslationProviderError("百度翻译单次请求不能超过 6000 个 UTF-8 字节。")
        
        # 百度翻译API支持的语言列表
        supported_langs = {
            'auto', 'zh', 'en', 'yue', 'wyw', 'ja', 'ko', 'fr', 'de', 'es', 'pt', 'pt-PT',
            'vi', 'id', 'th', 'ru', 'ar', 'hi', 'bn', 'pa', 'gu', 'or', 'ta', 'te', 'kn',
            'ml', 'mr', 'sa', 'yi', 'hi-IN', 'ne', 'my', 'si', 'km', 'lo', 'bo', 'ug',
            'ca', 'eu', 'gl', 'it', 'el', 'hu', 'cs', 'sk', 'sl', 'pl', 'ro', 'bg',
            'sr', 'hr', 'lv', 'lt', 'et', 'fi', 'sv', 'da', 'no', 'is', 'mt', 'tr',
            'hi-Latn', 'ur', 'fa', 'sq', 'bs', 'ms', 'ms-Arab', 'tl', 'ht', 'sw',
            'xh', 'zu', 'af', 'zu', 'xh', 'st', 'tn', 'ts', 'ss', 'nr', 've', 'wo',
            'rm', 'ga', 'gd', 'cy', 'kw', 'br', 'gv', 'gd', 'ga', 'eu', 'gl', 'ast',
            'ca', 'oc', 'fr-CA', 'it-CH', 'de-CH', 'gsw', 'als', 'lb', 'wa', 'nl',
            'fy', 'pap', 'sv-SE', 'sv-FI', 'nb', 'nn', 'da', 'is', 'fo', 'et', 'lv',
            'lt', 'pl', 'cs', 'sk', 'sl', 'hr', 'sr', 'bg', 'mk', 'sq', 'mt', 'tr',
            'az', 'uz', 'kk', 'ky', 'tg', 'tt', 'ba', 'cv', 'ce', 'yi', 'he', 'ar',
            'fa', 'ur', 'ps', 'sd', 'ku', 'ckb', 'syc', 'dv', 'hi', 'bn', 'as',
            'gu', 'or', 'ta', 'te', 'kn', 'ml', 'pa', 'mr', 'sa', 'kok', 'mai',
            'ne', 'si', 'my', 'km', 'lo', 'th', 'ja', 'ko', 'zh', 'zh-TW', 'zh-HK',
            'zh-CN', 'yue', 'wyw', 'vi', 'id', 'tl', 'ms', 'en', 'de', 'fr', 'es',
            'pt', 'ru', 'it', 'pl', 'ro', 'bg', 'cs', 'da', 'nl', 'fi', 'el', 'hu',
            'no', 'pt-PT', 'sv', 'tr', 'ar', 'hi', 'ja', 'ko', 'th', 'vi', 'zh',
            'zh-TW', 'en', 'de', 'fr', 'es', 'pt', 'ru', 'it', 'pl', 'ro', 'bg',
            'cs', 'da', 'nl', 'fi', 'el', 'hu', 'no', 'pt-PT', 'sv', 'tr', 'ar',
            'hi', 'ja', 'ko', 'th', 'vi', 'zh', 'zh-TW', 'en', 'de', 'fr', 'es',
            'pt', 'ru', 'it', 'pl', 'ro', 'bg', 'cs', 'da', 'nl', 'fi', 'el', 'hu',
            'no', 'pt-PT', 'sv', 'tr'
        }
        
        # 调试日志
        self.log(f"百度翻译: target_lang={target_lang}")
        
        # 百度翻译API语言代码映射表
        baidu_lang_map = {
            # 中文变体
            'zh-CN': 'zh', 'zh-TW': 'cht', 'zh-HK': 'zh', 'zh-SG': 'zh', 'zh-MO': 'zh',
            # 阿拉伯语变体
            'ar-SA': 'ar', 'ar-IQ': 'ar', 'ar-EG': 'ar', 'ar-LY': 'ar', 'ar-DZ': 'ar',
            'ar-MA': 'ar', 'ar-TN': 'ar', 'ar-OM': 'ar', 'ar-YE': 'ar', 'ar-SY': 'ar',
            'ar-JO': 'ar', 'ar-LB': 'ar', 'ar-KW': 'ar', 'ar-AE': 'ar', 'ar-BH': 'ar',
            'ar-QA': 'ar',
            # 德语变体
            'de-DE': 'de', 'de-CH': 'de', 'de-AT': 'de', 'de-LU': 'de', 'de-LI': 'de',
            # 英语变体
            'en-US': 'en', 'en-GB': 'en', 'en-AU': 'en', 'en-CA': 'en', 'en-NZ': 'en',
            'en-IE': 'en', 'en-ZA': 'en',
            # 西班牙语变体
            'es-ES': 'es', 'es-MX': 'es', 'es-GT': 'es', 'es-CR': 'es', 'es-PA': 'es',
            'es-DO': 'es', 'es-VE': 'es', 'es-CO': 'es', 'es-PE': 'es', 'es-AR': 'es',
            'es-EC': 'es', 'es-CL': 'es', 'es-UY': 'es', 'es-PY': 'es', 'es-BO': 'es',
            'es-SV': 'es', 'es-HN': 'es', 'es-NI': 'es', 'es-PR': 'es',
            # 法语变体
            'fr-FR': 'fr', 'fr-BE': 'fr', 'fr-CA': 'fr', 'fr-CH': 'fr', 'fr-LU': 'fr',
            'fr-MC': 'fr',
            # 意大利语变体
            'it-IT': 'it', 'it-CH': 'it',
            # 荷兰语变体
            'nl-NL': 'nl', 'nl-BE': 'nl',
            # 葡萄牙语变体
            'pt-PT': 'pt', 'pt-BR': 'pt',
            # 塞尔维亚语变体
            'sr-Latn': 'sr', 'sr-Cyrl': 'sr',
            # 瑞典语变体
            'sv-SE': 'sv', 'sv-FI': 'sv',
            # 乌兹别克语变体
            'uz-Latn': 'uz',
            # 阿塞拜疆语变体
            'az-Latn': 'az',
            # 其他语言
            'ca': 'ca', 'bg': 'bg', 'cs': 'cs', 'da': 'da', 'el': 'el', 'eo': 'eo',
            'et': 'et', 'fi': 'fi', 'he': 'he', 'hu': 'hu', 'is': 'is', 'ja': 'ja',
            'ko': 'ko', 'lv': 'lv', 'lt': 'lt', 'mk': 'mk', 'no': 'no', 'nn': 'no',
            'nb': 'no', 'pl': 'pl', 'ro': 'ro', 'ru': 'ru', 'hr': 'hr', 'sk': 'sk',
            'sl': 'sl', 'sq': 'sq', 'th': 'th', 'tr': 'tr', 'ur-PK': 'ur', 'id': 'id',
            'uk': 'uk', 'be': 'be', 'fa': 'fa', 'vi': 'vi', 'hy': 'hy', 'eu': 'eu',
            'af': 'af', 'ka': 'ka', 'fo': 'fo', 'hi': 'hi', 'ms-MY': 'ms', 'ms-BN': 'ms',
            'kk': 'kk', 'sw': 'sw', 'tt': 'tt', 'bn': 'bn', 'pa': 'pa', 'gu': 'gu',
            'or': 'or', 'ta': 'ta', 'te': 'te', 'kn': 'kn', 'ml': 'ml', 'as': 'as',
            'mr': 'mr', 'sa': 'sa', 'kok': 'kok'
        }
        
        # 获取百度语言代码
        if target_lang in baidu_lang_map:
            lang_code = baidu_lang_map[target_lang]
            self.log(f"百度翻译: 映射成功 {target_lang} -> {lang_code}")
        else:
            # 尝试仅使用语言代码部分
            base_lang = target_lang.split('-')[0] if '-' in target_lang else target_lang
            if base_lang in baidu_lang_map:
                lang_code = baidu_lang_map[base_lang]
                self.log(f"百度翻译: 使用基础语言代码 {target_lang} -> {base_lang} -> {lang_code}")
            else:
                self.log(f"警告: 百度翻译不支持语言 '{target_lang}'，使用英语")
                lang_code = 'en'
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                salt = str(uuid.uuid4())[:10]
                sign = app_id + text + salt + secret_key
                sign = hashlib.md5(sign.encode('utf-8')).hexdigest()
                
                url = 'https://fanyi-api.baidu.com/api/trans/vip/translate'
                data = {
                    'q': text,
                    'from': 'auto',
                    'to': lang_code,
                    'appid': app_id,
                    'salt': salt,
                    'sign': sign
                }
                
                self.log(f"百度翻译请求: from=auto, to={lang_code}, text_len={len(text)}")
                
                response = self._get_http_session().post(url, data=data, timeout=10)
                response.raise_for_status()
                result = response.json()
                
                if 'trans_result' in result and result['trans_result']:
                    return result['trans_result'][0]['dst']
                else:
                    if 'error_code' in result:
                        error_code = result['error_code']
                        error_msg = result.get('error_msg', '未知错误')
                        self.log(f"百度翻译API错误[{error_code}]: {error_msg}")
                        if error_code == '58001':
                            self.log(f"提示: 请检查百度翻译API账户是否开通了 {lang_code} 语言的翻译服务")
                    return None
            except (requests.RequestException, ValueError) as e:
                if attempt < max_retries - 1:
                    time.sleep((2 ** attempt) + random.uniform(0, 0.5))
                    continue
                self.log(f"百度翻译失败: {text} -> {target_lang}; 原因: {e}")
                return None
        
        return None
    
    def _base_language_code(self, target_lang):
        code = target_lang.replace('_', '-').lower()
        if code.startswith('zh-'):
            return 'zh'
        return code.split('-', 1)[0]

    def _provider_language_code(self, provider, target_lang):
        code = target_lang.replace('_', '-').lower()
        if provider in ('deepl_free', 'deepl_pro'):
            deepl_codes = {
                'en-us': 'EN-US', 'en-gb': 'EN-GB',
                'pt-br': 'PT-BR', 'pt-pt': 'PT-PT',
                'nb': 'NB', 'nn': 'NB', 'no': 'NB',
            }
            if code in deepl_codes:
                return deepl_codes[code]
            return self._base_language_code(code).upper()
        if provider == 'youdao':
            chinese_codes = {
                'zh-cn': 'zh-CHS', 'zh-sg': 'zh-CHS',
                'zh-tw': 'zh-CHT', 'zh-hk': 'zh-CHT', 'zh-mo': 'zh-CHT',
            }
            return chinese_codes.get(code, self._base_language_code(code))
        if provider == 'tencent':
            chinese_codes = {
                'zh-cn': 'zh', 'zh-sg': 'zh',
                'zh-tw': 'zh-TW', 'zh-hk': 'zh-TW', 'zh-mo': 'zh-TW',
            }
            return chinese_codes.get(code, self._base_language_code(code))
        if provider == 'google':
            return target_lang
        return self._base_language_code(code)

    def _require_provider_config(self, provider, config):
        missing = self._missing_provider_fields(provider, config)
        if missing:
            raise TranslationProviderError(f"缺少配置: {', '.join(missing)}。")

    def _request_json(self, method, url, *, data=None, json_data=None, headers=None):
        session = self._get_http_session()
        request_kwargs = {
            'data': data,
            'json': json_data,
            'headers': headers,
            'timeout': 10,
        }
        try:
            response = session.request(method, url, **request_kwargs)
        except requests.exceptions.SSLError as exc:
            # Some local proxies close the TLS handshake to DeepL with EOF.
            # Preserve proxy support, but retry once with a clean direct session.
            if not getattr(session, 'trust_env', True):
                raise TranslationProviderError(
                    f"网络请求失败: {exc}",
                    retryable=True
                ) from exc
            try:
                direct_session = self._get_http_session(use_environment_proxy=False)
                response = direct_session.request(method, url, **request_kwargs)
            except requests.RequestException as direct_exc:
                raise TranslationProviderError(
                    f"网络请求失败: {exc}；绕过系统代理后仍失败: {direct_exc}",
                    retryable=True
                ) from direct_exc
            self.log("检测到系统代理导致 TLS 握手失败，已自动改用直连请求。")
        except requests.RequestException as exc:
            raise TranslationProviderError(f"网络请求失败: {exc}", retryable=True) from exc

        if not response.ok:
            response_text = response.text.replace('\n', ' ').strip()[:240]
            retryable = response.status_code == 429 or response.status_code >= 500
            detail = f"HTTP {response.status_code}"
            if response_text:
                detail = f"{detail}: {response_text}"
            raise TranslationProviderError(detail, retryable=retryable)

        try:
            return response.json()
        except ValueError as exc:
            raise TranslationProviderError("服务返回了无法解析的 JSON 响应。") from exc

    def _require_translation(self, value, provider):
        if not isinstance(value, str) or not value.strip():
            raise TranslationProviderError(f"{API_PROVIDERS[provider]['label']} 未返回有效译文。")
        return value

    def _translate_google(self, text, target_lang, config):
        url = (
            'https://translate.googleapis.com/translate_a/single'
            f'?client=gtx&sl=zh-CN&tl={quote(self._provider_language_code("google", target_lang))}'
            f'&dt=t&q={quote(text)}'
        )
        payload = self._request_json('GET', url)
        if not isinstance(payload, list) or not payload or not isinstance(payload[0], list):
            raise TranslationProviderError("Google GTX 返回格式异常。")
        translated = ''.join(
            item[0] for item in payload[0]
            if isinstance(item, list) and item and isinstance(item[0], str)
        )
        return self._require_translation(translated, 'google')

    def _translate_deepl_batch_once(self, provider, texts, target_lang, config):
        self._require_provider_config(provider, config)
        endpoint = config['endpoint'].rstrip('/')
        url = endpoint if endpoint.endswith('/translate') else f'{endpoint}/translate'
        payload = self._request_json(
            'POST',
            url,
            json_data={
                'text': texts,
                'source_lang': 'ZH',
                'target_lang': self._provider_language_code(provider, target_lang),
            },
            headers={
                'Authorization': f"DeepL-Auth-Key {config['auth_key']}",
                'Content-Type': 'application/json',
            }
        )
        if not isinstance(payload, dict):
            raise TranslationProviderError("DeepL 返回格式异常。")
        translations = payload.get('translations')
        if not isinstance(translations, list):
            raise TranslationProviderError("DeepL 未返回 translations。")
        if not all(isinstance(item, dict) and isinstance(item.get('text'), str) for item in translations):
            raise TranslationProviderError("DeepL translations 包含无效译文。")
        return [item['text'] for item in translations]

    def _translate_deepl_text_batch(self, provider, texts, target_lang, config_override=None):
        return self._translate_batch_with_provider(
            provider,
            texts,
            target_lang,
            lambda batch, target, settings: self._translate_deepl_batch_once(
                provider,
                batch,
                target,
                settings
            ),
            config_override
        )

    def _translate_deepl(self, provider, text, target_lang, config):
        translated = self._translate_deepl_batch_once(provider, [text], target_lang, config)[0]
        return self._require_translation(translated, provider)

    def _youdao_sign_input(self, text):
        if len(text) <= 20:
            return text
        return f'{text[:10]}{len(text)}{text[-10:]}'

    def _translate_youdao(self, text, target_lang, config):
        self._require_provider_config('youdao', config)
        salt = uuid.uuid4().hex
        curtime = str(int(time.time()))
        sign_text = (
            config['app_key'] + self._youdao_sign_input(text) + salt + curtime + config['app_secret']
        )
        sign = hashlib.sha256(sign_text.encode('utf-8')).hexdigest()
        payload = self._request_json(
            'POST',
            'https://openapi.youdao.com/api',
            data={
                'q': text,
                'from': 'zh-CHS',
                'to': self._provider_language_code('youdao', target_lang),
                'appKey': config['app_key'],
                'salt': salt,
                'sign': sign,
                'signType': 'v3',
                'curtime': curtime,
            }
        )
        if not isinstance(payload, dict):
            raise TranslationProviderError("有道智云返回格式异常。")
        if str(payload.get('errorCode', '0')) != '0':
            raise TranslationProviderError(
                f"有道智云错误[{payload.get('errorCode')}]: {payload.get('errorMessage', '未知错误')}。"
            )
        translations = payload.get('translation')
        if isinstance(translations, list):
            return self._require_translation(''.join(str(value) for value in translations), 'youdao')
        return self._require_translation(translations, 'youdao')

    def _translate_niutrans(self, text, target_lang, config):
        self._require_provider_config('niutrans', config)
        payload = self._request_json(
            'POST',
            config['endpoint'],
            data={
                'from': 'zh',
                'to': self._provider_language_code('niutrans', target_lang),
                'apikey': config['api_key'],
                'src_text': text,
            }
        )
        if not isinstance(payload, dict):
            raise TranslationProviderError("小牛翻译返回格式异常。")
        if payload.get('error_code') or payload.get('errorCode'):
            code = payload.get('error_code', payload.get('errorCode'))
            message = payload.get('error_msg', payload.get('errorMsg', '未知错误'))
            raise TranslationProviderError(f"小牛翻译错误[{code}]: {message}。")
        return self._require_translation(payload.get('tgt_text'), 'niutrans')

    @staticmethod
    def _hmac_sha256(key, message):
        if isinstance(key, str):
            key = key.encode('utf-8')
        return hmac.new(key, message.encode('utf-8'), hashlib.sha256).digest()

    def _request_tencent_tmt(self, action, request_data, config):
        endpoint = config['endpoint'].rstrip('/')
        parsed = urlparse(endpoint)
        if not parsed.scheme or not parsed.netloc:
            raise TranslationProviderError("腾讯云服务地址必须是完整 HTTPS URL。")

        service = 'tmt'
        version = '2018-03-21'
        timestamp = int(time.time())
        date = datetime.fromtimestamp(timestamp, timezone.utc).strftime('%Y-%m-%d')
        payload_text = json.dumps(
            request_data,
            ensure_ascii=False,
            separators=(',', ':')
        )
        payload_hash = hashlib.sha256(payload_text.encode('utf-8')).hexdigest()
        canonical_uri = parsed.path or '/'
        canonical_headers = f'content-type:application/json; charset=utf-8\nhost:{parsed.netloc}\n'
        signed_headers = 'content-type;host'
        canonical_request = '\n'.join([
            'POST', canonical_uri, '', canonical_headers, signed_headers, payload_hash
        ])
        credential_scope = f'{date}/{service}/tc3_request'
        string_to_sign = '\n'.join([
            'TC3-HMAC-SHA256',
            str(timestamp),
            credential_scope,
            hashlib.sha256(canonical_request.encode('utf-8')).hexdigest(),
        ])
        secret_date = self._hmac_sha256(f"TC3{config['secret_key']}", date)
        secret_service = self._hmac_sha256(secret_date, service)
        secret_signing = self._hmac_sha256(secret_service, 'tc3_request')
        signature = hmac.new(
            secret_signing,
            string_to_sign.encode('utf-8'),
            hashlib.sha256
        ).hexdigest()
        authorization = (
            f"TC3-HMAC-SHA256 Credential={config['secret_id']}/{credential_scope}, "
            f"SignedHeaders={signed_headers}, Signature={signature}"
        )
        payload = self._request_json(
            'POST',
            endpoint,
            data=payload_text.encode('utf-8'),
            headers={
                'Authorization': authorization,
                'Content-Type': 'application/json; charset=utf-8',
                'Host': parsed.netloc,
                'X-TC-Action': action,
                'X-TC-Version': version,
                'X-TC-Region': config['region'],
                'X-TC-Timestamp': str(timestamp),
            }
        )
        if not isinstance(payload, dict) or not isinstance(payload.get('Response'), dict):
            raise TranslationProviderError("腾讯云 TMT 返回格式异常。")
        response = payload['Response']
        if isinstance(response.get('Error'), dict):
            error = response['Error']
            error_code = error.get('Code', 'unknown')
            raise TranslationProviderError(
                f"腾讯云 TMT 错误[{error_code}]: {error.get('Message', '未知错误')}。",
                retryable=error_code == 'RequestLimitExceeded'
            )
        return response

    def _translate_tencent(self, text, target_lang, config):
        self._require_provider_config('tencent', config)
        response = self._request_tencent_tmt(
            'TextTranslate',
            {
                'SourceText': text,
                'Source': 'zh',
                'Target': self._provider_language_code('tencent', target_lang),
                'ProjectId': 0,
            },
            config
        )
        return self._require_translation(response.get('TargetText'), 'tencent')

    def _translation_cache_key(self, provider, target_lang, text):
        return f'{provider}\0{target_lang}\0{text.strip()}'

    def _batch_text_error(self, provider, text):
        limits = API_BATCH_LIMITS[provider]
        max_text_length = limits.get('max_text_length')
        if max_text_length and len(text) > max_text_length:
            return f"{API_PROVIDERS[provider]['label']} 单条文本不能超过 {max_text_length} 个字符。"
        max_total_bytes = limits.get('max_total_bytes')
        if max_total_bytes and self._batch_text_bytes(provider, text) > max_total_bytes:
            return f"{API_PROVIDERS[provider]['label']} 单条文本超过批量请求字节限制。"
        max_total_characters = limits.get('max_total_characters')
        if max_total_characters and len(text) > max_total_characters:
            return f"{API_PROVIDERS[provider]['label']} 单条文本超过批量请求字符限制。"
        return ''

    def _batch_text_bytes(self, provider, text):
        if provider in ('deepl_free', 'deepl_pro'):
            # DeepL enforces the JSON request-body size, not only source UTF-8 bytes.
            return len(json.dumps(text, ensure_ascii=True).encode('utf-8')) + 1
        return len(text.encode('utf-8'))

    def _partition_provider_texts(self, provider, texts):
        limits = API_BATCH_LIMITS[provider]
        batches = []
        rejected = []
        current_batch = []
        current_characters = 0
        current_bytes = 0
        for text in texts:
            error = self._batch_text_error(provider, text)
            if error:
                rejected.append((text, error))
                continue

            text_characters = len(text)
            text_bytes = self._batch_text_bytes(provider, text)
            max_texts = limits['max_texts']
            max_total_characters = limits.get('max_total_characters')
            max_total_bytes = limits.get('max_total_bytes')
            exceeds_batch_limit = (
                len(current_batch) == max_texts
                or (max_total_characters and current_batch and current_characters + text_characters > max_total_characters)
                or (max_total_bytes and current_batch and current_bytes + text_bytes > max_total_bytes)
            )
            if exceeds_batch_limit:
                batches.append(current_batch)
                current_batch = []
                current_characters = 0
                current_bytes = 0
            current_batch.append(text)
            current_characters += text_characters
            current_bytes += text_bytes
        if current_batch:
            batches.append(current_batch)
        return batches, rejected

    def _partition_tencent_texts(self, texts):
        """Backward-compatible wrapper used by existing TMT callers."""
        batches, rejected = self._partition_provider_texts('tencent', texts)
        return batches, [text for text, _ in rejected]

    def _translate_batch_with_provider(
        self,
        provider,
        texts,
        target_lang,
        request_batch,
        config_override=None
    ):
        if not texts:
            return []
        if provider not in API_BATCH_LIMITS:
            raise TranslationProviderError(f"{API_PROVIDERS[provider]['label']} 不支持批量文本翻译。")
        if len(texts) > API_BATCH_LIMITS[provider]['max_texts']:
            raise TranslationProviderError(
                f"{API_PROVIDERS[provider]['label']} 单次批量翻译最多支持 "
                f"{API_BATCH_LIMITS[provider]['max_texts']} 条文本。"
            )

        config = self._provider_config(provider, config_override)
        self._require_provider_config(provider, config)
        results = [None] * len(texts)
        pending = []
        for index, text in enumerate(texts):
            error = self._batch_text_error(provider, text)
            if error:
                results[index] = TranslationResult(
                    False,
                    text,
                    error
                )
                continue
            cache_key = self._translation_cache_key(provider, target_lang, text)
            if cache_key in self.translation_cache:
                results[index] = TranslationResult(True, self.translation_cache[cache_key])
                continue
            pending.append((index, text, cache_key))

        if not pending:
            return results

        last_error = ''
        for attempt in range(3):
            try:
                self._wait_for_provider_slot(provider)
                translated_texts = request_batch(
                    [text for _, text, _ in pending],
                    target_lang,
                    config
                )
                if not isinstance(translated_texts, list) or len(translated_texts) != len(pending):
                    raise TranslationProviderError(
                        f"{API_PROVIDERS[provider]['label']} 批量响应与请求文本数量不一致。"
                    )
                for (index, _, cache_key), translated_text in zip(pending, translated_texts):
                    translated = self._require_translation(translated_text, provider)
                    self.translation_cache[cache_key] = translated
                    results[index] = TranslationResult(True, translated)
                return results
            except TranslationProviderError as exc:
                last_error = str(exc)
                if not exc.retryable or attempt == 2:
                    break
                time.sleep((2 ** attempt) + random.uniform(0, 0.5))

        self.log(
            f"{API_PROVIDERS[provider]['label']} 批量翻译失败: "
            f"{len(pending)} 条 -> {target_lang}; 原因: {last_error}"
        )
        for index, text, _ in pending:
            results[index] = TranslationResult(False, text, last_error)
        return results

    def _translate_tencent_batch_once(self, texts, target_lang, config):
        response = self._request_tencent_tmt(
            'TextTranslateBatch',
            {
                'SourceTextList': texts,
                'Source': 'zh',
                'Target': self._provider_language_code('tencent', target_lang),
                'ProjectId': 0,
            },
            config
        )
        translated_texts = response.get('TargetTextList')
        if not isinstance(translated_texts, list):
            raise TranslationProviderError("腾讯云 TMT 未返回 TargetTextList。")
        return translated_texts

    def _translate_tencent_text_batch(self, texts, target_lang, config_override=None):
        return self._translate_batch_with_provider(
            'tencent',
            texts,
            target_lang,
            self._translate_tencent_batch_once,
            config_override
        )

    def _build_volcengine_signed_request(self, config, body, action='TranslateText', timestamp=None):
        """Build the exact request bytes and Signature V4 headers sent to Volcengine."""
        endpoint = config['endpoint'].strip().rstrip('/')
        parsed = urlparse(endpoint)
        if parsed.scheme.lower() != 'https' or not parsed.netloc:
            raise TranslationProviderError("火山引擎服务地址必须是完整 HTTPS URL。")
        if parsed.query or parsed.fragment:
            raise TranslationProviderError("火山引擎服务地址不能包含 query 或 fragment。")

        access_key = config['access_key'].strip()
        secret_key = config['secret_key'].strip()
        region = VOLCENGINE_DEFAULT_REGION

        if isinstance(body, str):
            body = body.encode('utf-8')
        request_time = timestamp or datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
        date = request_time[:8]
        canonical_uri = quote(parsed.path or '/', safe='/-_.~')
        query_parameters = {
            'Action': action,
            'Version': VOLCENGINE_API_VERSION,
        }
        canonical_query = '&'.join(
            f'{quote(key, safe="-_.~")}={quote(value, safe="-_.~")}'
            for key, value in sorted(query_parameters.items())
        )
        payload_hash = hashlib.sha256(body).hexdigest()
        host = parsed.netloc
        canonical_headers_map = {
            'content-type': 'application/json',
            'host': host,
            'x-content-sha256': payload_hash,
            'x-date': request_time,
        }
        canonical_headers = ''.join(
            f'{key}:{canonical_headers_map[key]}\n'
            for key in sorted(canonical_headers_map)
        )
        signed_headers = ';'.join(sorted(canonical_headers_map))
        canonical_request = '\n'.join([
            'POST',
            canonical_uri,
            canonical_query,
            canonical_headers,
            signed_headers,
            payload_hash,
        ])
        credential_scope = f'{date}/{region}/{VOLCENGINE_SERVICE}/request'
        string_to_sign = '\n'.join([
            'HMAC-SHA256',
            request_time,
            credential_scope,
            hashlib.sha256(canonical_request.encode('utf-8')).hexdigest(),
        ])

        # Volcengine's official SignerV4 derives kDate directly from SK.
        date_key = self._hmac_sha256(secret_key, date)
        region_key = self._hmac_sha256(date_key, region)
        service_key = self._hmac_sha256(region_key, VOLCENGINE_SERVICE)
        signing_key = self._hmac_sha256(service_key, 'request')
        signature = hmac.new(
            signing_key,
            string_to_sign.encode('utf-8'),
            hashlib.sha256
        ).hexdigest()
        authorization = (
            f'HMAC-SHA256 Credential={access_key}/{credential_scope}, '
            f'SignedHeaders={signed_headers}, Signature={signature}'
        )
        url = f'{parsed.scheme}://{parsed.netloc}{canonical_uri}?{canonical_query}'
        headers = {
            'Authorization': authorization,
            'Content-Type': 'application/json',
            'Host': host,
            'X-Content-Sha256': payload_hash,
            'X-Date': request_time,
        }
        return url, headers, body

    def _translate_volcengine_batch_once(self, texts, target_lang, config):
        self._require_provider_config('volcengine', config)
        payload_text = json.dumps(
            {
                'SourceLanguage': 'zh',
                'TargetLanguage': self._provider_language_code('volcengine', target_lang),
                'TextList': texts,
            },
            ensure_ascii=False,
            separators=(',', ':')
        )
        url, headers, body = self._build_volcengine_signed_request(
            config,
            payload_text,
            action='TranslateText'
        )
        payload = self._request_json(
            'POST',
            url,
            data=body,
            headers=headers
        )
        if not isinstance(payload, dict):
            raise TranslationProviderError("火山引擎返回格式异常。")
        metadata = payload.get('ResponseMetadata', {})
        if isinstance(metadata, dict) and isinstance(metadata.get('Error'), dict):
            error = metadata['Error']
            raise TranslationProviderError(
                f"火山引擎错误[{error.get('Code', 'unknown')}]: {error.get('Message', '未知错误')}。"
            )
        translations = payload.get('TranslationList')
        if not isinstance(translations, list) or len(translations) != len(texts):
            raise TranslationProviderError("火山引擎未返回 TranslationList。")
        if not all(isinstance(item, dict) and isinstance(item.get('Translation'), str) for item in translations):
            raise TranslationProviderError("火山引擎 TranslationList 包含无效译文。")
        return [item['Translation'] for item in translations]

    def _translate_volcengine_text_batch(self, texts, target_lang, config_override=None):
        return self._translate_batch_with_provider(
            'volcengine',
            texts,
            target_lang,
            self._translate_volcengine_batch_once,
            config_override
        )

    def _translate_volcengine(self, text, target_lang, config):
        translated = self._translate_volcengine_batch_once([text], target_lang, config)[0]
        return self._require_translation(translated, 'volcengine')

    def _translate_aliyun(self, text, target_lang, config):
        self._require_provider_config('aliyun', config)
        endpoint = config['endpoint'].rstrip('/')
        parsed = urlparse(endpoint)
        if not parsed.scheme or not parsed.netloc:
            raise TranslationProviderError("阿里云服务地址必须是完整 HTTPS URL。")

        payload_text = json.dumps(
            {
                'FormatType': 'text',
                'SourceLanguage': 'zh',
                'TargetLanguage': self._provider_language_code('aliyun', target_lang),
                'SourceText': text,
                'Scene': config['scene'],
            },
            ensure_ascii=False,
            separators=(',', ':')
        )
        content_md5 = base64.b64encode(hashlib.md5(payload_text.encode('utf-8')).digest()).decode('ascii')
        request_date = formatdate(usegmt=True)
        nonce = uuid.uuid4().hex
        acs_headers = {
            'x-acs-signature-method': 'HMAC-SHA1',
            'x-acs-signature-nonce': nonce,
            'x-acs-version': '2019-01-02',
        }
        canonical_headers = ''.join(
            f'{key}:{value}\n' for key, value in sorted(acs_headers.items())
        )
        canonical_resource = parsed.path or '/'
        string_to_sign = '\n'.join([
            'POST',
            'application/json',
            content_md5,
            'application/json;charset=utf-8',
            request_date,
            canonical_headers + canonical_resource,
        ])
        signature = base64.b64encode(
            hmac.new(
                config['access_key_secret'].encode('utf-8'),
                string_to_sign.encode('utf-8'),
                hashlib.sha1
            ).digest()
        ).decode('ascii')
        headers = {
            'Accept': 'application/json',
            'Content-Type': 'application/json;charset=utf-8',
            'Content-MD5': content_md5,
            'Date': request_date,
            'Host': parsed.netloc,
            'Authorization': f"acs {config['access_key_id']}:{signature}",
        }
        headers.update({key.title(): value for key, value in acs_headers.items()})
        payload = self._request_json('POST', endpoint, data=payload_text.encode('utf-8'), headers=headers)
        if not isinstance(payload, dict):
            raise TranslationProviderError("阿里云返回格式异常。")
        if payload.get('errorCode'):
            raise TranslationProviderError(
                f"阿里云错误[{payload.get('errorCode')}]: {payload.get('errorMsg', '未知错误')}。"
            )
        data = payload.get('Data', payload)
        if not isinstance(data, dict):
            raise TranslationProviderError("阿里云未返回译文数据。")
        return self._require_translation(data.get('Translated'), 'aliyun')

    @staticmethod
    def _aliyun_percent_encode(value):
        return quote(str(value), safe='~')

    def _translate_aliyun_batch_once(self, texts, target_lang, config):
        self._require_provider_config('aliyun', config)
        endpoint = config['batch_endpoint'].rstrip('/')
        parsed = urlparse(endpoint)
        if not parsed.scheme or not parsed.netloc:
            raise TranslationProviderError("阿里云批量服务地址必须是完整 HTTPS URL。")

        source_ids = [str(index) for index in range(len(texts))]
        source_text = json.dumps(
            dict(zip(source_ids, texts)),
            ensure_ascii=False,
            separators=(',', ':')
        )
        parameters = {
            'Action': 'GetBatchTranslate',
            'Version': '2018-10-12',
            'Format': 'JSON',
            'AccessKeyId': config['access_key_id'],
            'SignatureMethod': 'HMAC-SHA1',
            'Timestamp': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
            'SignatureVersion': '1.0',
            'SignatureNonce': uuid.uuid4().hex,
            'RegionId': config['region_id'],
            'FormatType': 'text',
            'SourceLanguage': 'zh',
            'TargetLanguage': self._provider_language_code('aliyun', target_lang),
            'Scene': config['scene'],
            'ApiType': config['batch_api_type'],
            'SourceText': source_text,
        }
        canonicalized_query = '&'.join(
            f'{self._aliyun_percent_encode(key)}={self._aliyun_percent_encode(value)}'
            for key, value in sorted(parameters.items())
        )
        string_to_sign = f'POST&%2F&{self._aliyun_percent_encode(canonicalized_query)}'
        signature = base64.b64encode(
            hmac.new(
                f"{config['access_key_secret']}&".encode('utf-8'),
                string_to_sign.encode('utf-8'),
                hashlib.sha1
            ).digest()
        ).decode('ascii')
        parameters['Signature'] = signature
        payload = self._request_json('POST', endpoint, data=parameters)
        if not isinstance(payload, dict):
            raise TranslationProviderError("阿里云批量翻译返回格式异常。")
        response_code = payload.get('Code')
        if response_code not in (None, 200, '200'):
            raise TranslationProviderError(
                f"阿里云批量翻译错误[{response_code}]: {payload.get('Message', '未知错误')}。"
            )
        translated_list = payload.get('TranslatedList')
        if not isinstance(translated_list, list):
            raise TranslationProviderError("阿里云批量翻译未返回 TranslatedList。")

        translated_by_id = {}
        for item in translated_list:
            if not isinstance(item, dict):
                raise TranslationProviderError("阿里云批量翻译包含无效结果。")
            source_id = str(item.get('index', item.get('Index', '')))
            translated = item.get('translated', item.get('Translated'))
            if source_id not in source_ids or source_id in translated_by_id:
                raise TranslationProviderError("阿里云批量翻译返回了无效或重复索引。")
            translated_by_id[source_id] = self._require_translation(translated, 'aliyun')
        if set(translated_by_id) != set(source_ids):
            raise TranslationProviderError("阿里云批量翻译结果不完整。")
        return [translated_by_id[source_id] for source_id in source_ids]

    def _translate_aliyun_text_batch(self, texts, target_lang, config_override=None):
        return self._translate_batch_with_provider(
            'aliyun',
            texts,
            target_lang,
            self._translate_aliyun_batch_once,
            config_override
        )

    def _translate_provider_text_batch(self, provider, texts, target_lang, config_override=None):
        if provider == 'tencent':
            return self._translate_tencent_text_batch(texts, target_lang, config_override)
        if provider in ('deepl_free', 'deepl_pro'):
            return self._translate_deepl_text_batch(provider, texts, target_lang, config_override)
        if provider == 'volcengine':
            return self._translate_volcengine_text_batch(texts, target_lang, config_override)
        if provider == 'aliyun':
            return self._translate_aliyun_text_batch(texts, target_lang, config_override)
        raise TranslationProviderError(f"{API_PROVIDERS[provider]['label']} 不支持批量文本翻译。")

    def _translate_provider_once(self, provider, text, target_lang, config):
        if provider == 'google':
            return self._translate_google(text, target_lang, config)
        if provider in ('deepl_free', 'deepl_pro'):
            return self._translate_deepl(provider, text, target_lang, config)
        if provider == 'baidu':
            translated = self.baidu_translate(text, target_lang, config)
            if translated is None:
                raise TranslationProviderError("百度翻译未返回有效译文。")
            return self._require_translation(translated, 'baidu')
        if provider == 'youdao':
            return self._translate_youdao(text, target_lang, config)
        if provider == 'niutrans':
            return self._translate_niutrans(text, target_lang, config)
        if provider == 'tencent':
            return self._translate_tencent(text, target_lang, config)
        if provider == 'volcengine':
            return self._translate_volcengine(text, target_lang, config)
        if provider == 'aliyun':
            return self._translate_aliyun(text, target_lang, config)
        raise TranslationProviderError(f"不支持的翻译服务: {provider}。")

    def _translate_with_provider(self, provider, text, target_lang, config_override=None, use_cache=True):
        if not text.strip():
            return TranslationResult(True, text)
        if provider not in API_PROVIDERS:
            return TranslationResult(False, text, f"不支持的翻译服务: {provider}。")

        config = self._provider_config(provider, config_override)
        cache_key = self._translation_cache_key(provider, target_lang, text)
        if use_cache and cache_key in self.translation_cache:
            return TranslationResult(True, self.translation_cache[cache_key])

        last_error = ''
        for attempt in range(3):
            try:
                self._wait_for_provider_slot(provider)
                translated = self._translate_provider_once(provider, text, target_lang, config)
                if use_cache:
                    self.translation_cache[cache_key] = translated
                return TranslationResult(True, translated)
            except TranslationProviderError as exc:
                last_error = str(exc)
                if not exc.retryable or attempt == 2:
                    break
                time.sleep((2 ** attempt) + random.uniform(0, 0.5))

        preview = text.replace('\n', ' ')[:80]
        self.log(
            f"{API_PROVIDERS[provider]['label']} 翻译失败: {preview} -> {target_lang}; 原因: {last_error}"
        )
        return TranslationResult(False, text, last_error)

    def translate_text(self, text, target_lang):
        """Translate one text value through the currently selected provider."""
        return self._translate_with_provider(self.api_type, text, target_lang)
    
    def translate_text_multi_lang(self, text, target_langs):
        """
        批量翻译：相同文本翻译成多种目标语言。
        返回的每项包含成功状态，失败结果不会被缓存为原文。
        """
        results = {}
        for lang in target_langs:
            results[lang] = self.translate_text(text, lang)
        return results
    
    def start_translation(self):
        """
        开始翻译流程
        1. 检查输入
        2. 提取中文文本
        3. 执行翻译
        4. 更新进度
        """
        # 检查输入
        if not self.source_files:
            msg = "请先添加文件或文件夹，或点击DiskC工作流"
            if self.headless:
                self.log(f"错误: {msg}")
                if self.logger:
                    self.logger.error(msg)
                self.translation_failures = [('', '', msg)]
                self.translation_complete = True
                return
            else:
                messagebox.showwarning("警告", msg)
                return

        # 获取选中的语言（GUI模式）或保持已有设置（headless可在外部设置）
        if not self.headless:
            self.selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]

        if not self.selected_langs:
            msg = "请至少选择一种目标语言"
            if self.headless:
                self.log(f"错误: {msg}")
                if self.logger:
                    self.logger.error(msg)
                self.translation_failures = [('', '', msg)]
                self.translation_complete = True
                return
            else:
                messagebox.showwarning("警告", msg)
                return

        if self.api_type not in API_PROVIDERS:
            msg = f"不支持的翻译服务: {self.api_type}"
            if self.headless:
                self.log(f"错误: {msg}")
                self.translation_failures = [('', '', msg)]
                self.translation_complete = True
            else:
                messagebox.showwarning("API配置", msg)
            return

        missing = self._missing_provider_fields(self.api_type)
        if missing:
            msg = f"{API_PROVIDERS[self.api_type]['label']} 缺少配置: {', '.join(missing)}"
            if self.headless:
                self.log(f"错误: {msg}")
                self.translation_failures = [('', '', msg)]
                self.translation_complete = True
            else:
                messagebox.showwarning("API配置", msg)
            return
        
        # 初始化状态
        self.start_btn.config(state='disabled')
        self.confirm_btn.config(state='disabled')
        self.translation_complete = False
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.translation_failures = []
        self.log("开始提取中文文本...")
        
        # 提取中文文本
        texts = self.extract_chinese_texts()
        self.log(f"共需处理 {len(texts)} 条中文文本")
        task_texts = [
            text
            for text in {
                self._normalize_text(source_text)
                for source_text in texts
            }
            if text in self.translation_table
        ]
        self.log(
            f"本次任务限定为当前源文件的 {len(task_texts)} 条文本，"
            "不会翻译翻译记忆库中的历史条目"
        )
        
        # 计算需要翻译的数量
        total_needed = 0
        for text in task_texts:
            langs = self.translation_table[text]
            for lang in self.selected_langs:
                if lang not in langs:
                    langs[lang] = ''
                if not langs[lang] or langs[lang] == text:
                    total_needed += 1
        
        # 如果没有需要翻译的内容
        if total_needed == 0:
            self.log("所有翻译已完成")
            self.translation_complete = True
            self.confirm_btn.config(state='normal')
            self.start_btn.config(state='normal')
            return
        
        self.log(f"需要翻译: {total_needed} 条")
        
        # 创建翻译线程
        def translate_thread():
            from concurrent.futures import ThreadPoolExecutor, as_completed

            # 按文本分组：相同文本的多语言翻译请求合并
            # 结构: {text: [lang1, lang2, ...]}
            text_to_langs = {}
            for text in task_texts:
                for lang in self.selected_langs:
                    if lang not in self.translation_table[text]:
                        self.translation_table[text][lang] = ''
                    if not self.translation_table[text][lang] or self.translation_table[text][lang] == text:
                        if text not in text_to_langs:
                            text_to_langs[text] = []
                        text_to_langs[text].append(lang)

            worker_count = self._provider_worker_count()
            concurrency_log = (
                f"优化后: {len(text_to_langs)} 个唯一文本，{total_needed} 个翻译任务，"
                f"并发数 {worker_count}"
            )
            requests_per_second = self._provider_requests_per_second(self.api_type)
            if requests_per_second:
                concurrency_log = f"{concurrency_log}，全局限流 {requests_per_second:g} 次/秒"
            self.log(concurrency_log)

            completed = 0
            succeeded = 0
            failed = 0
            failures = []
            last_saved_completed = 0
            lock = threading.Lock()

            def commit_results(result_items):
                nonlocal completed, succeeded, failed, last_saved_completed
                result_items = list(result_items)
                if not result_items:
                    return

                should_save = False
                with lock:
                    for text, lang, translation_result in result_items:
                        if translation_result and translation_result.success:
                            self.translation_table[text][lang] = translation_result.text
                            succeeded += 1
                        else:
                            # Leave failed cells pending so a later run can retry them.
                            self.translation_table[text][lang] = ''
                            failed += 1
                            error = translation_result.error if translation_result else "翻译服务未返回结果。"
                            failures.append((text, lang, error))
                        completed += 1

                    current_completed = completed
                    current_progress = round(completed / total_needed * 100, 1)
                    if completed - last_saved_completed >= 50:
                        last_saved_completed = completed
                        should_save = True

                if should_save:
                    self.save_translation_table()
                self._run_on_ui(
                    lambda p=current_progress, c=current_completed: self._update_progress(p, c)
                )
                self.log(f"进度: {current_progress}% ({current_completed}/{total_needed})")

            def translate_one_text(item):
                text, langs = item
                lang_codes = [LANG_MAP[lang]['code'] for lang in langs]
                results = self.translate_text_multi_lang(text, lang_codes)
                return text, langs, results

            if self.api_type in API_BATCH_LIMITS:
                batch_provider = self.api_type
                texts_by_language = {}
                for text, langs in text_to_langs.items():
                    for lang in langs:
                        texts_by_language.setdefault(lang, []).append(text)

                batch_tasks = []
                rejected_count = 0
                for lang, texts_for_language in texts_by_language.items():
                    batches, rejected_texts = self._partition_provider_texts(
                        batch_provider,
                        texts_for_language
                    )
                    batch_tasks.extend((lang, batch) for batch in batches)
                    rejected_count += len(rejected_texts)
                    commit_results(
                        (
                            text,
                            lang,
                            TranslationResult(False, text, error)
                        )
                        for text, error in rejected_texts
                    )

                limits = API_BATCH_LIMITS[batch_provider]
                self.log(
                    f"{API_PROVIDERS[batch_provider]['label']} 批量模式："
                    f"{len(batch_tasks)} 个请求批次，每批最多 {limits['max_texts']} 条文本"
                )
                if rejected_count:
                    self.log(
                        f"{API_PROVIDERS[batch_provider]['label']} 跳过 {rejected_count} 条"
                        "超过批量限制的文本。"
                    )

                def translate_provider_batch(task):
                    lang, batch = task
                    target_lang = LANG_MAP[lang]['code']
                    results = self._translate_provider_text_batch(
                        batch_provider,
                        batch,
                        target_lang
                    )
                    return lang, batch, results

                with ThreadPoolExecutor(max_workers=worker_count) as executor:
                    futures = [
                        executor.submit(translate_provider_batch, task)
                        for task in batch_tasks
                    ]
                    for future in as_completed(futures):
                        lang, batch, results = future.result()
                        commit_results(zip(batch, [lang] * len(batch), results))
            else:
                with ThreadPoolExecutor(max_workers=worker_count) as executor:
                    futures = {
                        executor.submit(translate_one_text, (text, langs)): text
                        for text, langs in text_to_langs.items()
                    }
                    for future in as_completed(futures):
                        text, langs, results = future.result()
                        commit_results(
                            (
                                text,
                                lang,
                                results.get(LANG_MAP[lang]['code'])
                            )
                            for lang in langs
                        )

            self.log("执行智能缩写...")
            self.abbreviate_translations(task_texts)
            self.save_translation_table()
            self._run_on_ui(lambda: self._update_progress(100, total_needed))
            self.translation_failures = failures
            if failed:
                self.log(f"翻译结束：成功 {succeeded} 条，失败 {failed} 条。失败项已保留为空，可切换服务后重试。")
            else:
                self.log(f"翻译完成！成功 {succeeded} 条。")
            self.translation_complete = True
            self._run_on_ui(lambda: self.confirm_btn.config(state='normal'))
            self._run_on_ui(lambda: self.start_btn.config(state='normal'))
        
        # 启动翻译线程
        threading.Thread(target=translate_thread, daemon=True).start()
    
    def abbreviate_translations(self, texts=None):
        import re

        max_lengths = {
            'default': 40,
            'JPN': 22, 'KOR': 22,
            'GER': 35, 'DES': 35, 'DEA': 35, 'DEL': 35, 'DEC': 35,
            'FRA': 32, 'FRB': 32, 'FRC': 32, 'FRS': 32, 'FRL': 32, 'FRM': 32,
            'ITA': 30, 'ITS': 30,
            'ESP': 28, 'ESM': 28, 'ESG': 28, 'ESC': 28, 'ESA': 28, 'ESD': 28,
            'ESV': 28, 'ESO': 28, 'ESR': 28, 'ESS': 28, 'ESF': 28, 'ESL': 28,
            'ESY': 28, 'ESB': 28, 'ESE': 28, 'ESH': 28, 'ESN': 28, 'ESU': 28, 'ESP2': 28,
            'PTG': 30, 'PTB': 30,
            'RUS': 25,
            'ARA': 40, 'ARL': 40, 'ARG': 40, 'ARM': 40, 'ART': 40,
            'ARO': 40, 'ARY': 40, 'ARS': 40, 'ARJ': 40, 'ARB': 40,
            'ARK': 40, 'ARU': 40, 'ARH': 40, 'ARQ': 40,
        }

        CN_FUNCTION_WORDS = set(
            '的 是 在 和 与 及 了 有 我 你 他 她 它 能 可 将 已 被 从 到 对 为 于 以 此 其 该 这 那 些 等 或 但 如 若 则 因 所 并 且 而 又 也 更 最 很 太 较 稍 略 仅 只 均 皆 各 每 任 某 另 再 还 仍 尚 须 应 需 要 会 得 让 把 比 按 照 根 据 由 向 往 朝 沿 顺 随 跟 同 非 无 不 没 未 否 勿 别 莫 休 罢 着 过 起 来 去 出 入 上 下 左 右 前 后 内 外 中 间 里 边 旁 侧 底 顶 头 尾 首 末 始 终 全 整 部 分 段 节 项 条 类 种 型 式 样 法 方 式 途 径 路 线 面 体 点 位 置 处 所 场 地 域 区 范 围 界 限 度 量 数 值 参 据 信 息 号 码 标 志 记 符 名 称 编 序 代 码 键 值 组 列 表 格 框 窗 页 屏 幕 板 卡 盘 钮 键 开 关 启 停 复 重 设 置 配 选 确 取 消 删 除 增 加 修 改 更 新 换 替 转 变 化 显 示 隐 藏 查 找 搜 索 读 写 存 取 载 卸 装 连 接 断 开 通 发 送 收 受 传 输 播 放 录 记 打 印 预 览 扫 描 检 测 试 验 证 核 校 正 调 整 优 化 改 善 提 升 降 低 减 多 扩 大 缩 拉 伸 压 移 动 拖 拽 滚 翻 页 跳 转 返 回 退 进'
        )

        EN_STOP_WORDS = set(
            'a an the and or but if is are was were be been being have has had do does did will would shall should can could may might must need want like love hate think know believe see hear feel make get give take come go use find tell ask say speak talk walk run work play start stop begin end open close save load create delete add remove set get put show hide enable disable select choose click press enter exit cancel ok yes no true false on off in at to from by for of with about above after before between under over through into onto upon within without during since until while as than up down left right back front near far all any each every some many few more most less much such other another same different new old first last next previous current default custom user system file folder directory window page menu button label text input output setting option mode state status type kind form format size color style value name number id index key code data info message error warning success fail pass result total count sum average min max low high auto manual public private local global remote main sub help tip note log view edit copy cut paste paste undo redo clear reset refresh reload search sort filter group merge split join connect disconnect attach detach lock unlock freeze thaw zoom in out expand collapse'
        )

        JP_PARTICLES = set('の は が を に へ と で から まで も など しか も か なら けれど けど し て に な だ です ます た ない ぬ ね よ わ ぁ い う え お'.split())

        KR_PARTICLES = set('의 에서 은 는 를 을 가 과 와 도 부터 까지 조차 만큼 처럼'.split())

        def abbreviate_cn(text, limit):
            chars = list(text)
            kept = [c for c in chars if c not in CN_FUNCTION_WORDS]
            result = ''.join(kept)
            if len(result) <= limit:
                return result
            head = result[:limit//2]
            tail = result[-(limit//2-2):] if len(result) > limit else ''
            return f"{head}…{tail}" if tail else f"{result[:limit-1]}…"

        def abbreviate_en(text, limit):
            words = text.split()
            if len(words) <= 2:
                return text[:limit] + ('...' if len(text) > limit else '')
            kept = []
            for i, w in enumerate(words):
                w_lower = w.lower().rstrip('.,;:!?')
                if i == 0:
                    kept.append(w)
                elif w_lower in EN_STOP_WORDS or len(w_lower) <= 1:
                    continue
                elif len(w) > 7:
                    kept.append(w[:5] + '.')
                else:
                    kept.append(w)
                if len(' '.join(kept)) >= limit - 3:
                    break
            result = ' '.join(kept)
            if len(result) <= limit:
                return result
            return f"{result[:limit//2]}…{result[-(limit//2-2):]}" if len(result) > limit else f"{result[:limit-1]}…"

        def abbreviate_jp(text, limit):
            chars = list(text)
            kept = [c for c in chars if c not in JP_PARTICLES]
            result = ''.join(kept)
            if len(result) <= limit:
                return result
            return f"{result[:limit-1]}…"

        def abbreviate_kr(text, limit):
            syllables = re.findall(r'[\uAC00-\uD7AF\u1100-\u11FF\u3130-\u318F]|[^\s]', text)
            result = ''.join(syllables)
            if len(result) <= limit:
                return result
            return f"{result[:limit-1]}…"

        def abbreviate_default(text, limit):
            words = text.split()
            if len(words) <= 2:
                return text[:limit] + ('...' if len(text) > limit else '')
            kept = [words[0]]
            for w in words[1:]:
                w_clean = w.lower().rstrip('.,;:!?')
                if w_clean in EN_STOP_WORDS or len(w_clean) <= 1:
                    continue
                elif len(w) > 6:
                    kept.append(w[:4] + '.')
                else:
                    kept.append(w)
                if len(' '.join(kept)) >= limit - 3:
                    break
            result = ' '.join(kept)
            return result if len(result) <= limit else f"{result[:limit-1]}…"

        texts_to_check = (
            list(self.translation_table)
            if texts is None
            else [text for text in texts if text in self.translation_table]
        )
        abbreviated_count = 0
        total_to_check = sum(
            1 for text in texts_to_check
            for lang in self.selected_langs
            if lang in self.translation_table[text] and self.translation_table[text][lang]
            and len(self.translation_table[text][lang]) > max_lengths.get(lang, max_lengths['default'])
            and lang not in ('CHT', 'CHS', 'ZHH', 'ZHI', 'ZHM')
        )
        checked = 0
        for text in texts_to_check:
            for lang in self.selected_langs:
                if lang not in self.translation_table[text] or not self.translation_table[text][lang]:
                    continue
                original = self.translation_table[text][lang]
                max_len = max_lengths.get(lang, max_lengths['default'])
                if len(original) <= max_len:
                    continue
                if lang in ('CHT', 'CHS', 'ZHH', 'ZHI', 'ZHM'):
                    continue
                elif lang == 'JPN':
                    new_val = abbreviate_jp(original, max_len)
                elif lang == 'KOR':
                    new_val = abbreviate_kr(original, max_len)
                elif lang in ('ENG', 'USA', 'ENA', 'ENC', 'ENZ', 'ENI', 'ENS'):
                    new_val = abbreviate_en(original, max_len)
                elif lang in ('GER', 'DES', 'DEA', 'DEL', 'DEC'):
                    new_val = abbreviate_en(original, max_len)
                elif lang in ('FRA', 'FRB', 'FRC', 'FRS', 'FRL', 'FRM'):
                    new_val = abbreviate_en(original, max_len)
                elif lang in ('ITA', 'ITS'):
                    new_val = abbreviate_en(original, max_len)
                elif lang in ('ESP', 'ESM', 'ESG', 'ESC', 'ESA', 'ESD',
                              'ESV', 'ESO', 'ESR', 'ESS', 'ESF', 'ESL',
                              'ESY', 'ESB', 'ESE', 'ESH', 'ESN', 'ESU', 'ESP2'):
                    new_val = abbreviate_en(original, max_len)
                elif lang in ('PTG', 'PTB'):
                    new_val = abbreviate_en(original, max_len)
                elif lang == 'RUS':
                    new_val = abbreviate_default(original, max_len)
                elif lang in ('ARA', 'ARL', 'ARG', 'ARM', 'ART', 'ARO',
                              'ARY', 'ARS', 'ARJ', 'ARB', 'ARK', 'ARU', 'ARH', 'ARQ'):
                    new_val = abbreviate_default(original, max_len)
                else:
                    new_val = abbreviate_default(original, max_len)
                    new_val = abbreviate_default(original, max_len)
                self.translation_table[text][lang] = new_val
                abbreviated_count += 1
                checked += 1
                if checked % 200 == 0:
                    pct = round(checked / max(total_to_check, 1) * 100, 1)
                    self.log(f"智能缩写: {pct}% ({checked}/{total_to_check})")

        self.log(f"智能缩写完成，共缩写 {abbreviated_count} 条翻译")
    
    def _update_progress(self, progress, count):
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{progress}%")
        if hasattr(self, 'progress_label') and progress > 0:
            master = self.progress_label.master
            for child in master.winfo_children():
                if isinstance(child, ttk.Label) and child.cget('text').startswith(' '):
                    child.config(text="  处理中..." if progress < 100 else "  完成!")
                    break
    
    def confirm_and_generate(self):
        """
        确认翻译并生成XML文件
        需要用户确认后才执行生成操作
        """
        if not self.translation_complete:
            messagebox.showwarning("警告", "请先完成翻译")
            return
        
        if self.translation_failures:
            prompt = (
                f"有 {len(self.translation_failures)} 条翻译失败，未翻译内容会保留原文。\n"
                "建议先切换或修复 API 后重试。仍要生成 XML 文件吗？"
            )
            result = messagebox.askyesno("翻译存在失败", prompt)
        else:
            result = messagebox.askyesno("确认生成", "翻译已完成，是否确认生成XML文件？")
        if result:
            self.generate_xml_files()

    def _build_translated_resmap_xml(self, content, lang, include_language_identity=False):
        messages = []
        for match in re.finditer(
            r'<Message\b[^>]*/>',
            content,
            flags=re.IGNORECASE | re.DOTALL
        ):
            raw_elem = match.group(0)
            content_match = (
                re.search(r'Content\s*=\s*"(.*?)"', raw_elem) or
                re.search(r"Content\s*=\s*'(.*?)'", raw_elem)
            )
            original_content = content_match.group(1) if content_match else ''
            import html
            messages.append((html.unescape(original_content), raw_elem))

        lines = [
            '<?xml version="1.0" encoding="utf-8"?>',
            '<ResMap>',
        ]
        if include_language_identity:
            lang_identity = LANG_NATIVE_NAME.get(lang, lang)
            lang_identity = lang_identity.replace('\n', '&#xA;').replace('"', '&quot;')
            lines.append(f'  <Message ID="{lang}" Content="{lang_identity}" />')

        for original_content, raw_elem in messages:
            normalized_content = self._normalize_text(original_content)
            if normalized_content in self.translation_table and lang in self.translation_table[normalized_content]:
                translated = self.translation_table[normalized_content][lang]
                content_val = translated if translated and translated != original_content else original_content
            else:
                content_val = original_content

            content_val = content_val.replace('\n', '&#xA;')
            content_val = content_val.replace('"', '&quot;')
            new_elem = re.sub(
                r'Content\s*=\s*"[^"]*"',
                f'Content="{content_val}"',
                raw_elem
            )
            lines.append(f'  {new_elem.strip()}')

        lines.append('</ResMap>')
        return '\n'.join(lines), len(messages)

    def generate_xml_files(self):
        """
        生成各语言版本的XML文件
        将翻译结果写入对应语言文件夹
        """
        if not self.source_files:
            msg = "请先选择XML源文件"
            if self.headless:
                self.log(f"错误: {msg}")
                if self.logger:
                    self.logger.error(msg)
                return
            else:
                messagebox.showwarning("警告", msg)
                return

        if not self.selected_langs:
            msg = "请至少选择一种目标语言"
            if self.headless:
                self.log(f"错误: {msg}")
                if self.logger:
                    self.logger.error(msg)
                return
            else:
                messagebox.showwarning("警告", msg)
                return

        if self.mb_backup_mode:
            self.generate_mb_backup_file()
            return
        
        self.log("开始生成XML文件...")
        
        for xml_file in self.source_files:
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()

                self.log(f"从 {os.path.basename(xml_file)} 提取 Message，准备生成翻译文件")

                file_name = os.path.basename(xml_file)
                if file_name.endswith('.xml'):
                    base_name = file_name[:-4]
                else:
                    base_name = file_name

                is_diskc_res = self.diskc_mode and xml_file == self.diskc_res_source
                is_diskc_xml = self.diskc_mode and xml_file in self.diskc_xml_map
                is_diskc_plugin = self.diskc_mode and xml_file == self.diskc_plugin_source
                is_simulator_xml = (
                    self.simulator_mode
                    and xml_file in self.simulator_xml_map
                )

                # 为每种选中的语言生成XML文件
                for lang in self.selected_langs:
                    if is_diskc_res:
                        language_dir = os.path.join(self.diskc_root, 'OpenCNC', 'Bin', 'Language')
                        os.makedirs(language_dir, exist_ok=True)
                        xml_output_path = os.path.join(language_dir, lang)
                        pack_after_write = True
                    elif is_diskc_xml:
                        rel_path = self.diskc_xml_map[xml_file]
                        string_dir = os.path.join(self.diskc_root, 'OpenCnc Shared', 'OCRes', lang, 'String')
                        os.makedirs(os.path.join(string_dir, os.path.dirname(rel_path)), exist_ok=True)
                        xml_output_path = os.path.join(string_dir, rel_path)
                        pack_after_write = False
                    elif is_diskc_plugin:
                        plugin_dir = os.path.join(self.diskc_root, 'OpenCNC', 'Bin', 'Plugin', 'Config')
                        os.makedirs(plugin_dir, exist_ok=True)
                        xml_output_path = os.path.join(plugin_dir, lang + '.xml')
                        pack_after_write = False
                    elif is_simulator_xml:
                        rel_path = self.simulator_xml_map[xml_file]
                        string_dir = os.path.join(
                            self.simulator_root,
                            'OpenCnc Shared',
                            'OCRes',
                            lang,
                            'String'
                        )
                        os.makedirs(
                            os.path.join(string_dir, os.path.dirname(rel_path)),
                            exist_ok=True
                        )
                        xml_output_path = os.path.join(string_dir, rel_path)
                        pack_after_write = False
                    else:
                        xml_output_path = os.path.join(self.output_dir, lang, 'String', base_name + '.xml')
                        os.makedirs(os.path.dirname(xml_output_path), exist_ok=True)
                        pack_after_write = self.pack_var.get()

                    rendered_xml, message_count = self._build_translated_resmap_xml(
                        content,
                        lang,
                        include_language_identity=pack_after_write
                    )
                    if message_count == 0:
                        content_preview = content.strip()[:120]
                        self.log(
                            f"警告: {os.path.basename(xml_file)} 未匹配到任何Message标签，"
                            f"文件预览: {content_preview}"
                        )

                    try:
                        with open(xml_output_path, 'w', encoding='utf-8') as f_out:
                            f_out.write(rendered_xml)
                        self.log(
                            f"生成 {lang}: {os.path.basename(xml_output_path)} "
                            f"({message_count} 条Message)"
                        )
                    except Exception as write_e:
                        self.log(f"写入文件失败: {write_e}")
                        if self.logger:
                            self.logger.exception(write_e)

                    if pack_after_write:
                        self.pack_to_res(xml_output_path)
            except Exception as e:
                self.log(f"生成XML失败: {e}")
                if self.logger:
                    self.logger.exception(e)
        
        self.log("所有XML文件已生成完成！")
        
        self._cleanup_temp_files()
        
        if getattr(self, 'simulator_mode', False):
            messagebox.showinfo(
                "完成",
                "模拟器工作流处理完成！\n"
                "XML文件已输出至: OpenCnc Shared/OCRes/{LANG}/String/"
            )
        elif self.diskc_mode:
            messagebox.showinfo("完成", f"DiskC工作流处理完成！\n"
                                f"RES文件已打包至: OpenCNC/Bin/Language/\n"
                                f"XML文件已输出至: OpenCnc Shared/OCRes/\n"
                                f"临时文件已清理")
        elif self.pack_var.get():
            messagebox.showinfo("完成", "XML文件已成功生成并打包为.res文件！\n临时文件已清理")
        else:
            messagebox.showinfo("完成", "XML文件已成功生成！\n临时文件已清理")

    @staticmethod
    def _file_crc32(file_path):
        checksum = 0
        with open(file_path, 'rb') as source_file:
            while True:
                chunk = source_file.read(1024 * 1024)
                if not chunk:
                    break
                checksum = zlib.crc32(chunk, checksum)
        return checksum & 0xffffffff

    @staticmethod
    def _mb_backup_output_path(archive_path, checksum):
        archive_dir = os.path.dirname(archive_path)
        archive_stem = os.path.splitext(os.path.basename(archive_path))[0]
        if re.fullmatch(r'.+_[0-9A-Fa-f]{8}', archive_stem):
            archive_stem = archive_stem.rsplit('_', 1)[0]
        return os.path.join(archive_dir, f'{archive_stem}_{checksum:08X}.zip')

    @staticmethod
    def _copy_zip_entry(source_archive, target_archive, source_info):
        target_info = copy.copy(source_info)
        target_info.header_offset = 0
        data = b'' if source_info.is_dir() else source_archive.read(source_info)
        target_archive.writestr(
            target_info,
            data,
            compress_type=source_info.compress_type
        )

    @staticmethod
    def _generated_zip_info(source_info, target_entry):
        target_info = copy.copy(source_info)
        target_info.filename = target_entry
        target_info.orig_filename = target_entry
        target_info.header_offset = 0
        return target_info

    def _cleanup_mb_backup_temp_files(self):
        work_dir = self.mb_backup_work_dir
        if not work_dir:
            return
        if os.path.isdir(work_dir):
            try:
                shutil.rmtree(work_dir)
                self.log(f"已删除 MB备份临时目录: {work_dir}")
            except OSError as error:
                self.log(f"删除 MB备份临时目录失败: {work_dir} -> {error}")
                return
        self.mb_backup_work_dir = ""

    def generate_mb_backup_file(self):
        if (
            not self.mb_backup_archive_path
            or not self.mb_backup_output_base_path
            or not self.mb_backup_sources
        ):
            message = "MB备份工作流未准备可生成的资源。"
            self.log(message)
            if not self.headless:
                messagebox.showerror("MB备份工作流", message)
            return None
        if not self.selected_langs:
            message = "请至少选择一种目标语言。"
            self.log(message)
            if not self.headless:
                messagebox.showwarning("MB备份工作流", message)
            return None

        target_languages = [lang for lang in self.selected_langs if lang != 'CHS']
        if len(target_languages) != len(self.selected_langs):
            self.log("MB备份工作流: 源语言 CHS 无需生成，已跳过。")
        if not target_languages:
            message = "MB备份工作流没有可生成的非 CHS 目标语言。"
            self.log(message)
            if not self.headless:
                messagebox.showwarning("MB备份工作流", message)
            return None

        temporary_output_path = None
        try:
            generated_entries = {}
            with zipfile.ZipFile(self.mb_backup_archive_path, 'r') as source_archive:
                source_infos = {
                    info.filename: info
                    for info in source_archive.infolist()
                    if not info.is_dir()
                }
                for source in self.mb_backup_sources.values():
                    source_info = source_infos.get(source.archive_entry)
                    if source_info is None:
                        raise ValueError(
                            f"原始MB备份中缺少已准备资源: {source.archive_entry}"
                        )
                    try:
                        source_content = source_archive.read(source_info).decode('utf-8')
                    except UnicodeDecodeError as error:
                        raise ValueError(
                            f"MB备份资源不是 UTF-8 XML: {source.archive_entry}"
                        ) from error

                    for lang in target_languages:
                        target_entry = source.target_entry(lang)
                        if target_entry in generated_entries:
                            raise ValueError(
                                f"MB备份目标资源路径重复: {target_entry}"
                            )
                        rendered_xml, message_count = self._build_translated_resmap_xml(
                            source_content,
                            lang
                        )
                        if message_count == 0:
                            self.log(
                                f"MB备份工作流: {source.archive_entry} 未匹配到 Message，"
                                "将保留空 ResMap 输出。"
                            )
                        generated_entries[target_entry] = (
                            source_info,
                            rendered_xml.encode('utf-8'),
                        )

                output_dir = os.path.dirname(self.mb_backup_output_base_path)
                file_descriptor, temporary_output_path = tempfile.mkstemp(
                    prefix='.mb_backup.',
                    suffix='.zip',
                    dir=output_dir
                )
                os.close(file_descriptor)
                with zipfile.ZipFile(
                    temporary_output_path,
                    'w',
                    compression=zipfile.ZIP_DEFLATED,
                    allowZip64=True
                ) as target_archive:
                    target_archive.comment = source_archive.comment
                    for source_info in source_archive.infolist():
                        if source_info.filename in generated_entries:
                            continue
                        self._copy_zip_entry(
                            source_archive,
                            target_archive,
                            source_info
                        )
                    for target_entry in sorted(generated_entries):
                        source_info, rendered_xml = generated_entries[target_entry]
                        target_archive.writestr(
                            self._generated_zip_info(source_info, target_entry),
                            rendered_xml,
                            compress_type=source_info.compress_type
                        )

            checksum = self._file_crc32(temporary_output_path)
            output_path = self._mb_backup_output_path(
                self.mb_backup_output_base_path,
                checksum
            )
            source_zip_path = os.path.normcase(
                os.path.abspath(self.mb_backup_archive_path)
            )
            if (
                self.mb_backup_input_type == 'zip'
                and os.path.normcase(os.path.abspath(output_path)) == source_zip_path
            ):
                raise ValueError("生成后的 MB备份 CRC32 未变化，已拒绝覆盖原始备份。")
            if (
                self.mb_backup_input_type == 'folder'
                and os.path.normcase(os.path.abspath(output_path)) == os.path.normcase(
                    f"{os.path.abspath(self.mb_backup_root_path)}.zip"
                )
            ):
                raise ValueError("生成后的 MB备份将覆盖同名原始ZIP，已拒绝写入。")
            os.replace(temporary_output_path, output_path)
            temporary_output_path = None
        except (OSError, RuntimeError, ValueError, zipfile.BadZipFile, zipfile.LargeZipFile) as error:
            if temporary_output_path and os.path.exists(temporary_output_path):
                try:
                    os.remove(temporary_output_path)
                except OSError as cleanup_error:
                    self.log(f"清理 MB备份临时ZIP失败: {temporary_output_path} -> {cleanup_error}")
            message = f"生成MB备份失败: {error}"
            self.log(message)
            if not self.headless:
                messagebox.showerror("MB备份工作流", message)
            return None

        self.mb_backup_output_path = output_path
        self._cleanup_mb_backup_temp_files()
        message = (
            f"MB备份工作流处理完成：已生成 {len(target_languages)} 种语言，"
            f"共写入 {len(generated_entries)} 个翻译资源。\n{output_path}"
        )
        self.log(message)
        if not self.headless:
            messagebox.showinfo("MB备份工作流", message)
        return output_path

    def _cleanup_temp_files(self):
        import shutil
        
        self.log("开始清理临时文件...")
        
        conv_dir = os.path.join(self.output_dir, '_res_converted')
        if os.path.exists(conv_dir):
            try:
                shutil.rmtree(conv_dir)
                self.log(f"已删除: {conv_dir}")
            except Exception as e:
                self.log(f"删除失败: {conv_dir} -> {e}")
        
        if self.diskc_mode and self.diskc_root:
            lang_dir = os.path.join(self.diskc_root, 'OpenCNC', 'Bin', 'Language')
            if os.path.isdir(lang_dir):
                deleted_count = 0
                for f in os.listdir(lang_dir):
                    fp = os.path.join(lang_dir, f)
                    if os.path.isfile(fp) and not os.path.splitext(fp)[1]:
                        try:
                            os.remove(fp)
                            deleted_count += 1
                        except Exception as e:
                            self.log(f"删除失败: {fp} -> {e}")
                if deleted_count > 0:
                    self.log(f"已删除 Language 目录下 {deleted_count} 个无后缀文件")
        
        self.log("临时文件清理完成")
    
    def pack_to_res(self, xml_file_path):
        """
        将XML文件打包为GZIP格式的.res文件
        """
        try:
            with open(xml_file_path, 'rb') as f_in:
                content = f_in.read()
            
            import zlib
            crc = zlib.crc32(content) & 0xffffffff
            compressed = zlib.compress(content, 9)
            
            # 构造GZIP文件头（确保FLG=0x00）
            gzip_header = b'\x1f\x8b\x08\x00'  # ID1, ID2, CM, FLG
            gzip_header += b'\x00\x00\x00\x00'  # MTIME (0)
            gzip_header += b'\x00'              # XFL (0)
            gzip_header += b'\xff'              # OS (255 = unknown)
            
            # 生成.res文件路径
            if xml_file_path.endswith('.xml'):
                res_path = xml_file_path[:-4] + '.res'
            else:
                res_path = xml_file_path + '.res'
            
            # 如果目标文件已存在，先删除以确保覆盖
            if os.path.exists(res_path):
                os.remove(res_path)
                self.log(f"覆盖已存在的文件: {res_path}")
            
            # 写入文件
            with open(res_path, 'wb') as f_out:
                f_out.write(gzip_header)
                f_out.write(compressed[2:-4])  # 移除zlib头和adler32
                # 添加GZIP尾部（CRC32和原始大小）
                f_out.write(crc.to_bytes(4, 'little'))
                f_out.write(len(content).to_bytes(4, 'little'))
            
            self.log(f"打包: {res_path}")
        except Exception as e:
            self.log(f"打包失败: {e}")
    
    def view_translation_table(self):
        """
        打开翻译表文件
        """
        table_path = os.path.join(self.table_dir, 'translation_table.json')
        if os.path.exists(table_path):
            os.startfile(table_path)
        else:
            messagebox.showinfo("提示", "翻译表文件不存在")
    
    def export_log(self):
        """
        导出日志文件
        """
        log_content = self.log_text.get('1.0', tk.END)
        log_path = os.path.join(self.output_dir, 'translation_log.txt')
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(log_content)
        messagebox.showinfo("提示", f"日志已导出到: {log_path}")
    
    def export_to_excel(self):
        """
        导出翻译表到Excel文件
        方便人工翻译和编辑
        """
        if not self.translation_table:
            messagebox.showwarning("警告", "翻译表为空，请先进行翻译")
            return
        
        # 动态获取当前勾选的语言
        current_selected_langs = [lang for lang, var in self.lang_vars.items() if var.get()]
        if not current_selected_langs:
            messagebox.showwarning("警告", "请先选择目标语言")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="导出Excel文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "翻译表"
            
            # 表头：原文 + 各语言（显示语言缩写和名称）
            headers = ["原文"]
            for lang in current_selected_langs:
                lang_name = LANG_MAP.get(lang, {}).get('name', lang)
                headers.append(f"{lang} - {lang_name}")
            ws.append(headers)
            
            # 填充数据
            for original_text, translations in self.translation_table.items():
                row = [original_text]
                for lang in current_selected_langs:
                    row.append(translations.get(lang, ''))
                ws.append(row)
            
            # 自动调整列宽
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width
            
            wb.save(file_path)
            self.log(f"翻译表已导出到Excel: {file_path}")
            messagebox.showinfo("完成", f"翻译表已成功导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出Excel失败: {str(e)}")
    
    def import_from_excel(self):
        """
        从Excel文件导入翻译结果
        支持人工编辑后重新导入
        """
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            
            # 获取表头
            headers = []
            for cell in ws[1]:
                headers.append(cell.value)
            
            if headers[0] != "原文":
                messagebox.showwarning("警告", "Excel文件格式不正确，第一列必须是'原文'")
                return
            
            # 获取语言列表（解析列标题，提取语言缩写）
            import_langs = []
            for header in headers[1:]:
                if header:
                    # 解析格式: "语言缩写 - 语言名称"
                    parts = str(header).split(' - ', 1)
                    lang_code = parts[0].strip() if parts else str(header).strip()
                    import_langs.append(lang_code)
            
            # 导入数据
            updated_count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                
                original_text = row[0]
                
                # 创建或更新翻译表条目
                if original_text not in self.translation_table:
                    self.translation_table[original_text] = {}
                
                # 更新各语言翻译
                for i, lang in enumerate(import_langs):
                    if i + 1 < len(row) and row[i + 1]:
                        old_value = self.translation_table[original_text].get(lang, '')
                        if old_value != row[i + 1]:
                            self.translation_table[original_text][lang] = row[i + 1]
                            updated_count += 1
            
            # 保存翻译表
            self.save_translation_table()
            
            self.log(f"从Excel导入完成，共更新 {updated_count} 条翻译")
            messagebox.showinfo("完成", f"成功导入 {updated_count} 条翻译")
        except Exception as e:
            messagebox.showerror("错误", f"导入Excel失败: {str(e)}")

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def run_as_admin(self):
        if platform.system() != "Windows":
            return False
        
        script = os.path.abspath(__file__)
        params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
        
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
            return True
        except Exception as e:
            self.log(f"提权失败: {e}")
            return False

    def check_and_request_admin(self):
        if not self.is_admin():
            self.log("检测到需要管理员权限，正在请求提权...")
            if self.run_as_admin():
                self.log("程序将以管理员权限重新启动...")
                sys.exit(0)
            else:
                self.log("提权失败，程序将继续以普通权限运行")
                return False
        return True

    def validate_image(self, image_path):
        try:
            test_image = Image.open(image_path)
            test_image.verify()
            test_image.close()
            return True
        except Exception:
            return False

    def clear_icon_cache(self):
        try:
            user_profile = os.path.expandvars("%USERPROFILE%")
            cache_paths = [
                os.path.join(user_profile, "AppData", "Local", "IconCache.db"),
                os.path.join(user_profile, "AppData", "Local", "Microsoft", "Windows", "Explorer"),
            ]
            
            cache_cleared = False
            
            for cache_path in cache_paths:
                if os.path.exists(cache_path):
                    if os.path.isfile(cache_path):
                        try:
                            os.remove(cache_path)
                            self.log(f"已删除图标缓存文件: {cache_path}")
                            cache_cleared = True
                        except Exception as e:
                            self.log(f"无法删除图标缓存文件 {cache_path}: {e}")
                    elif os.path.isdir(cache_path):
                        try:
                            import glob
                            icon_cache_files = glob.glob(os.path.join(cache_path, "iconcache*.db"))
                            for cache_file in icon_cache_files:
                                try:
                                    os.remove(cache_file)
                                    self.log(f"已删除图标缓存文件: {cache_file}")
                                    cache_cleared = True
                                except Exception as e:
                                    self.log(f"无法删除图标缓存文件 {cache_file}: {e}")
                        except Exception as e:
                            self.log(f"扫描图标缓存目录失败: {e}")
            
            if cache_cleared:
                try:
                    SHCNE_ASSOCCHANGED = 0x08000000
                    SHCNF_IDLIST = 0x0000
                    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
                    self.log("已刷新图标缓存（无需重启资源管理器）")
                    return True
                except Exception as e:
                    self.log(f"刷新图标缓存失败: {e}")
                    return True
            else:
                self.log("未找到需要清理的图标缓存文件")
                return True
        except Exception as e:
            self.log(f"清理图标缓存时出错: {e}")
            return False

    def select_logo_image(self):
        file_types = [
            ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.ico"),
            ("所有文件", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择图片文件",
            filetypes=file_types
        )
        
        if filename:
            self.logo_image_path = filename
            self.logo_file_label.config(text=os.path.basename(filename))
            self.logo_action_btn.config(state='normal')
            self.logo_status_var.set(f"已选择: {os.path.basename(filename)}")

    def convert_logo_images(self):
        if not self.logo_image_path:
            messagebox.showwarning("警告", "请先选择图片文件")
            return
        
        try:
            self.logo_progress.start()
            self.logo_status_var.set("正在转换图片...")
            self.root.update()
            
            target_dir = os.path.join(self.output_dir, "OpenCNC", "Bin", "Logo")
            os.makedirs(target_dir, exist_ok=True)
            
            original_image = Image.open(self.logo_image_path)
            
            ico_filename = self.ico_filename_var.get().strip() or "LOGO"
            ico_path = os.path.join(target_dir, f"{ico_filename}.ico")
            
            ico_image = original_image.resize((128, 128), Image.Resampling.LANCZOS)
            
            if ico_image.mode == 'RGBA':
                pass
            elif ico_image.mode == 'LA':
                ico_image = ico_image.convert('RGBA')
            elif ico_image.mode != 'RGB':
                ico_image = ico_image.convert('RGBA')
            
            ico_image.save(ico_path, format='ICO')
            
            gif_image = original_image.resize((75, 94), Image.Resampling.LANCZOS)
            gif_path = os.path.join(target_dir, "LoadingImage.gif")
            
            if gif_image.mode == 'RGBA':
                pass
            elif gif_image.mode == 'LA':
                gif_image = gif_image.convert('RGBA')
            elif gif_image.mode != 'RGB':
                gif_image = gif_image.convert('RGBA')
            
            gif_image.save(gif_path, format="GIF")
            
            ico_valid = self.validate_image(ico_path)
            gif_valid = self.validate_image(gif_path)
            
            self.logo_progress.stop()
            
            if ico_valid and gif_valid:
                messagebox.showinfo("成功", f"图片转换完成！\n文件已保存到: {target_dir}")
                self.logo_status_var.set("转换完成")
                self.log(f"Logo图片转换完成: {ico_path}, {gif_path}")
            else:
                messagebox.showwarning("警告", "文件已生成，但验证时发现可能存在问题")
                self.logo_status_var.set("转换完成（有警告）")
                
        except Exception as e:
            self.logo_progress.stop()
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            self.logo_status_var.set("转换失败")
            self.log(f"Logo转换失败: {e}")

    def modify_logo_and_create_shortcut(self):
        if not self.logo_image_path:
            messagebox.showwarning("警告", "请先选择图片文件")
            return
        
        try:
            self.logo_progress.start()
            self.logo_status_var.set("正在转换图片格式...")
            self.root.update()
            
            target_dir = os.path.join(self.output_dir, "OpenCNC", "Bin", "Logo")
            os.makedirs(target_dir, exist_ok=True)
            
            original_image = Image.open(self.logo_image_path)
            
            ico_filename = self.ico_filename_var.get().strip() or "LOGO"
            ico_path = os.path.join(target_dir, f"{ico_filename}.ico")
            
            ico_image = original_image.resize((128, 128), Image.Resampling.LANCZOS)
            
            if ico_image.mode == 'RGBA':
                pass
            elif ico_image.mode == 'LA':
                ico_image = ico_image.convert('RGBA')
            elif ico_image.mode != 'RGB':
                ico_image = ico_image.convert('RGBA')
            
            ico_image.save(ico_path, format='ICO')
            
            gif_image = original_image.resize((75, 94), Image.Resampling.LANCZOS)
            gif_path = os.path.join(target_dir, "LoadingImage.gif")
            
            if gif_image.mode == 'RGBA':
                pass
            elif gif_image.mode == 'LA':
                gif_image = gif_image.convert('RGBA')
            elif gif_image.mode != 'RGB':
                gif_image = gif_image.convert('RGBA')
            
            gif_image.save(gif_path, format="GIF")
            
            ico_valid = self.validate_image(ico_path)
            gif_valid = self.validate_image(gif_path)
            
            if not (ico_valid and gif_valid):
                messagebox.showwarning("警告", "图片转换完成，但验证时发现可能存在问题")
            
            self.logo_status_var.set("正在创建桌面快捷方式...")
            self.root.update()
            
            exe_path = os.path.join(self.output_dir, "OpenCNC", "Bin", "SyntecLaserMarking.exe")
            
            if not os.path.exists(exe_path):
                messagebox.showerror("错误", f"可执行文件不存在: {exe_path}\n\n请先通过「DiskC工作流」设置正确的DiskC根目录")
                self.logo_progress.stop()
                return
            
            self.logo_status_var.set("正在刷新图标缓存...")
            self.root.update()
            self.clear_icon_cache()
            
            script_dir = os.path.dirname(os.path.abspath(__file__))
            vbs_script = os.path.join(script_dir, "快速设定logo和快捷方式", "create_shortcut.vbs")
            
            if os.path.exists(vbs_script):
                bin_dir = os.path.join(self.output_dir, "OpenCNC", "Bin")
                company_name = self.ico_filename_var.get().strip() or ""
                # 传递: exePath, workingDir, iconFilename, companyName
                result = subprocess.run([
                    "cscript", "//Nologo", vbs_script,
                    exe_path, bin_dir, ico_filename, company_name
                ], capture_output=True, text=True)
                
                self.logo_progress.stop()
                
                if result.returncode == 0:
                    messagebox.showinfo("成功", "LOGO修改完成！桌面快捷方式创建成功！图标缓存已刷新。")
                    self.logo_status_var.set("LOGO修改和快捷方式创建成功")
                    self.log("Logo和快捷方式创建成功")
                else:
                    messagebox.showerror("错误", f"创建快捷方式失败: {result.stderr}")
                    self.logo_status_var.set("快捷方式创建失败")
            else:
                self.logo_progress.stop()
                messagebox.showerror("错误", f"VBScript脚本文件不存在: {vbs_script}")
                
        except Exception as e:
            self.logo_progress.stop()
            messagebox.showerror("错误", f"操作失败: {str(e)}")
            self.logo_status_var.set("操作失败")
            self.log(f"Logo和快捷方式操作失败: {e}")

    def create_shortcut_only(self):
        try:
            exe_path = os.path.join(self.output_dir, "OpenCNC", "Bin", "SyntecLaserMarking.exe")
            ico_filename = self.ico_filename_var.get().strip() or "LOGO"
            ico_path = os.path.join(self.output_dir, "OpenCNC", "Bin", "Logo", f"{ico_filename}.ico")
            
            if not os.path.exists(exe_path):
                messagebox.showerror("错误", f"可执行文件不存在: {exe_path}\n\n请先通过「DiskC工作流」设置正确的DiskC根目录")
                return
            
            if not os.path.exists(ico_path):
                messagebox.showwarning("警告", "请先转换图片生成logo.ico文件")
                return
            
            self.logo_status_var.set("正在清理图标缓存...")
            self.root.update()
            self.clear_icon_cache()
            
            vbs_script = os.path.join(script_dir, "快速设定logo和快捷方式", "create_shortcut.vbs")
            
            if os.path.exists(vbs_script):
                bin_dir = os.path.join(self.output_dir, "OpenCNC", "Bin")
                company_name = self.ico_filename_var.get().strip() or ""
                # 传递: exePath, workingDir, iconFilename, companyName
                result = subprocess.run([
                    "cscript", "//Nologo", vbs_script,
                    exe_path, bin_dir, ico_filename, company_name
                ], capture_output=True, text=True)
                
                if result.returncode == 0:
                    messagebox.showinfo("成功", "桌面快捷方式创建成功！图标缓存已刷新。")
                    self.logo_status_var.set("快捷方式创建成功")
                    self.log("快捷方式创建成功")
                else:
                    messagebox.showerror("错误", f"创建快捷方式失败: {result.stderr}")
                    self.logo_status_var.set("快捷方式创建失败")
            else:
                messagebox.showerror("错误", f"VBScript脚本文件不存在: {vbs_script}")
                
        except Exception as e:
            messagebox.showerror("错误", f"创建快捷方式时出错: {str(e)}")
            self.logo_status_var.set("快捷方式创建失败")
            self.log(f"创建快捷方式失败: {e}")


if __name__ == '__main__':
    """
    程序入口
    创建主窗口并启动应用
    """
    parser = argparse.ArgumentParser(description='自动化翻译处理工具（GUI/命令行）')
    parser.add_argument('--input', '-i', help='输入文件或文件夹路径（支持XML或RES）')
    parser.add_argument('--mode', choices=['batch', 'single'], default='batch', help='处理模式: batch=文件夹, single=单文件')
    parser.add_argument('--output', '-o', help='输出目录（默认为脚本目录）')
    parser.add_argument('--langs', help='逗号分隔的目标语言缩写（例如: CHS,ENG）')
    parser.add_argument('--api', choices=list(API_PROVIDERS), help='翻译服务，例如 deepl_free、baidu、tencent')
    parser.add_argument('--pack', action='store_true', help='处理后是否打包为.res')
    parser.add_argument('--diskc', help='DiskC工作流模式: 指定DiskC根目录路径')
    parser.add_argument('--mb-backup', help='MB备份工作流模式: 指定MB备份ZIP或已解压文件夹路径')
    parser.add_argument('--simulator', help='模拟器工作流模式: 指定DISKC_V1.1.6或DiskC根目录路径')
    parser.add_argument('--log', help='将详细日志写入指定文件')
    parser.add_argument('--nogui', action='store_true', help='无界面模式（仅命令行）')
    args = parser.parse_args()

    if args.nogui or args.input or args.diskc or args.mb_backup or args.simulator:
        # headless / CLI 模式
        root = tk.Tk()
        root.withdraw()
        app = TranslationApp(root)
        app.headless = True
        if args.api:
            app.api_type = args.api
        # 设置输出目录
        if args.output:
            app.output_dir = args.output
        # 准备日志
        if args.log:
            app._ensure_logger(args.log)
        else:
            app._ensure_logger()

        # 设置打包选项
        app.pack_var.set(bool(args.pack))

        # 设置语言列表
        if args.langs:
            langs = [s.strip() for s in args.langs.split(',') if s.strip()]
            app.selected_langs = langs
        else:
            app.selected_langs = app.default_langs.copy()

        # 准备输入文件列表
        if args.simulator:
            if os.path.isdir(args.simulator):
                app.selected_workflow = 'simulator'
                app.log(f'模拟器工作流模式: {args.simulator}')
                app.setup_simulator_sources(args.simulator)
            else:
                app.log(f'模拟器路径无效: {args.simulator}')
                exit(1)
        elif args.mb_backup:
            if os.path.isfile(args.mb_backup) or os.path.isdir(args.mb_backup):
                app.selected_workflow = 'mb_backup'
                app.log(f'MB备份工作流模式: {args.mb_backup}')
                app.setup_mb_backup_sources(args.mb_backup)
            else:
                app.log(f'MB备份文件路径无效: {args.mb_backup}')
                exit(1)
        elif args.diskc:
            if os.path.isdir(args.diskc):
                app.log(f'DiskC工作流模式: {args.diskc}')
                app.setup_diskc_sources(args.diskc)
            else:
                app.log(f'DiskC路径无效: {args.diskc}')
                exit(1)
        elif args.input:
            if os.path.isdir(args.input) and args.mode == 'batch':
                app.source_folder = args.input
                app.refresh_file_list()
            elif os.path.isfile(args.input) and args.mode == 'single':
                # 单文件处理
                prepared = app.prepare_file(args.input)
                if prepared:
                    app.source_files = [prepared]
                else:
                    app.log('无法准备输入文件，退出')
                    exit(1)
            else:
                # 如果用户指定文件，但没有选择single模式，尝试作为单文件处理
                if os.path.isfile(args.input):
                    prepared = app.prepare_file(args.input)
                    if prepared:
                        app.source_files = [prepared]
                    else:
                        app.log('无法准备输入文件，退出')
                        exit(1)
                else:
                    app.log('输入路径无效')
                    exit(1)

        # 启动处理流程（同步等待完成）
        app.log('命令行模式: 开始处理')
        app.start_translation()

        # 等待翻译完成（轮询）
        try:
            while not app.translation_complete:
                time.sleep(0.5)
        except KeyboardInterrupt:
            app.log('用户中断')
            exit(2)

        if app.translation_failures:
            app.log(f'命令行模式: 检测到 {len(app.translation_failures)} 条失败，未生成 XML 文件')
            if app.logger:
                app.logger.error('翻译失败，命令行以状态码 3 退出')
            exit(3)

        # 生成XML
        app.generate_xml_files()
        app.save_translation_table()
        app.log('命令行模式: 处理完成')
        if app.logger:
            app.logger.info('处理完成')
        exit(0)
    else:
        root = tk.Tk()
        app = TranslationApp(root)
        root.mainloop()