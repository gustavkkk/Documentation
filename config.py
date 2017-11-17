# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 16:11:37 2017

@author: Frank
"""
# PATH TO DATA
PATH_IMAGE = ""
PATH_DOCUMENT = ""
PATH_CERTIFICATE_INDIVIDUAL = ""
PATH_CERTIFICATE_COMPANY = ""
# COMPANY
COMPANY_NAME = ""
COMPANY_FOUNDATION_DATE = ""
COMPANY_ADDRESS = ""
COMPANY_SITE_ADDR = ""
COMPANY_PHONE_NUMBER = ""
COMPANY_FAX_NUMBER = ""
COMPANY_POST_ADDR = ""

COMPANY_OWNER_NAME = ""
COMPANY_OWNER_BIRTHDAY = ""
COMPANY_OWNER_POSITION = ""
COMPANY_OWNER_GREENCARD_NUMBER = ""

COMPANY_QUALITY_LEVEL = ""

COMPANY_CERTIFICATE_NUMBER = ""
COMPANY_REGISTRATION_BANK_NAME = ""
COMPANY_REGISTRATION_BANK_ACCOUNT = ""
COMPANY_REGISTRATION_BUDGET = ""

COMPANY_BUSINESS_SCOPE = ""

COMPANY_LINK_MAN = ""

COMPANY_HUMAN_RESOURCE_STRUCTURE_IMAGE_PATH = ""
COMPANY_HUMAN_RESOURCE_DESCRIPTIONs = []
COMPANY_PROJECT_MANAGER_NUMBER = ""
COMPANY_ENGINEER_INTERMEDIATE_NUMBER = ""
COMPANY_ENGINEER_ELEMENTARY_NUMBER = ""
COMPANY_TECHNICAN_NUMBER = ""
# INDIVIDUAL
PROJECT_MANAGER = ""
TECH_MANAGER = ""
SECURITY_MANAGER = ""
MATERIAL_MANAGER = ""
PLAN_MANAGER = ""
CONSTRUCTION_MANAGER = ""
QUALITY_MANAGER = ""
FINANCE_MANAGER = ""

ID = ""
NAME = ""
POSTION = ""
BIRTHDAY = ""
EDUCATION_BACKGROUND = ""
SCHOOL_OF_GRADUATION = ""
POSITION_TECHNICAL = ""
GREEN_CARD_NUMBER = ""
PHONE_NUMBER = ""
PHONE_NUMBER_MOBILE = ""

CERTIFICATE_NAME = ""
CERTIFICATE_DATE = ""
CERTIFICATE_LEVEL = ""
CERTIFICATE_NUMBER = ""
MAJOR = ""
INSURANCE_STATE = ""

# FINANCE
YEAR = ""
REGISTRATION_BUDGET = ""
PURE_ASSET = ""
TOTAL_ASSET = ""
FIXED_ASSET = ""
FLOATING_ASSET = ""
FLOATING_DEBT = ""
TOTAL_DEBT = ""
OPERATING_RECEIPT = ""
RETAINED_PROFIT = NET_MARGIN = ""

# PROJECT　INFO
PROJECT_PUBLISHER = "" 
PROJECT_NAME = ""
PROJECT_DATE = ""
PROJECT_COST = ""
PROJECT_TIME = ""
BIDDING_DATE = ""
# WORK HISTORY
PROJECT_NAME = ""
PROJECT_LOCATION = ""

PROJECT_PUBLISHER_NAME = ""
PROJECT_PUBLISHER_ADDRESS = ""
PROJECT_PUBLISHER_PHONENUMBER = ""

PROJECT_BUDGET = ""
PROJECT_START_DATE = ""
PROJECT_END_DATE = ""

PROJECT_CONTENT = ""
PROJECT_QUALITY = ""
PROJECT_MANAGER = ""
PROJECT_MANAGER_TECH = ""

PROJECT_DESCRIPTION = ""
PROJECT_REMARKS = ""

index = ['投标文件格式',
        '投标函',
        '投标函附录',
        '承诺书',
        '法定代表人身份证明',
        '授权委托书',
        '联合体协议书',
        '投标保证金',
        '已标价工程量清单',
        '施工组织设计',
        '项目管理机构表',
        '拟分包项目情况表',
        '资格审查资料',
        '原件的复印件',
        '其他材料',
        '资格审查原件登记表',
        '符合性审查表'
        ]

keywords = ['目录',
            '投标文件格式'
           ]

cover = ['投标文件格式',
         '招标编号',
         '项目名称',
         '投标文件',
         '投标人',
         '盖单位章',
         '法定代表',
         '签字',
         '年月日',
        ]

content = [
           '评审因素索引表',
           '标段名称',
           '招标文件',
           '投标总报价',
           '项目负责人',
           '技术负责人',
           '投标人名称',
           '招标人名称',
           '授权委托书',
           '身份证号码',
           '委托代理人',
           '法定代表人',
           '法定代表人身份证复印件',
           '代理人身份证复印件',
           '异议函'
           ]

certificate = [
                '身份证',
                '建造师证',
                '安考A证',
                '安考B证',
                '安考C证',
                '职称证',
                '学历证',
                '上岗证',
                '会计证',                
                ]
dic = index + cover + content