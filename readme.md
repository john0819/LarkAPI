# 飞书API与数据处理项目

# Lark API and Data Processing Project



## 简介

## Introduction

本项目致力于利用飞书API对公司内部知识空间wiki的子目录进行读取、筛选，并对提测报告和每日报告进行批量处理。项目的主要功能包括解析URL形式存在的Jira报告、统一不同来源和格式的报告信息、处理文件内容更新问题等。项目的最终目标是归类并分析公司多阶段的数据EVENT和Jira数据，并为它们添加标签，以便后续的分类和检索。 

This project is dedicated to reading and filtering the subdirectories of the company's internal knowledge space wiki using the Fishu API, and batch processing of the mention test reports and daily reports. The main functions of the project include parsing Jira reports existing in URL form, unifying report information from different sources and formats, and handling file content updates. The ultimate goal of the project is to categorize and analyze the company's multi-stage data EVENT and Jira data, and add tags to them for subsequent classification and retrieval. 



## 背景

## Background

随着公司业务的发展，内部生成的各种报告和数据日益增多，因此对这些数据进行有效管理和利用成了一个挑战。特别是当数据存在于各种不同的系统和格式中时，如知识空间wiki、提测报告、每日报告、Jira报告等，进行统一处理和分析就更加困难。另外，由于历史遗留问题和命名规范问题，使得报告的整理工作变得复杂和混乱。 因此，我开发了这个项目，旨在通过飞书API集成和自动化处理这些数据，使其能够被更好地使用和分析。通过我的努力，现在我们可以更容易地从飞书获取和管理数据，甚至对那些历史零碎的报告文件进行整理，为公司的决策制定提供了数据支持。

As a company's business grows, the number of various reports and data generated internally is increasing, making it a challenge to manage and utilize these data effectively. Especially when the data exists in various systems and formats, such as knowledge space wiki, titration reports, daily reports, Jira reports, etc., it becomes more difficult to conduct uniform processing and analysis. In addition, historical legacy issues and naming convention problems make the organization of reports complicated and confusing. Therefore, I developed this project with the aim of integrating and automating the processing of this data through the FeiBook API so that it can be better used and analyzed. Through my efforts, we can now more easily access and manage data from FeiBook, and even organize those historical fragmented report files to provide data support for company's decision making.

## 安装

以下是在本地环境设置项目的步骤。 首先，克隆这个仓库到你的本地环境



进入项目目录：

```bash
cd your_project
```

然后，建议创建一个新的Python虚拟环境来安装和管理项目的依赖项。使用以下命令可以创建并激活虚拟环境（需要预先安装 `python3-venv`）:

```bash
python3 -m venv env
source env/bin/activate
```

接着，使用pip安装必要的依赖项：

```bash
pip install requests jsonlib os requests_toolbelt time bs4
```

现在，你应该已经安装好了所有所需的依赖项，并可以运行项目了



## 使用

1. 对于知识空间中文档的创建

​		a. 你可以创建针对你任务的文件格式（sheet, word）

​	如果是使用sheet类型的，使用create_wiki_sheet函数，参数为表格的列名

2. 从云文档移动到知识空间下

   a. 由于通过飞书机器人创建的文件放在机器人的知识库下，需要移动到我们自己的知识空间下

​	使用move_doc_to_wiki，并且需要输入相应的知识库id

3. 把自己个人的账号添加权限

   a. 当你添加了个人权限后，你才能对知识空间下的文件进行使用

​	使用chmod_doc函数，输入个人用户的账号，和权限等级

4. 解析报告提取所需要的内容

​		a. write_report_or_wulingshan

​	可以针对自己的任务进行修改



简单步骤：
    1. 先创建sheet获取token和id (FeishuFunction.py)
        create_wiki_sheet
           2. 从云文档移动到知识空间下 (feishu.py)
          move_doc_to_wiki
             3. 给自己账号附加权限 --> 以便于操作 (feishu.py)
            chmod_doc
            4. 读取雾灵山报告、提测报告获取url (FeishuFunction.py)
            get_all_url_in_wulingshan/ get_all_url_in_ticereport
                5. 根据url进行写入 (FeishuFunction.py)
            write_report_or_wulingshan   



1. For creating documents in Knowledge Space a. You can create document formats (sheet, word) specific to your task. If you are using sheet type, use the `create_wiki_sheet` function, with column names as arguments.
2. Moving from Cloud Document to Knowledge Space a. Since the file created through Lark robot is stored under the robot's knowledge base, it needs to be moved to our own knowledge space. Use `move_doc_to_wiki`, and you need to enter the corresponding knowledge base id.
3. Adding access permissions for your personal account a. After you have added personal permissions, you will be able to use the files under the Knowledge Space. Use `chmod_doc` function, inputting individual user's account and permission level.
4. Parsing report to extract required content a. `write_report_or_wulingshan` Modifications can be made depending on your tasks.



Steps:

\1. Create a sheet first to get token and id (`FeishuFunction.py`)

​      `create_wiki_sheet`

​             \2. Move from Cloud Document to Knowledge Space (`feishu.py`)

​                `move_doc_to_wiki`

​                       \3. Add permissions to your own account –> for operation (`feishu.py`)

​                          `chmod_doc`

​                                 \4. Read Wulingshan report/test report to get url (`FeishuFunction.py`)

​                                    `get_all_url_in_wulingshan`/ `get_all_url_in_ticereport`

​                                           \5. Write according to url (`FeishuFunction.py`)

​                                              `write_report_or_wulingshan`

## 使用许可证

由于使用的过程中涉及到token以及用户名，你需要使用你自己的，并申请相应的token

As the process of using involves token and username, you need to use your own and apply for the corresponding token



在对飞书文档操作的类函数 中需要你自己填写一些内容

```python
class DocPub()
```

```python
def __init__(self):
    self.app_id = ''
    self.app_secret = ''
    self.root_folder_token = ''      
    self.parent_token = ''  
```

```python
def get_wiki_node_by_parent(self, parent_node_token, space_id='', page_size = '', page_token = None)
```

```python
def move_doc_to_wiki(self, obj_token, parent_token=None, obj_type='sheet'):
    if not parent_token:
        parent_token = self.parent_token
        url = 'https://open.feishu.cn/open-apis/wiki/v2/spaces/space_id/nodes/move_docs_to_wiki'   
```

```python
def chmod_doc(self, obj_token, perm, obj_type='wiki', member_id=' ', member_type='email'):
```



```python
class MsgSend
```

```python
def __init__(self):
        self.app_id = ""
        self.app_secret = ""
        self.chat_ids = {
            "daily_test": "",        
            "pnpp_and_system": "",  
            "crash_alert": "",       # LPNP路测Crash报警
            "version_pub": "",       
            "test": ""               
            }
        self.user_ids= {}
        self.tenant_access_token = ""
        self.last_tenant_access_time = ""
```



```python
class DocProcess():
    def __init__(self):
        self.doc = DocPub()
        self.sheet_token = ""
        self.sheet_id = ""
        self.report_type = -1 # 日报（直接EL）：1； 提测（jira）：0
        # 提测- 雾灵山- 
        self.write_sheet_token = # 已经创建好了，也可以自己创建新的
        # 提测- 雾灵山- 
        self.write_sheet_id = 
        self.values = [[]] # 读取日报/ 提测的内容
        # 提取jira
        self.username = 
        self.userpwd = 
```



