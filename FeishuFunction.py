from feishu import DocPub
import requests
from bs4 import BeautifulSoup
import re

'''
    下方的sheet token和sheet id --> 可以进行init替换
    雾灵山：
    提测报告：
    步骤：
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
'''

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

    # 获取 token id 还有类型
    '''
        注意：
            这一份是 url中带有wiki-判断为EL，并且判断为
    '''
    def url_to_token(self, url):
        '''
            function:
                把url转换成对应的token，区分两种报告类型
                获取 sheet_token, sheet_id

            Args:
                url: 用户输出 提测报告 or 日报
            Returns:
        '''
        # 提测报告
        if "?sheet=" in url:
            start_token = url.find("wiki/") + len("wiki/")
            end_token = url.find("?")
            wiki_token = url[start_token:end_token]
            self.sheet_token = self.doc.get_wiki_node(wiki_token)
            self.sheet_id = self.doc.read_sheet(self.sheet_token)
            self.report_type = 0
        else:
        # 日报
            parts = url.split("/")
            wiki_token = parts[-1]
            self.sheet_token = self.doc.get_wiki_node(wiki_token)
            self.sheet_id = self.doc.read_sheet(self.sheet_token)
            self.report_type = 1

    # 判断雾灵山和提测 来实现 获取token和id
    def judge_wulingshan_or_report(self, url):
        parts = url.split('/')
        wiki_id = parts[-1]
        # print(wiki_id)
        if "?sheet=" in url:
            # 日报不需要
            # self.sheet_token = self.doc.get_wiki_node(wiki_id.split('=')[-1])
            # self.sheet_id = wiki_id.split('?')[0]
            self.sheet_token = self.doc.get_wiki_node(wiki_id.split('?')[0])
            self.sheet_id = wiki_id.split('=')[-1]
            # print(self.sheet_token)
        else:
            # 雾灵山需要进一步解析
            # shtcntHgqTVS0G8DKbDNXBhUi1d - 日报需要进行修改token
            self.sheet_token = self.doc.get_wiki_node(wiki_id)
            self.sheet_id = self.doc.read_sheet(self.sheet_token)
            # self.sheet_token = wiki_id

            # print(self.sheet_token)
        # print(self.sheet_id)
        # parts = url.split('/')
        # wiki_id = parts[-1]
        # if "?sheet=" in url:
        #     self.sheet_token = wiki_id.split('=')[-1]
        #     self.sheet_id = wiki_id.split('?')[0]
        #
        # else:
        #     self.sheet_token = parts[-1]
        #     self.sheet_id = self.doc.read_sheet(self.sheet_token)

    # 专门jira的 雾灵山+提测
    # 报告都是jira类型的
    # 雾灵山url： 
    # 提测： 
    def read_jira_sheet(self, url, username=None, userpwd=None, start_col='A', end_col='B'):
        if username is None:
            username = self.username
        if userpwd is None:
            userpwd = self.userpwd
        # 判断类型
        self.judge_wulingshan_or_report(url)
        # print(self.sheet_token)
        # print(self.sheet_id)

        sheet_infos = self.doc.read_sheet_info(self.sheet_token, self.sheet_id, start_col, end_col)
        data_values = sheet_infos['data']['valueRange']['values']
        # print(sheet_infos)
        # print('data_values')
        # print(data_values)

        # 筛选第一列为url的内容，不是url的进行剔除
        filtered_data = []
        for item in data_values:
            if isinstance(item[0], list) and any(subitem.get('type') == 'url' for subitem in item[0]):
                filtered_data.append(item)

        # print('filtered_data')
        filtered_data = self.delete_at_info(filtered_data)
        # print(filtered_data)

        # 由于有两种类型 旧版text 新版url
        links = []
        reasons = []
        for entry in filtered_data:
            # 刷新link
            link = []
            url_data = next(subitem for subitem in entry[0] if subitem.get('type') == 'url')

            if 'link' in url_data:
                link = url_data['link']
            elif 'text' in url_data:
                link = url_data['text']

            links.append(link)
            reasons.append(entry[1])

        # print('reasons')
        # print(reasons)

        # print(links)
        # print(reasons)
        # ##### 测试
        self.read_jira_url(links, reasons)
        # print(self.values)

    # 读取 提测报告中event 而不是 jira
    def read_event_sheet(self, url, username=None, userpwd=None, start_col='A', end_col='B'):
        if username is None:
            username = self.username
        if userpwd is None:
            userpwd = self.userpwd
        # 获取token 和 id
        self.judge_wulingshan_or_report(url)
        # print(self.sheet_token)
        # print(self.sheet_id)
        sheet_infos = self.doc.read_sheet_info(self.sheet_token, self.sheet_id, start_col, end_col)
        data_values = sheet_infos['data']['valueRange']['values']

        filtered_data = []
        for item in data_values:
            if isinstance(item[0], list) and any(subitem.get('type') == 'url' for subitem in item[0]):
                filtered_data.append(item)

        filtered_data = self.delete_at_info(filtered_data)
        zipped_data = [(entry[0][0]["link"].split("?id=")[1], entry[0][0]["link"], entry[1]) for entry in filtered_data]
        self.values = [[id, link, reason] for id, link, reason in zipped_data]
        # print(self.values)

    def read_wiki_node_sheet(self, url, username=None, userpwd=None, start_col='A', end_col='B'):
        '''
        function:
            读取已有的wiki中的sheet文档

        Args:
            url: 报告的url
            start_col: 报告读取的第一个单元格
            end_col: 报告读取的最后一个单元格
        Returns:
        '''
        if username is None:
            username = self.username
        if userpwd is None:
            userpwd = self.userpwd

        self.url_to_token(url)
        sheet_infos = self.doc.read_sheet_info(self.sheet_token, self.sheet_id, start_col, end_col)
        data_values = sheet_infos['data']['valueRange']['values']
        print(data_values)
        # print(sheet_infos)
        # 筛选第一列为url的内容，不是url的进行剔除
        filtered_data = []
        for item in data_values:
            if isinstance(item[0], list) and item[0][0].get('type') == 'url':
                filtered_data.append(item)

        # print(filtered_data)
        # 把 @ 的消息删除，仅保留纯文字
        filtered_data = self.delete_at_info(filtered_data)
        print(filtered_data)

        # 日报和提测分开
        # 日报提取三个内容的方法如下
        # 日报：获取三个数据
        # reason消息里面：保存着当时 @ 的人内容（消息有点长）- 已解决
        if self.report_type == 1:
            # 根据 filtered_data 生成 ids、links 和 reasons 的元组列表
            zipped_data = [(entry[0][0]["link"].split("?id=")[1], entry[0][0]["link"], entry[1]) for entry in filtered_data]
            # 将元组列表转换为所需的数据格式
            self.values = [[id, link, reason] for id, link, reason in zipped_data]
            # print(self.values)
        elif self.report_type == 0: # 提测：到jira中进行提取 event_id event_url
            # 把jira的url读取
            # zipped_data = [(entry[0][0]["link"], entry[1]) for entry in filtered_data]
            links = [entry[0][0]["link"] for entry in filtered_data]
            reasons = [entry[1] for entry in filtered_data]
            # print(links)
            self.read_jira_url(links, reasons)
            # print(reasons)

        print(self.values)


    def extract_text(self, data):
        # 提取除了@消息以外的 问题分析
        filter_string = ''
        for item in data:
            if item["type"] == "text":
                filter_string = filter_string + item["text"]
                # return item["text"]
        return filter_string

    def delete_at_info(self, filtered_datas):
        '''
        function:
            把问题分析中的 @ 信息进行剔除，需要区分是否需要进行此操作
            图片也同样进行删除
        Args:
            filtered_datas 需要进行处理的数据

        Returns:
        '''
        process_data = []

        for one_data in filtered_datas:
            # 进行每次刷新
            filtered_last_element = ''
            image_text = False
            # 图片数据：换成None
            if isinstance(one_data[-1], list):
                # 图片类型
                for item in one_data[-1]:
                    if item["type"] == "embed-image":
                        image_text = True
                        break
                if image_text is True:
                    break
                # @类型
                filtered_last_element = self.extract_text(one_data[-1])

            elif isinstance(one_data[-1], str):  # 没有@内容 str
                filtered_last_element = one_data[-1]
            # else: # 包含@的内容
            #     # 轮询从1开始
            #     filtered_last_element = self.extract_text(one_data[-1])

            # @ 消息
            # if isinstance(one_data[-1], list):  # 有@内容
            #     filtered_last_element = self.extract_text(one_data[-1])
            # else:   # 没有@内容
            #     filtered_last_element = one_data[-1]
            if image_text is True:
                output_data = [one_data[0], None]
                process_data.append(output_data)
            else:
                output_data = [one_data[0], filtered_last_element]
                process_data.append(output_data)
        # print(process_data)
        return process_data

    def read_jira_url(self, links, reasons):
        '''
        function:
            根据jira地址来获取 event id 和 event url

        Args:
            links: jira_url
            reasons: reasons

        Returns:
        '''
        results = []  # 用于保存结果的列表，只包含[id, link]

        for url in links:
            response = requests.get(url, auth=(self.username, self.userpwd))
            html_content = response.text
            soup = BeautifulSoup(html_content, "html.parser")

            if response.status_code == 200:
                element = soup.find('div', id='customfield_14660-val')
                div_element = soup.find('div', {'id': 'customfield_14661-val'})

                value = None
                target_url = None

                if element is not None:
                    value = element.text.strip()

                if div_element is not None:
                    target_url = div_element.find('a')['href']

                if value is None and target_url is None:
                    results.append(['None', url])
                else:
                    results.append([value, target_url])
            # if response.status_code == 200:
            #     element = soup.find('div', id='customfield_14660-val')
            #
            #     if element is not None:
            #         value = element.text.strip()
            #
            #         # 查找id为'customfield_14661-val'的<div>元素
            #         div_element = soup.find('div', {'id': 'customfield_14661-val'})
            #
            #         # 从<div>元素中查找<a>元素并获取其href属性
            #         # https://event.momenta.works/events/detail?id= event_id 拼接也可以
            #         # url : https://event.momenta.works/events/detail?id=646c753d3be60e47a53d4eac
            #         target_url = div_element.find('a')['href']
            #
            #         # 将value和target_url组合成一个列表，并添加到results列表中
            #         results.append([value, target_url])

            else:
                print(f"Error: Request failed with status code {response.status_code}")

        final_results = []
        for result, reason in zip(results, reasons):
            final_results.append(result + [reason])
        # print(final_results)
        self.values = final_results
        # print([final_results])

    def create_wiki_sheet(self, attributes_name:list):
        '''
        function:
            第一次创建

        Args:
            attributes_name: 属性，第一行的值，各列

        Returns:
        '''
        # 先插入第一行数据 标题
        # 创建表格
        self.write_sheet_token = self.doc.create_sheet()
        self.write_sheet_id = self.doc.read_sheet(self.write_sheet_token)
        print("####很重要的token 和 id ####")
        print("请妥善保存下方信息")
        print("sheet_token = ", self.write_sheet_token)
        print("sheet_id = ", self.write_sheet_id)
        # 插入数据 头数据
        first_line = [attributes_name]
        res = self.doc.sheet_add_values(self.write_sheet_token, self.write_sheet_id, first_line)
        if res == 0:
            print("create new sheet")

    def write_wiki_sheet(self, start_unit='A', end_unit='B'):
        '''
        function:
            继续插入

        Args:
            start_unit、end_unit 从哪一行开始插入

        Returns:
        '''
        res = self.doc.sheet_add_values(self.write_sheet_token, self.write_sheet_id, self.values, start_unit, end_unit)
        if res == 0:
            print("add values successfully")
            return True
        elif res == 90204:
            print("add values error")
            return False

    # 把 某个日测报告的内容 保存到 知识库下的sheet，单一API
    def read_report_to_write_sheet(self, report_url, username=None, userpwd=None, start_col='A', end_col='B'):
        '''
            function:
                默认都保存在一个 固定的 sheet中，省略第一次创建表格
                读取 日测/ 提测的报告，再把它转换到 专门存储的 sheet中

            Args:
                report_url 所需要读取报告的url
                start_unit、end_unit 从哪一行开始插入
                username 用户名（可以输入自己的）
                userpwd 用户密码（可以输入自己的）
            Returns:
            '''
        # 读取报告的内容保存在 values中
        self.read_wiki_node_sheet(report_url, username, userpwd, start_col, end_col)
        self.write_wiki_sheet(start_col, end_col)

    # 写入提测报告和雾灵山报告
    def write_report_or_wulingshan(self, report_url, username=None, userpwd=None, start_col='A', end_col='B'):
        self.read_jira_sheet(report_url)
        res = self.write_wiki_sheet(start_col, end_col)
        return res # 错误内容进行保存

    # 读取知识空间下某一个parent token的所有内容
    # 获取所有的提测报告
    # title = 'ET车系统集成提测报告 xxx'
    def get_all_url_in_wulingshan(self, parent_token, space_id = None, page_size = None, page_token = None):
        # 读取内容保存在list中（所有页数）
        all_data = self.doc.get_wiki_node_by_parent(parent_token)
        # 检查是否有has_child内容，有的话保存其对应node_token
        data_with_child = []
        for item in all_data:
            if item['has_child']:
                data_with_child.append(item['node_token'])
        # print(data_with_child)
        # 根据node_token进行检查title
        # 遍历所有有child的token进行检查title
        all_data_with_child = []
        for item in data_with_child:
            # 查看子token的title
            signal_result = self.doc.get_wiki_node_by_parent(item)
            # 雾灵山
            # 遍历所有
            for signal_report in signal_result:
                if 'ET车系统集成提测报告' in signal_report['title']:
                    all_data_with_child.append('https://momenta.feishu.cn/wiki/' +  signal_report['node_token'])
        return all_data_with_child
        # print(all_data_with_child)
        # print(all_data_with_child)

    # 提测报告
    def get_all_url_in_ticereport(self, parent_token, space_id=None, page_size=None, page_token=None):
        # 读取内容保存在list中（所有页数）
        all_data = self.doc.get_wiki_node_by_parent(parent_token)
        # 检查是否有has_child内容，有的话保存其对应node_token
        data_with_child = []
        for item in all_data:
            # 先筛选title 必须为 2023年的报告 并且有 has child
            if '2023' in item['title'] and item['has_child']:
                # 所有的 表格token
                data_with_child.append(item['node_token'])
        # print(data_with_child)
        # 遍历所有的table
        all_object_token_list = []
        all_node_token_list = []
        for item in data_with_child:
            # 查看子token的title
            signal_result = self.doc.get_wiki_node_by_parent(item)
            for signal_report in signal_result:
                # 找到表格 需要继续找table
                if '系统集成提测报告' in signal_report['title']:
                    # 读取sheet
                    # 找到子sheet： 问题类型对应jira类型
                    # 保存object_token
                    all_object_token_list.append(signal_report['obj_token'])
                    # 还需要保存 node_token
                    all_node_token_list.append(signal_report['node_token'])
        # 读取table
        sheet_id_list = []
        for signal_object_token in all_object_token_list:
            # 所有table的信息
            table_info = self.doc.read_table_by_sheet_token(signal_object_token)
            # print(table_info)
            # 遍历每个table
            for table in table_info['data']['sheets']:
                if table['title'] == '问题类型对应jira详情':
                    sheet_id_list.append(table['sheet_id'])
                    break

        url_list = []
        # 进行拼接 node_token + sheet_id
        # url = 'https://momenta.feishu.cn/wiki/NgK0wu45eiq53uk51cjcOl6bn2d?sheet=CyuJPi'
        for one_node, one_id in zip(all_node_token_list, sheet_id_list):
            one_url = 'https://momenta.feishu.cn/wiki/' + one_node + '?sheet=' + one_id
            url_list.append(one_url)

        # print(url_list)
        return url_list
        # print(len(sheet_id_list))
        # print(len(all_node_token_list))

    # 现在日报需要用这份，对应event
    def write_report_by_event(self, report_url, username=None, userpwd=None, start_col='A', end_col='B'):
        self.read_event_sheet(report_url)
        res = self.write_wiki_sheet(start_col, end_col)
        return res # 错误内容进行保存