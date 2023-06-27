# coding=utf8
import requests
import json
import os
from requests_toolbelt import MultipartEncoder
import time

# from pyepl.cp_system_evaluation.system_evaluation_pipeline.feishu_handler import FeishuHandler


# 飞书文档对象
class FeishuTextRun():
    def __init__(self, text="", style={}):
        self.type = 'textRun'
        self.textRun = {
            'text': text,
            'style': style,
        }

class FeishuParagraph():
    def __init__(self, style={'align': 'center'}):
        self.type = 'paragraph'
        self.paragraph = {'elements': []}
        self.style = style

    def add_element(self, element):
        self.paragraph['elements'].append(element.__dict__)
    
class FeishuTableCell():
    def __init__(self, col_index=0):
        self.columnIndex = col_index
        self.body = {'blocks': []}
        
    def add_block(self, block):
        self.body['blocks'].append(block.__dict__)
   
class FeishuTableRow():
    def __init__(self, row_index=0):
        self.rowIndex = row_index
        self.tableCells = []
        
    def add_cell(self, cell: FeishuTableCell):
        self.tableCells.append(cell.__dict__)
        
    def add_para_cell(self, col_index, text, style):
        cell = FeishuTableCell(col_index)
        for i in range(len(text)):
            text_run = FeishuTextRun(text[i], style[i])
            para = FeishuParagraph()
            para.add_element(text_run)
            cell.add_block(para)
        self.add_cell(cell)
    
    def add_empty_cell(self, col_index):
        cell = FeishuTableCell(col_index)
        self.add_cell(cell)

class FeishuMergedCell():
    def __init__(self, row_start, row_end, col_start, col_end):
        self.rowStartIndex = row_start
        self.rowEndIndex = row_end
        self.columnStartIndex = col_start
        self.columnEndIndex = col_end

class FeishuTable():
    def __init__(self, column_size=0, row_size=0):  
        self.type = 'table'
        self.row_index = 0
        self.table = {
            'rowSize': row_size,
            'columnSize': column_size,
            'tableRows': [],
            'mergedCells': [],
            'tableStyle': {'tableColumnProperties': []}
        }
    
    def set_col_properties(self, col_properties):
        self.table['tableStyle']['tableColumnProperties'] = col_properties
    
    def add_row(self, row: FeishuTableRow):
        row.rowIndex = self.row_index
        self.row_index += 1
        self.table['tableRows'].append(row.__dict__)
        
    def add_merged_cell(self, merged_cell):
        self.table['mergedCells'].append(merged_cell.__dict__)


# 飞书文档操作
class DocPub():
    def __init__(self):
        self.app_id = ''
        self.app_secret = ''
        self.root_folder_token = ''      
        self.parent_token = ''           
        self.ext_to_type = {'docx': 'docx', 'doc': 'docx', 'txt': 'docx', 'md': 'docx', 'mark': 'docx', 'markdown': 'docx', 'html': 'docx',  # 除html外，也可以将type设置为doc而非docx
                       'xlsx': 'sheet', 'xls': 'sheet', 'csv': 'sheet'}         # xlsx和csv也可以将type设置为 bitable

    # 操作API之前的令牌获取
    def get_tenant_access_token(self):
        url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'
        headers = {'Content-type': 'application/json; charset=utf-8'}
        data = {'app_id': self.app_id, 'app_secret': self.app_secret}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        return res['tenant_access_token']

    # 调用飞书 API 来获取根文件夹的 token ： 操作根目录 app_id和app_secret不变则不变
    def get_root_folder_token(self):
        url = 'https://open.feishu.cn/open-apis/drive/explorer/v2/root_folder/meta'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Authorization': authorization}
        res = json.loads(requests.get(url, headers=headers).text)
        folder_root_token = str(res['data']['token'])

    # 获取sheet表格的 - sheet id
    # 获取表格数据: 路径参数
    # get的方式
    def read_sheet(self, spreadsheetToken):
        url = f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheetToken}/metainfo'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        res = json.loads(requests.get(url, headers=headers).text)
        # print("sheetId = ", res['data']['sheets'][0]['sheetId'])
        print(res)
        return res['data']['sheets'][0]['sheetId']

    # 获取sheet表格信息（多个子表格table）- 输入sheet token
    # token: sht 开头
    def read_table_by_sheet_token(self, sheet_token):
        url = 'https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/' + sheet_token + '/sheets/query'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        res = json.loads(requests.get(url, headers=headers).text)
        # print(res)
        return res

    # 获取wiki空间节点信息 - wikcn -xxx 开头的url
    # 从节点信息中获取token --> sheet的token
    # 文档的节点token
    def get_wiki_node(self, file_token):
        url = f'https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={file_token}'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Authorization': authorization}
        res = json.loads(requests.get(url, headers=headers).text)
        print(res)
        return res['data']['node']['obj_token']

    # 读取知识库节点下的某个特定子节点 parent token
    # 需要分页
    # 直接把parent token下所有页数保存读取
    def get_wiki_node_by_parent(self, parent_node_token, space_id='', page_size = '', page_token = None):
        all_nodes = []
        count = 0
        while True:
            count += 1
            # 构建URL
            url = 'https://open.feishu.cn/open-apis/wiki/v2/spaces/' + space_id + '/nodes?parent_node_token=' \
                  + parent_node_token + '&page_size=' + page_size

            if page_token:
                url += '&page_token=' + page_token

            authorization = 'Bearer ' + self.get_tenant_access_token()
            headers = {'Authorization': authorization}
            res = json.loads(requests.get(url, headers=headers).text)

            nodes = res['data']['items']
            all_nodes.extend(nodes)

            # 测试
            # if count == 5:
            #     break
            # 'page_token' in res['data'] and
            if res['data']['has_more'] == True:
                # 如果存在下一页，更新page_token并继续循环
                page_token = res['data']['page_token']
            else:
                break
        # print(all_nodes)
        return all_nodes

    # 读取单元格数据
    # 0b**12!A:B
    def read_sheet_info(self, spreadsheetToken, sheet_id, start_col='A', end_col='C'):
        url = f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheetToken}/values/{sheet_id}!' \
              f'{start_col}:{end_col}'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        res = json.loads(requests.get(url, headers=headers).text)
        return res

    # 创建sheet - 并获取对应的sheet token
    def create_sheet(self):
        url = 'https://open.feishu.cn/open-apis/sheets/v3/spreadsheets'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        data = {'FolderToken': self.root_folder_token}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        print(res)
        # print("sheet token = ", res['data']['spreadsheet']['spreadsheet_token'])
        return res['data']['spreadsheet']['spreadsheet_token']

    # sheet追加数据
    def sheet_add_values(self, sheet_token, sheet_id, values, start_unit='A', end_unit='B'):
        url = 'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/' + sheet_token + '/values_append'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        data = {
            "valueRange": {
                "range": sheet_id + "!" + start_unit + ":" + end_unit,
                "values": values
            }
        }
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        print(res)
        return res["code"]

    # 创建文档
    def create_doc(self, content):
        url = 'https://open.feishu.cn/open-apis/doc/v2/create'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        data = {'FolderToken': self.root_folder_token , 'Content':content}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        # print('######## create doc ###########')
        # print(res)
        return res['data']['objToken']

    # parent_token 知识空间下几级节点 url: 知识空间的space_id
    # 指定文档移动到特定 wiki 空间
    def move_doc_to_wiki(self, obj_token, parent_token=None, obj_type='sheet'):
        if not parent_token:
            parent_token = self.parent_token
        url = 'https://open.feishu.cn/open-apis/wiki/v2/spaces/space_id/nodes/move_docs_to_wiki'   
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        data = {'parent_wiki_token': parent_token, 'obj_type': obj_type, 'obj_token': obj_token}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        print(res)
        if 'wiki_token' in res['data']:
            return res['data']['wiki_token']
        status = 1
        task_id = res['data']['task_id']
        while status == 1:
            url = 'https://open.feishu.cn/open-apis/wiki/v2/tasks/' + task_id     
            authorization = 'Bearer ' + self.get_tenant_access_token()
            headers = {'Authorization': authorization, 'task_type': 'move'}
            data = {'task_type': 'move'}
            res = json.loads(requests.get(url, params=data, headers=headers).text)
            # print(res)
            status = res['data']['task']['move_result'][0]['status']
            if status == 0:
                return res['data']['task']['move_result'][0]['node']['node_token']
            elif status == -1:
                print('######## move doc to wiki failed! ' + obj_token + '########') 
                return ''

    # obj_type可选值: doc~文档，sheet~电子表格，file~云空间文件，wiki~知识库节点，bitable~多维表格，docx~新版文档，folder~文件夹，mindnote~思维笔记
    # member_type可选值: email~飞书邮箱，openid~开放平台ID，openchat~开放平台群组ID，opendepartmentid~开放平台部门ID，userid~用户自定义ID
    # perm可选值: view~可阅读角色，edit~可编辑角色，full_access~可管理角色
    def chmod_doc(self, obj_token, perm, obj_type='wiki', member_id=' ', member_type='email'):
        url = 'https://open.feishu.cn/open-apis/drive/v1/permissions/' + obj_token + '/members?type=' + obj_type
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        data = {'member_type': member_type, 'member_id': member_id, 'perm': perm}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        # print('######## chmod doc ###########')
        # print(res)

    def delete_doc(self, obj_token):
        url = 'https://open.feishu.cn/open-apis/drive/explorer/v2/file/docs/' + obj_token
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Authorization': authorization}
        res = json.loads(requests.delete(url, headers=headers).text)
        # print(res)
        
    def file_upload(self, file_path):
        url = 'https://open.feishu.cn/open-apis/drive/v1/files/upload_all'
        form = {
            'file_name': os.path.basename(file_path), 
            'parent_type': 'explorer', 
            'parent_node': self.root_folder_token,
            'size': str(os.path.getsize(file_path)),
            'file': (open(file_path, 'rb'))
            }
        multi_form = MultipartEncoder(form)
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': multi_form.content_type, 'Authorization': authorization}
        res = json.loads(requests.post(url, data=multi_form, headers=headers).text)
        if res['code'] == 0:
            return res['data']['file_token']
        else:
            print('######## medias upload failed! ###########')
            print(res)
        
    # https://open.feishu.cn/document/uAjLw4CM/ukTMukTMukTM/reference/drive-v1/import_task/create
    def import_task(self, file_path, file_token):
        url = 'https://open.feishu.cn/open-apis/drive/v1/import_tasks'
        authorization = 'Bearer ' + self.get_tenant_access_token()
        headers = {'Content-type': 'application/json; charset=utf-8', 'Authorization': authorization}
        file_extension = file_path.split('.')[-1]
        data = {
            'file_extension': file_extension, 
            'file_token': file_token, 
            'type': self.ext_to_type[file_extension],
            'point':{
                'mount_type': 1, 
                'mount_key': self.root_folder_token
                }
            }
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        if res['code'] == 0:
            return res['data']['ticket']
        else:
            print('######## import task failed! ###########')
            print(res)
    
    def get_import_status(self, ticket):    
        url = 'https://open.feishu.cn/open-apis/drive/v1/import_tasks/' + ticket
        headers = { 'Authorization':'Bearer ' + self.get_tenant_access_token() }
        res = json.loads(requests.get(url, headers=headers).text)
        if res['code'] != 0:
            print('######## get import status failed! ###########')
            print(res)
        else:
            return res['data']['result']
            
    # mount_key: mount_key是目标挂载目录，如为空，则挂载到导入者的云空间根目录下
    # 详见medias_upload()方法
    def import_doc(self, file_path, parent_node=None):
        try:
            token = self.file_upload(file_path)
            print('######## file uploaded! ###########')
            ticket = self.import_task(file_path, token)
            print('######## importing task... ###########')
            res = self.get_import_status(ticket)
            while res['job_status'] == 1 or res['job_status'] == 2:                       # 1:初始化，2:处理中
                time.sleep(1)
                res = self.get_import_status(ticket)
            print('######## doc imported ###########')
            self.chmod_doc(res['token'], 'full_access', 'sheet')                          # 0表示成功，这里不判断。非0由异常捕获
            obj_type = self.ext_to_type[file_path.split('.')[-1]]
            token = self.move_doc_to_wiki(res['token'], parent_node, obj_type)
            url = 'https://momenta.feishu.cn/wiki/' + token
            return url
        except Exception as e:
            print('######## exception occured importing doc! ###########')
            print(type(e), e)
            


# 飞书消息
class MsgSend:
    def __init__(self):
        self.app_id = ""
        self.app_secret = ""
        self.chat_ids = {
            "daily_test": "",        
            "pnpp_and_system": "",  
            "crash_alert": "",       
            "version_pub": "",       
            "test": ""               
            }
        self.user_ids= {}
        self.tenant_access_token = ""
        self.last_tenant_access_time = ""

    def get_tenant_access_token(self):
        url = 'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal'
        headers = {'Content-type': 'application/json; charset=utf-8'}
        data = {'app_id': self.app_id, 'app_secret': self.app_secret}
        data_json = json.dumps(data, ensure_ascii=True).encode('utf-8')
        res = json.loads(requests.post(url, data=data_json, headers=headers).text)
        return res['tenant_access_token']
    
    def get_chats_id(self):     
        url = "https://open.feishu.cn/open-apis/im/v1/chats"
        headers = {
            "Content-type": "application/json; charset=utf-8",
            "Authorization":"Bearer " + self.get_tenant_access_token(),
            "Connection":"close"
        }
        res = json.loads(requests.get(url, headers=headers).text)["data"]["items"]
        ret = {}
        for item in res:
            ret[item["name"]] = item["chat_id"]
        return ret
    
    # 获取飞书user_id
    def get_user_id(self, name):       
        mail_address = name + "@momenta.ai"
        url = "https://open.feishu.cn/open-apis/user/v1/batch_get_id?emails=" + mail_address
        headers = {
            "Content-type": "application/json; charset=utf-8",
            "Authorization":"Bearer " + self.get_tenant_access_token(),
            "Connection":"close"
        }
        res = json.loads(requests.get(url, headers=headers).text)["data"]["email_users"]
        # 收到信息后将user_id加到代码里，后续就不用请求了
        self.user_ids[name] = res[mail_address][0]["user_id"]
        return self.user_ids[name]

    # 获取文档内容(适用于飞书旧版文档)
    # 要创建文档，可以先在飞书上创建一篇文档，再用这个方法获取文档模板
    def get_doc_content(self, docToken):
        url = 'https://open.feishu.cn/open-apis/doc/v2/' + docToken + '/content'
        headers = {
            "Content-type": "application/json; charset=utf-8",
            "Authorization":"Bearer " + self.get_tenant_access_token(),
        }
        res = json.loads(requests.get(url, headers=headers).text)
        with open('doc_content_' + docToken, 'w+') as f:
            json.dump(res, f)     
        return res
    
    # 发送单聊消息
    def send_msg_user(self, msg, user_name):
        if user_name not in self.user_ids:           
                self.get_user_id(user_name)
        self.send_msg("user_id", self.user_ids[user_name], msg) 

    # 发送群聊消息
    # @at_list: 艾特用户名列表，其中的用户会被艾特。"all"艾特所有人。可以为空。
    def send_msg_group(self, msg, chat_name, at_list=[]):
        if chat_name not in self.chat_ids:
            msg = "group " + chat_name + " does not exist, please check!\n The msg is:\n" + msg
            self.send_msg_user(msg, "feifei.xu")
            return   
        if "all" not in at_list:
            for name in at_list:
                if name not in self.user_ids:           
                    self.get_user_id(name)
        self.send_msg("chat_id", self.chat_ids[chat_name], msg, at_list)

    # 发送飞书消息
    # @receive_type: 
    #   "user_id": 给单用户发送私聊信息
    #   "chat_id": 发送群聊消息。默认发送给崩溃报警群
    # @receive_id: 用户的user_id或者群聊的chat_id
    # @msg: 消息内容。当msg_type为"file"时， msg应为文件路径，如"/tmp/test.txt"
    # @at_list: 艾特用户名列表，其中的用户会被艾特。"all"艾特所有人。可以为空。
    def send_msg(self, receive_type, receive_id, msg, at_list=[], msg_type="text"):
        url = "https://open.feishu.cn/open-apis/im/v1/messages"
        params = {"receive_id_type" : receive_type}
        at_str = ""
        if "all" in at_list:
            at_str = "<at user_id=\\\"all\\\">所有人</at>\\n"
        elif len(at_list) != 0:
            for name in at_list:
                at_str += "<at user_id=\\\"" + self.user_ids[name] + "\\\">" + name + "</at>"
            at_str += "\\n"
        context_str = msg
        if msg_type == "text":
            context_str = "{\"text\":\"" + at_str + msg.replace("\n", "\\n") +  "\"}"
        payload = {
            "receive_id": receive_id,
            "content":  context_str,
            "msg_type": msg_type,
        }
        headers = {
            "Authorization": "Bearer " + self.get_tenant_access_token(),
            "Content-Type": "application/json",
            "Connection":"close"
        }
        res = requests.request("POST", url, params=params, headers=headers, data=json.dumps(payload))
        return res


# if __name__ == '__main__':
#     DP = DocPub()