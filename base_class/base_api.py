import requests, time, random, re, json
from baseopensdk.api.base.v1 import *
from baseopensdk import BaseClient

class BaseClass(object):
    def __init__(self, *args, **kwargs):
        self._args = args
        self._kwargs = kwargs

        self._app_token = ""
        self._personal_base_token = ""
        self._table_id = ""
        self._view_id = ""

        self._page_token = ""
        self._filter_info = ""
        self._field_names = ""

        self._record_list = []
        self._record_ids = None

        self._msg = "缺少 {} 参数"

        self._step = 500
        self._number_of_retries = 3
        

    ###########################   判断 value 是否是数字   ###########################
    def is_number(self, value):
        
        try:
            float(value)
            return True
        except ValueError:
            pass
    
        try:
            import unicodedata
            unicodedata.numeric(value)
            return True
        except (TypeError, ValueError):
            pass
    
        return False
    

    ###########################  从 data_list 随机抽取 num 个选项   ###########################
    def random_samples(self, data_list, value):

        sampled_indices = random.sample(data_list, value)

        return sampled_indices
    

    ###########################   判断字符串前后是否包含{{}}   ###########################
    def is_variable(self, string):
        # 判断是否为映射字段模式，^开头$结尾，中间包含两对大括号{}
        pattern = r'^\{\{.*\}\}$'
        if re.match(pattern, string):
            return True
    
        else:
            return False
        
    
    ###########################   基于字段映射表和字段名，从字段返回值中获取实际值   ###########################
    def get_field_value(self, field_type, field_value_item):
        field_value = ""

        ## 文本、Email、条码
        if field_type == 1:
            field_value = field_value_item[0].get("text")

        ## 数字、进度、货币、评分
        elif field_type == 2:
            field_value = str(field_value_item)

        ## 单选
        elif field_type == 3:
            field_value = field_value_item
            
        ## 多选
        elif field_type == 4:
            field_value = ",".join(field_value_item)

        ## 日期、创建时间、最后更新时间
        elif field_type == 5 or field_type == 1001 or field_type == 1002:
            local_time = time.localtime(field_value_item / 1000)
            formatted_local_time = time.strftime('%Y-%m-%d %H:%M:%S', local_time)
            if len(re.findall(r'\s00:00:00$', formatted_local_time)) > 0:
                formatted_local_time = formatted_local_time.replace(" 00:00:00", "")
            field_value = formatted_local_time

        ## 复选框
        elif field_type == 7:
            field_value = str(field_value_item)

        ## 人员、创建人、修改人
        elif field_type == 11 or field_type == 1003 or field_type == 1004:
            user_list = []
            for user_item in field_value_item:
                user_list.append(user_item.get("name"))
            field_value_item = ",".join(user_list)
            field_value = field_value_item

        ## 电话号码
        elif field_type == 13:
            field_value = field_value_item

        ## 超链接
        elif field_type == 15:
            field_value = field_value_item.get("link")

        ## 附件
        elif field_type == 17:
            # print(field_type, field_value_item)
            if "image" in field_value_item[0].get("type"):
                field_value = field_value_item[0].get("file_token")
            else:
                field_value = ""

        ## 单向关联和双向关联
        elif field_type == 18 or field_type == 21:
            try:
                field_value = ",".join(field_value_item.get("link_record_ids"))
            except Exception as e:
                field_value = ""

        ## 公式和查找引用
        elif field_type == 19 or field_type == 20:
            if field_value_item.get("type") == 1:
                field_value = field_value_item.get("value")[0].get("text")
            else:
                field_value = ""

        # 地理位置
        elif field_type == 22:
            field_value = field_value_item.get("full_address")

        # 群组
        elif field_type == 23:
            item_tmp = []
            for group_item in field_value_item:
                item_tmp.append(group_item.get("name"))
            field_value_item = ",".join(item_tmp)
            field_value = field_value_item

        ## 自动编号
        elif field_type == 1005:
            field_value = field_value_item

        else:
            field_value = ""

        return field_value
    

    ###########################  构建client   ###########################
    def _base_client(self):

        if self._app_token == "" or self._personal_base_token == "":
            return self._msg.format("app_token 或 personal_base_token")

        client: BaseClient = BaseClient.builder() \
            .app_token(self._app_token) \
            .personal_base_token(self._personal_base_token) \
            .build()
        
        return client
    

    ###########################   批量创建记录, 数据列表拆分成每500条记录写入一次   ###########################
    def batch_create_record(self, app_token, personal_base_token, table_id, record_list):
        
        self._app_token = app_token
        self._personal_base_token = personal_base_token
        self._table_id = table_id
        self._record_list = record_list

        client = self._base_client()
        
        try:
            
            data = []
            for i in range(0, len(self._record_list), self._step):
                _record_list = self._record_list[i:i + self._step]
                retry = 0
                while retry < self._number_of_retries:
                    try:
                        
                        batch_create_records_request = BatchCreateAppTableRecordRequest().builder() \
                        .table_id(self._table_id) \
                        .request_body(
                            BatchCreateAppTableRecordRequestBody.builder() \
                            .records(_record_list) \
                            .build()
                        ) \
                        .build()

                        batch_create_records_response = client.base.v1.app_table_record.batch_create(
                            batch_create_records_request)
                        retry = self._number_of_retries
                        # print(batch_create_records_response.data.records)
                        for record_item in batch_create_records_response.data.records:
                            data.append(record_item.__dict__)

                    except Exception as e:
                        time.sleep(2)
                        retry = retry + 1
                        if retry == self._number_of_retries:
                            raise "重试超过 {} 次, 数据操作失败, 请检查网络！".format(self._number_of_retries)

            if batch_create_records_response.msg == 'success':
                result = {
                    "code": batch_create_records_response.code,
                    "msg": batch_create_records_response.msg,
                    "data": data
                }
            else:
                result = {
                    "code": batch_create_records_response.code,
                    "msg": batch_create_records_response.msg
                }
            return result

        except Exception as e:
            result = {
                "code": -1,
                "msg": "数据创建失败, 请检查网络或参数后重试"
            }
            return result
    

    ###########################   批量更新记录, 数据列表拆分成每500条记录更新一次   ###########################
    def batch_update_record(self, app_token, personal_base_token, table_id, record_list):
        
        self._app_token = app_token
        self._personal_base_token = personal_base_token
        self._table_id = table_id
        self._record_list = record_list

        client = self._base_client()

        try:
            for i in range(0, len(self._record_list), self._step):
                _record_list = self._record_list[i:i + self._step]
                retry = 0
                while retry < self._number_of_retries:
                    try:
                        batch_update_records_request = BatchUpdateAppTableRecordRequest().builder() \
                        .table_id(self._table_id) \
                        .request_body(
                            BatchUpdateAppTableRecordRequestBody.builder() \
                            .records(_record_list) \
                            .build()
                        ) \
                        .build()

                        batch_update_records_response = client.base.v1.app_table_record.batch_update(
                            batch_update_records_request)
                        retry = self._number_of_retries

                    except Exception as e:
                        time.sleep(2)
                        retry = retry + 1
                        if retry == self._number_of_retries:
                            raise "重试超过 {} 次, 数据操作失败, 请检查网络！".format(self._number_of_retries)
                        
                if batch_update_records_response.msg == 'success':
                    result = {
                        "code": batch_update_records_response.code,
                        "msg": batch_update_records_response.msg,
                        "data": batch_update_records_response.data.records
                    }
                else:
                    result = {
                        "code": batch_update_records_response.code,
                        "msg": batch_update_records_response.msg
                    }

            return result

        except Exception as e:
            result = {
                "code": -1,
                "msg": "数据更新失败, 请检查网络或参数后重试"
            }
            return result

    
    ###########################   调用 base 开放接口查询记录   ###########################
    def search_records(self, app_token: str, personal_base_token : str, table_id: str, view_id: str|None, page_token: str, filter_info: dict|None, field_names: list|None):
    
        self._app_token = app_token
        self._personal_base_token = personal_base_token
        self._table_id = table_id
        self._view_id = view_id
        self._page_token = page_token
        self._filter_info = filter_info
        self._field_names = field_names

        # 查询记录接口
        url = 'https://base-api.feishu.cn/open-apis/bitable/v1/apps/' + self._app_token + '/tables/' + self._table_id + '/records/search'

        req_headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': 'Bearer ' + self._personal_base_token
        }

        req_queries = {
            'page_token': self._page_token,
            'page_size': self._step
        }

        req_body = {}
        if self._view_id != "" or self._view_id is not None:
            req_body['view_id'] = self._view_id

        if self._filter_info is not None:
            req_body['filter'] = self._filter_info

        if self._field_names is not None:
            req_body['field_names'] = self._field_names

        resp = ""
        retry = 0
        while retry < self._number_of_retries:
            try:
                resp = requests.post(url=url, params=req_queries, headers=req_headers, json=req_body)
                # print(resp.json())
                if resp.status_code == 200:
                    retry = self._number_of_retries
                    resp = resp.json()
                    # print(resp)
                elif resp.status_code == 429:
                    print("请求太频繁，等待重试")
                    raise "请求太频繁，等待重试"
            except Exception as e:
                retry = retry + 1
                if retry == self._number_of_retries:
                    print("获取数据超过 {} 次".format(self._number_of_retries))
                    resp = {}
                else:
                    retry_time = retry * 2
                    time.sleep(retry_time)
                    print("正在尝试第 " + str(retry) + " 次查询记录")

            # print(resp.status_code)
        return resp
    

    
    ###########################   调用 base 开放接口批量获取记录   ###########################
    def batch_get_records(self, app_token: str, personal_base_token : str, table_id: str, record_ids: list|None):
    
        self._app_token = app_token
        self._personal_base_token = personal_base_token
        self._table_id = table_id
        self._record_ids = record_ids

        # 批量获取记录接口
        url = 'https://base-api.feishu.cn/open-apis/bitable/v1/apps/' + self._app_token + '/tables/' + self._table_id + '/records/batch_get'

        req_headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': 'Bearer ' + self._personal_base_token
        }

        req_body = {}
        if self._record_ids is not None:
            req_body['record_ids'] = self._record_ids
        req_body['with_shared_url'] = True

        resp = ""
        retry = 0
        while retry < self._number_of_retries:
            try:
                resp = requests.post(url=url, headers=req_headers, json=req_body)
                # print(resp.json())
                if resp.status_code == 200:
                    retry = self._number_of_retries
                    resp = resp.json()
                    # print(resp)
                elif resp.status_code == 429:
                    print("请求太频繁，等待重试")
                    raise "请求太频繁，等待重试"
            except Exception as e:
                retry = retry + 1
                if retry == self._number_of_retries:
                    print("获取数据超过 {} 次".format(self._number_of_retries))
                    resp = {}
                else:
                    retry_time = retry * 2
                    time.sleep(retry_time)
                    print("正在尝试第 " + str(retry) + " 次查询记录")

            # print(resp.status_code)
        return resp
    

    ###########################   调用 base 开放接口获取字段列表   ###########################
    def list_fields(self, app_token: str, personal_base_token : str, table_id: str):
    
        self._app_token = app_token
        self._personal_base_token = personal_base_token
        self._table_id = table_id

        # 批量获取记录接口
        url = 'https://base-api.feishu.cn/open-apis/bitable/v1/apps/' + self._app_token + '/tables/' + self._table_id + '/fields'

        req_headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': 'Bearer ' + self._personal_base_token
        }

        resp = ""
        retry = 0
        while retry < self._number_of_retries:
            try:
                resp = requests.get(url=url, headers=req_headers)
                # print(resp.json())
                if resp.status_code == 200:
                    retry = self._number_of_retries
                    resp = resp.json()
                    # print(resp)
                elif resp.status_code == 429:
                    print("请求太频繁，等待重试")
                    raise "请求太频繁，等待重试"
            except Exception as e:
                retry = retry + 1
                if retry == self._number_of_retries:
                    print("获取数据超过 {} 次".format(self._number_of_retries))
                    resp = {}
                else:
                    retry_time = retry * 2
                    time.sleep(retry_time)
                    print("正在尝试第 " + str(retry) + " 次查询记录")

            # print(resp.status_code)
        return resp
    
    
    ###########################   调用 base 开放接口上传素材   ###########################
    def upload_all(self, personal_base_token : str, multi_form):
    
        self._personal_base_token = personal_base_token

        url = 'https://base-api.feishu.cn/open-apis/drive/v1/medias/upload_all'

        req_headers = {
        'Content-Type': multi_form.content_type,
        'Authorization': 'Bearer ' + self._personal_base_token
        }

        req_body = multi_form

        resp = ""
        retry = 0
        while retry < self._number_of_retries:
            try:
                resp = requests.post(url=url, headers=req_headers, data=req_body)
                # print(resp.json())
                if resp.status_code == 200:
                    retry = self._number_of_retries
                    resp = resp.json()
                    # print(resp)
                elif resp.status_code == 429:
                    print("请求太频繁，等待重试")
                    raise "请求太频繁，等待重试"
            except Exception as e:
                retry = retry + 1
                if retry == self._number_of_retries:
                    print("获取数据超过 {} 次".format(self._number_of_retries))
                    resp = {}
                else:
                    retry_time = retry * 2
                    time.sleep(retry_time)
                    print("正在尝试第 " + str(retry) + " 次查询记录")

            # print(resp.status_code)
        return resp
    

        
    ###########################   调用 base 开放接口下载多维表格中的附件   ###########################
    def download_attachment(self, personal_base_token: str, file_token: str, extra: object):
    
        self._personal_base_token = personal_base_token

        extra = json.dumps(extra)

        url = 'https://base-api.feishu.cn/open-apis/drive/v1/medias/{}/download'.format(file_token)
        
        req_headers = {
            'Host': 'base-api.feishu.cn',
            'Authorization': 'Bearer ' + self._personal_base_token
        }

        req_queries = {
            'extra': extra
        }

        resp = ""
        retry = 0
        while retry < self._number_of_retries:
            try:
                resp = requests.get(url=url, params=req_queries, headers=req_headers)
                # print(resp.__dict__)
                # print(resp.json())
                if resp.status_code == 200:
                    retry = self._number_of_retries
                    # print(resp)
                elif resp.status_code == 429:
                    print("请求太频繁，等待重试")
                    raise "请求太频繁，等待重试"
            except Exception as e:
                retry = retry + 1
                if retry == self._number_of_retries:
                    print("获取数据超过 {} 次".format(self._number_of_retries))
                    resp = {}
                else:
                    retry_time = retry * 2
                    time.sleep(retry_time)
                    print("正在尝试第 " + str(retry) + " 次查询记录")

            # print(resp.status_code)
        return resp
