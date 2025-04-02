from flask import Flask, jsonify, request, render_template, send_file
import os, platform, subprocess, json
import requests, shutil, time, uuid, re, base64
from requests_toolbelt import MultipartEncoder
from docx import Document
from docx.shared import Pt, Inches
from docx.shared import RGBColor
from werkzeug.utils import secure_filename
from docx2pdf import convert
from apscheduler.schedulers.background import BackgroundScheduler
from base_class.base_api import BaseClass
from base_class.generator import generate_qrcode, generate_barcode



app = Flask(__name__, static_folder="static", static_url_path="/static")
app.config['JSON_AS_ASCII'] = False
app.config['JSON_SORT_KEYS'] = False
app.json.ensure_ascii = False

# 当前脚本的目录
fp = os.path.dirname(os.path.abspath(__file__))  

## 定义模板文件的保存路径和文件名尾缀
UPLOAD_FOLDER = os.path.join(fp, 'template_files')
ALLOWED_EXTENSIONS = {'docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

## 独立生成二维码和条形码的保存路径
GENERATE_FOLDER = os.path.join(fp, 'generate_files')
app.config['GENERATE_FOLDER'] = GENERATE_FOLDER


# timestamp = time.time()
# local_time = time.localtime()
# formatted_local_time = time.strftime('%Y-%m-%d %H:%M:%S', local_time)
# formatted_local_time = time.strftime('%Y-%m-%d', local_time)



## 基于上传的模板将多维表格中的记录数据导出到word文件，并回传到当前记录的附件字段中
def export_to_doc(app_token, personal_token, table_id, record_id, info_json, file_name, file_field, field_id_map, file_type):
    '''
        基于上传的模板将多维表格中的记录数据导出到word文件，并回传到当前记录的附件字段中\r\n
        pramas:\r\n
        - app_token: 多维表格ID\r\n
        - personal_token: 多维表格授权码\r\n
        - table_id: 数据表ID\r\n
        - record_id: 记录ID\r\n
        - info_json: 字段名与字段值的映射\r\n
        - file_name: 导出成附件的文件名，默认设定为多维表格中某个字段中的值\r\n
        - file_field: 导出成附件后回传的附件字段名\r\n
        - field_id_map: 字段名与字段ID的映射\r\n
        - file_type: 导出文件类型，默认为docx，当前在linux环境下面部署后导出为pdf有问题\r\n
    '''
    # print(info_json)
    # print("*"*30)
    # print(field_id_map)
    # print("*"*30)

    msg = "生成附件成功"

    # 个人主目录
    main_path = os.path.join(app.config['UPLOAD_FOLDER'], personal_token)
    # print(str(main_path))
    # print("*"*30)
    
    # 个人模板文件路径
    template_file_path = os.path.join(main_path, "template.docx")
    # print(template_file_path)

    # 个人生成的word文件路径
    target_file_path = os.path.join(main_path, file_name + ".docx")
    # print(target_file_path)

    # 个人图片文件路径
    image_file_path = os.path.join(main_path, file_name + ".jpg")

    # 如果模板文件不存在，则直接返回
    if not os.path.isfile(template_file_path):
        print("模板文件不存在，请先上传模板文件")
        return "模板文件不存在，请先上传模板文件"

    # 如果个人生成的word文件存在，则删除
    if os.path.isfile(target_file_path):
        os.remove(target_file_path)

    # 如果个人图片文件存在，则删除
    if os.path.isfile(image_file_path):
        os.remove(image_file_path)

    # 如果个人生成的word文件不存在，则从模板文件复制一个副本
    if not os.path.exists(target_file_path):
        shutil.copy(template_file_path, target_file_path)

    # 基于个人生成word文件副本初始化一个文档实例
    doc = Document(target_file_path)

    ## 遍历文档中的所有文本段落
    for paragraph in doc.paragraphs:
        # for run in paragraph.runs:
        #   print(run.text)
        index = 0
        # 根据获取段落文本，遍历字段名列表进行占位字符的替换
        for key, value in info_json.items():
            # print(key, value)

            # 判断段落文本中是否包含有`{{字段名` 这样的信息，如果存在，则表示存在占位符否则不处理当前段落文本
            if '{{' + key in paragraph.text:
                # print(key, paragraph.text)

                # 遍历段落中的所有文本片断
                for run in paragraph.runs:

                    # 判断文本片断中是否包含有`{{字段名` 这样的信息，如果存在，则表示存在占位符
                    if '{{' + key in run.text:
                        print(key, run.text)
                        
                        # 获取当前文本片断的样式，当前只处理字体大小、颜色、加粗和斜体四种样式
                        font_size = run.font.size  # 假设所有格式相同，这里仅取第一个run的格式
                        color = run.font.color.rgb  # 保存颜色，如果有的话
                        bold = run.font.bold is not None  # 保存粗体状态
                        italic = run.font.italic is not None  # 保存斜体状态
                        
                        # 根据占位符格式对文本片断进行替换
                        try:
                            # 如果占位符包含有`:image`，替换占位符为图片
                            if ":image" in run.text:
                                # 将文本片断中的`{{`和`}}`替换为空，保留有用信息
                                run_text = run.text.replace("{{","").replace("}}","")
                                # 将以上信息用`:`分割，生成列表
                                key_split = run_text.split(":")
                                key = key_split[0]
                                # 如果列表的长度为3，则以上信息中包含有图片尺寸
                                if len(key_split) == 3:
                                    # 对图片尺寸进行分割后并进行变量赋值
                                    size = key_split[2].split("*")
                                    width = float(size[0])
                                    height = float(size[1])

                                # 如果列表的长度不为3，则以上信息中不包含有图片尺寸，则不处理图片大小
                                else:
                                    width = None
                                    height = None
                                    
                                # 生成多维表格附件下载的extra信息，并进行附件下载，返回附件文件的二进制流信息
                                extra = {"bitablePerm":{"tableId":table_id,"attachments":{field_id_map[key]:{record_id:[value]}}}}
                                attachment_resp = BaseClass().download_attachment(personal_token, value, extra)
                                # print(attachment_resp)

                                # 将二进制流信息写入到个人生成的图片附件中
                                with open(image_file_path, 'wb') as file:
                                    file.write(attachment_resp.content)
                                file.close()

                                # 将图片替换到图片占位符所在位置，并把原有的占位符文本置为空
                                try:
                                    paragraph.add_run().add_picture(image_file_path, width=Inches(width), height=Inches(height))
                                except Exception as e:
                                    paragraph.add_run().add_picture(image_file_path)
                                run.text = ""

                            # 如果占位符包含有`:barcode`，替换占位符为条形码
                            elif ":barcode" in run.text:
                                run_text = run.text.replace("{{","").replace("}}","")
                                key_split = run_text.split(":")
                                key = key_split[0]
                                if len(key_split) == 3:
                                    size = key_split[2].split("*")
                                    width = float(size[0])
                                    height = float(size[1])
                                else:
                                    width = None
                                    height = None
                                    
                                # 调用接口生成条形码，默认为模板文件相同路径，并以当前字段的值作为文件名，默认使用`code128`的样式，并重新设定条形码的文件路径
                                barcode_file_name = generate_barcode(value, 'code128', None, main_path)
                                barcode_file_path = os.path.join(main_path, barcode_file_name)

                                # 将条形码替换到条形码占位符所在位置，并把原有的占位符文本置为空
                                try:
                                    paragraph.add_run().add_picture(barcode_file_path, width=Inches(width), height=Inches(height))
                                except Exception as e:
                                    paragraph.add_run().add_picture(barcode_file_path)
                                run.text = "" 

                            # 如果占位符包含有`:qrcode`，替换占位符为二维码
                            elif ":qrcode" in run.text:
                                run_text = run.text.replace("{{","").replace("}}","")
                                key_split = run_text.split(":")
                                key = key_split[0]
                                if len(key_split) == 3:
                                    size = key_split[2].split("*")
                                    width = float(size[0])
                                    height = float(size[1])
                                else:
                                    width = None
                                    height = None
                                    
                                # 调用接口生成二维码，默认为模板文件相同路径，并以当前字段的值作为文件名，默认使用`code128`的样式，并重新设定条形码的文件路径
                                qrcode_file_name = generate_qrcode(value, {}, main_path)
                                qrcode_file_path = os.path.join(main_path, qrcode_file_name)

                                # 将二维码替换到二维码占位符所在位置，并把原有的占位符文本置为空
                                try:
                                    paragraph.add_run().add_picture(qrcode_file_path, width=Inches(width), height=Inches(height))
                                except Exception as e:
                                    paragraph.add_run().add_picture(qrcode_file_path)
                                run.text = ""

                            # 如果不是以上三种情况，则直接替换为对应字段的值
                            else:
                                run.text = run.text.replace("{{"+ key + "}}", value, 1)

                        # 如果替换失败，则将当前文本片断置为空，继续后面的执行
                        except Exception as e:
                            run.text = ""
                            
                        # 应用之前保存的样式
                        if font_size:  # 如果存在字体大小，设置字体大小
                            run.font.size = Pt(font_size.pt)
                        if color:  # 如果存在颜色，则设置颜色
                            run.font.color.rgb = RGBColor(*color)
                        if bold:  # 如果存在粗体，则设置粗体
                            run.font.bold = bold
                        if italic:  # 如果存在斜体，则设置斜体
                            run.font.italic = italic
                        break  # 只考虑第一个出现的占位符
                
                index = index + 1


    ## 遍历文档中的所有表格
    for table in doc.tables:
        # 遍历表格中的每一行
        for row in table.rows:
            # 遍历行中的每一个单元格
            for cell in row.cells:
                # 遍历单元格中的每一个段落
                for paragraph in cell.paragraphs:
                    text = paragraph.text.replace('\n', '').replace('\r', '').replace('\r\n', '').strip()
                    if BaseClass().is_variable(text):
                        key = text.replace("{{", "").replace("}}", "")
                        # paragraph.text = paragraph.text.replace(text, info_json[key])

                        font_size = paragraph.runs[0].font.size  # 假设所有格式相同，这里仅取第一个run的格式
                        color = paragraph.runs[0].font.color.rgb  # 保存颜色，如果有的话
                        bold = paragraph.runs[0].font.bold is not None  # 保存粗体状态
                        italic = paragraph.runs[0].font.italic is not None  # 保存斜体状态
                        
                        # 遍历段落中的每一个run（文本片段）
                        text_tmp = ""
                        for run in paragraph.runs:
                            # print(run.text)
                            text_tmp = text_tmp + run.text
                            if text_tmp == text:
                                # print(text_tmp)
                                # print(key)
                                # print(info_json[key])
                                # 如果run的文本包含占位符，则替换它
                                try:
                                    if  ":image" in text_tmp:
                                        key_split = key.split(":")
                                        key = key_split[0]
                                        if len(key_split) == 3:
                                            size = key_split[2].split("*")
                                            width = float(size[0])
                                            height = float(size[1])
                                        else:
                                            width = None
                                            height = None
                                            
                                        extra = {"bitablePerm":{"tableId":table_id,"attachments":{field_id_map[key]:{record_id:[info_json[key]]}}}}
                                        attachment_resp = BaseClass().download_attachment(personal_token, info_json[key], extra)
                                        # print(attachment_resp)

                                        with open(image_file_path, 'wb') as file:
                                            file.write(attachment_resp.content)
                                        file.close()

                                        try:
                                            paragraph.add_run().add_picture(image_file_path, width=Inches(width), height=Inches(height))
                                            # print(paragraph.text)
                                        except Exception as e:
                                            paragraph.add_run().add_picture(image_file_path)
                                        run.text = ""

                                    elif ":barcode" in text_tmp:
                                        key_split = key.split(":")
                                        key = key_split[0]
                                        if len(key_split) == 3:
                                            size = key_split[2].split("*")
                                            width = float(size[0])
                                            height = float(size[1])
                                        else:
                                            width = None
                                            height = None

                                        barcode_file_name = generate_barcode(info_json[key], 'code128', None, main_path)
                                        barcode_file_path = os.path.join(main_path, barcode_file_name)

                                        try:
                                            paragraph.add_run().add_picture(barcode_file_path, width=Inches(width), height=Inches(height))
                                        except Exception as e:
                                            paragraph.add_run().add_picture(barcode_file_path)
                                        run.text = ""

                                    elif ":qrcode" in text_tmp:
                                        key_split = key.split(":")
                                        key = key_split[0]
                                        if len(key_split) == 3:
                                            size = key_split[2].split("*")
                                            width = float(size[0])
                                            height = float(size[1])
                                        else:
                                            width = None
                                            height = None

                                        qrcode_file_name = generate_qrcode(info_json[key], {}, main_path)
                                        qrcode_file_path = os.path.join(main_path, qrcode_file_name)

                                        try:
                                            paragraph.add_run().add_picture(qrcode_file_path, width=Inches(width), height=Inches(height))
                                        except Exception as e:
                                            paragraph.add_run().add_picture(qrcode_file_path)
                                        run.text = ""
                                        
                                    else:
                                        run.text = info_json[key]
                                except Exception as e:
                                    run.text = ""

                                if font_size:  # 如果存在字体大小，设置字体大小
                                    run.font.size = Pt(font_size.pt)
                                if color:  # 如果存在颜色，则设置颜色
                                    run.font.color.rgb = RGBColor(*color)
                                if bold:  # 如果存在粗体，则设置粗体
                                    run.font.bold = bold
                                if italic:  # 如果存在斜体，则设置斜体
                                    run.font.italic = italic
                                text_tmp = ""
                                break  # 因为我们只处理第一个出现的占位符，所以找到后退出循环
                            else:
                                run.text = ""
                                

    # 保存修改后的文档
    doc.save(target_file_path)

    if file_type == 'pdf':
        #获取文件名称
        filename=target_file_path.split(".docx")[0]
        pdf_target_file_path = f"{filename}.pdf"
        system = platform.system()
        convert_flag = True
        
        # 将 docx 文档转换为 PDF，如果转换失败，将上传 docx 文件到附件字段中
        if system == 'Windows':
            try:
                convert(target_file_path, pdf_target_file_path)

            except Exception as e:
                print("当前系统未安装Office软件，PDF转换失败")
                msg = "当前系统未安装Office软件，PDF转换失败"
                convert_flag = False

        elif system == 'Linux':
            command = [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                pdf_target_file_path,
                target_file_path
            ]
            try:
                subprocess.run(command, check=True)
                
            except subprocess.CalledProcessError as e:
                print("当前系统未安装LibreOffice软件，PDF转换失败")
                msg = "当前系统未安装LibreOffice软件，PDF转换失败"

                try:
                    # 更新包列表并安装LibreOffice
                    subprocess.run(["sudo", "apt", "update"], check=True)
                    subprocess.run(["sudo", "apt", "install", "libreoffice", "-y"], check=True)

                    try:
                        subprocess.run(command, check=True)
                        
                    except subprocess.CalledProcessError as e:
                        print("PDF转换失败")
                        msg = "PDF转换失败"
                        convert_flag = False

                except subprocess.CalledProcessError as e:
                    print("安装LibreOffice软件失败")
                    msg = "安装LibreOffice软件失败"
                    convert_flag = False

                # print(f"Linux 系统转换出错: {e}")

        if convert_flag == True:
            file = (open(pdf_target_file_path, 'rb'))
            req_body = {
                "file_name": file_name + ".pdf",
                "parent_type": "bitable_file",
                "parent_node": app_token,
                "size": str(os.path.getsize(pdf_target_file_path)),
                "file": file
            }
        else:
            file = (open(target_file_path, 'rb'))
            req_body = {
                "file_name": file_name + ".docx",
                "parent_type": "bitable_file",
                "parent_node": app_token,
                "size": str(os.path.getsize(target_file_path)),
                "file": file
            }

    # 如果 file_type 不为 `pdf` 时执行以下代码
    else:
        file = (open(target_file_path, 'rb'))
        req_body = {
            "file_name": file_name + ".docx",
            "parent_type": "bitable_file",
            "parent_node": app_token,
            "size": str(os.path.getsize(target_file_path)),
            "file": file
        }


    multi_form = MultipartEncoder(req_body)
    # 上传附件到多维表格空间
    response = BaseClass().upload_all(personal_token, multi_form)
    # print(response)
    file.close()

    record_list = []
    field_list = {}
    fields = {}
    if response.get("code") == 0:
        file_token = [response.get("data")]
        fields[file_field] = file_token
        field_list["fields"] = fields
        field_list["record_id"] = record_id
        record_list.append(field_list)

        # print(record_list)

        # 更新多维表格记录
        response = BaseClass().batch_update_record(app_token, personal_token, table_id, record_list)
        # print(response)
        if response.get("code") == 0:
            msg = "生成附件成功"

            # 附件更新成功后，将模板目录中的临时文件全部删除
            if os.path.isfile(target_file_path):
                file.close()
                try:
                    os.remove(target_file_path)
                    os.remove(image_file_path)
                    os.remove(barcode_file_path)
                    os.remove(qrcode_file_path)
                    os.remove(pdf_target_file_path)
                except Exception as e:
                    pass

    return msg



## 判断文件名是否在允许的格式范围内
def allowed_file(filename):
    """
        检验文件名尾缀是否满足格式要求\r\n
        :param filename:\r\n
        :return:\r\n
    """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


## 删除指定路径下面的所有文件
def delete_files_in_directory(directory):
    try:
        file_list = os.listdir(directory)
        file_list.remove(".gitkeep")
        for file_name in file_list:
            file_path = os.path.join(directory, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
    except Exception as e:
        return



## 上传模板文件接口
@app.route('/upload_template', methods=['GET', 'POST'])
def upload_file():
    """
    上传文件到 template_files 文件夹下对应的 personal_token 下
    """

    if 'filePicker' not in request.files:
        return "No file part"
    
    # print(request.files)
    file_list = dict(request.files.lists()).get("filePicker")

    # print(file_list)

    personal_token = dict(request.form.lists()).get("personal_token")[0]
    # print(personal_token)
    if personal_token == "":
        return "多维表格授权码不能为空"

    result_msg = ""
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], personal_token)
    # print(file_path)

    if not os.path.exists(file_path):
        os.mkdir(file_path)
    else:
        delete_files_in_directory(file_path)

    for file in file_list:
        # print(file)
        # print(file.filename)
        if file.filename == '':
            return 'No selected file'
        elif file.filename != 'template.docx':
            return '模板文件名必须为 template.docx'
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            try:
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], personal_token, filename))
                result_msg = result_msg + '\'' + filename + '\' file uploaded successfully<br><br>'
                server_url = request.headers.get("Origin")
                identifier = str(uuid.uuid1())
                # print(identifier)
                result_msg = result_msg + server_url + "/generate_attachment?identifier=" + identifier
            except Exception as e:
                # print(e)
                result_msg = result_msg + '\'' + filename + '\' file uploaded Fail<br>'


    return result_msg



## 生成多维表格附件接口
@app.route("/generate_attachment", methods=['POST'])
def generate_attachment():

    result_msg = ""
    result_code = 200

    request_body = json.loads(request.data.decode("utf-8"))
    # print(request_body)

    response = BaseClass().list_fields(request_body.get("app_token"), request_body.get("personal_base_token"), request_body.get("table_id"))
    # print(response)
    field_map = {}
    field_id_map = {}
    if response.get("code") == 0:
        field_items = response.get("data").get("items")
        for item in field_items:
            field_map[item.get("field_name")] = item.get("type")
            field_id_map[item.get("field_name")] = item.get("field_id")


    # print(field_map)
    # print(field_id_map)
    # print("*" * 50)

    response = BaseClass().batch_get_records(request_body.get("app_token"), request_body.get("personal_base_token"), request_body.get("table_id"), [request_body.get("record_id")])
    # print(response)
    field_list = {}
    if response.get("code") == 0:
        records = response.get("data").get("records")[0].get("fields")
        shared_url = response.get("data").get("records")[0].get("shared_url")
        field_list["记录链接"] = shared_url
        # print(records)
        # print("*" * 50)
        for key, item in records.items():
            if key != request_body.get("file_field"):
                field_value = BaseClass().get_field_value(field_map[key], item)
            else:
                field_value = ""

            # print(key, ":", field_value)

            field_list[key] = field_value
        
        try:
            msg = export_to_doc(request_body.get("app_token"), request_body.get("personal_base_token"), request_body.get("table_id"), request_body.get("record_id"), field_list, request_body.get("file_name"), request_body.get("file_field"), field_id_map, request_body.get("file_type", None))
            # result_msg = "生成附件成功"
            result_msg = msg

        except Exception as e:
            result_msg = "生成附件失败，请联系管理员查看日志"

    else:
        result_msg = "获取记录失败，请重试！"
        result_code = -1
    
    return {"msg": result_msg, "code": result_code}


## 删除 template_files 目录下面生成的条形码和二维码文件接口
@app.route("/clean_generate_files")
def clean_generate_files():
    file_path = app.config['GENERATE_FOLDER']
    delete_files_in_directory(file_path)

    timestamp = time.time()
    local_time = time.localtime(timestamp)
    formatted_local_time = time.strftime('%Y-%m-%d %H:%M:%S', local_time)

    print("【{}】template_files 目录下文件删除成功".format(formatted_local_time))
    return {"code": 200, "msg": "文件删除成功"}


## 条形码和二维码下载接口，返回文件的二进制流
@app.route("/download_file")
def download_file():
    """
    下载template_files目录下面的文件
    params:\r\n
    - file_name: 指定要下载的文件名\r\n
    - return_type: 指定返回的类型，不指定此参数默认为文件的二进制信息，可设置为 base64 生成图片的 base64 编码\r\n
    :return:
    """

    # # 读取图片文件并编码为base64
    # with open('path/to/your/image.png', 'rb') as image_file:
    #     encoded_string = base64.b64encode(image_file.read())
    # encoded_str = encoded_string.decode('utf-8')
    # print("Encoded Image:", encoded_str)
    
    # # 解码base64字符串并保存为图片
    # encoded_data = encoded_str.encode('utf-8')
    # decoded_data = base64.b64decode(encoded_data)
    # with open('path/to/save/image_decoded.png', 'wb') as decoded_file:
    #     decoded_file.write(decoded_data)
        

    return_type = request.args.get("return_type", "file")


    file_name = request.args.get('file_name')
    file_path = os.path.join(fp, app.config['GENERATE_FOLDER'], file_name)
    if os.path.isfile(file_path):
        if return_type == 'file':
            return send_file(file_path, as_attachment=True)
        elif return_type == 'base64':
            with open(file_path, 'rb') as image_file:
                encoded_string = base64.b64encode(image_file.read())
            image_file.close()
            encoded_str = encoded_string.decode('utf-8')
            # print("Encoded Image:", encoded_str)
            return {"code": 200, "msg": "下载成功","data": "data:image/png;base64," + encoded_str}
    else:
        return {"code": -1, "msg":"下载的文件不存在，请尝试重新生成"}
    


## 生成条形码接口，返回下载链接
@app.route('/generate_barcode', methods=['POST'])
def barcode():

    # "ean8": EAN8,
    # "ean8-guard": EAN8_GUARD,
    # "ean13": EAN13,
    # "ean13-guard": EAN13_GUARD,
    # "ean": EAN13,
    # "gtin": EAN14,
    # "ean14": EAN14,
    # "jan": JAN,
    # "upc": UPCA,
    # "upca": UPCA,
    # "isbn": ISBN13,
    # "isbn13": ISBN13,
    # "gs1": ISBN13,
    # "isbn10": ISBN10,
    # "issn": ISSN,
    # "code39": Code39,
    # "pzn": PZN,
    # "code128": Code128,
    # "itf": ITF,
    # "gs1_128": Gs1_128,
    # "codabar": CODABAR,
    # "nw-7": CODABAR,
    
    '''
    # POST 请求体参数格式如下：
    {
        "barcode_class": "code128",
        "options": {
            "module_width": 0.3,
            "module_height": 15.0,
            "font_size": 10,
            "text_distance": 5.0,
            "quiet_zone": 10.5
        }
    }
    '''

    request_body = {}

    try:
        request_body = json.loads(request.data.decode("utf-8"))

    except Exception as e:
        return '请求参数格式错误'

    # print(request_body)

    barcode_class = request_body.get("barcode_class", "code128")

    options = request_body.get("options", {})

    content = request.args.get("content", None)
    if content is None:
        content = request_body.get("content", None)
        if content is None:
            return "条码内容为空，请添加查询参数 content"

    file_path = os.path.join(fp, app.config['GENERATE_FOLDER'])

    result = generate_barcode(content, barcode_class, options, file_path)

    server_url = request.headers.get("Host")

    return {"code": 200, "msg":"生成成功","data":'https://' + server_url + '/download_file?file_name=' + result}



## 生成二维码接口，返回下载链接
@app.route('/generate_qrcode', methods=['POST'])
def qrcode():

    '''
    # POST 请求体参数格式如下：
    {
        "version": 2,
        "error_correction": "ERROR_CORRECT_H",
        "box_size": 12,
        "border": 2,
        "fill_color": "green",
        "back_color":"white"
    }
    '''

    request_body = {}
    try:
        request_body = json.loads(request.data.decode("utf-8"))

    except Exception as e:
        return '请求参数格式错误'

    # print(request_body)

    content = request.args.get("content", None)
    if content is None:
        content = request_body.get("content", None)
        if content is None:
            return "二维码内容为空，请添加查询参数 content"

    file_path = os.path.join(fp, app.config['GENERATE_FOLDER'])

    result = generate_qrcode(content, request_body, file_path)

    server_url = request.headers.get("Host")

    return {"code": 200, "msg":"生成成功","data":'https://' + server_url + '/download_file?file_name=' + result}



## 插件主页，用于上传模板文件
@app.route('/', methods=['GET'])
def index():
    identifier = str(uuid.uuid1())
    return render_template("index.html", identifier=identifier)


# 创建一个调度器
scheduler = BackgroundScheduler()
# 启动调度器
scheduler.start()
# 添加定时任务
task = scheduler.add_job(clean_generate_files, 'cron', hour=0, minute=30, id='task')


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=3300, debug=True, use_reloader=True)