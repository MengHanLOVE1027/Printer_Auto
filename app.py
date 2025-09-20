# ================== 导入模块 ==================

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
import os
import uuid
import hashlib
import win32print
import win32com.client
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants
import pythoncom
import win32api
import win32con
from PIL import Image
import threading
from datetime import datetime
from werkzeug.utils import secure_filename
from payment import init_payment_routes, create_payment_order, check_payment_status
from task_storage import get_print_tasks, set_print_tasks, get_task, update_task

# ================== 项目信息 ==================

# 项目基本信息
PROJECT_NAME = "自助打印系统"
VERSION = "v1.0.1"
AUTHOR = "梦涵LOVE"
CONTACT = "QQ 2193438288"

# ================== 配置和常量定义 ==================

# Flask应用配置
app = Flask(__name__)
app.secret_key = 'your_secret_key'  # 用于flash消息

# 文件上传配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = set()  # 空集合表示允许所有文件类型
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB

# 打印机配置
PRINTER_NAME = "HP2C05A1 (HP DeskJet 2800 series)"  # 默认打印机，根据检测到的实际打印机名称修改
HP_PRINTER_IP = "192.168.3.225"  # HP打印机IP地址
HP_PRINTER_NAME = "HP DeskJet 2800e"  # HP打印机名称

# 任务状态跟踪
TASK_STATUS = {
    "pending": "等待中",
    "processing": "处理中",
    "completed": "已完成",
    "failed": "失败",
    "pending_payment": "等待支付",
    "paid": "已支付"
}

# 支付配置
PAYMENT_KEY = "1145141919810"  # 支付密钥
PAYMENT_API_URL = "https://example.web/mapi.php"  # 支付API接口
PAYMENT_SUBMIT_URL = "https://example.web/submit.php"  # 支付提交接口
ENABLE_PAYMENT = True  # 是否启用支付功能，True为启用，False为禁用

# 打印价格配置（每页价格）
PRICE_PER_PAGE_MONOCHROME = 0.1  # 黑白打印每页价格
PRICE_PER_PAGE_COLOR = 0.3  # 彩色打印每页价格

def generate_sign(params, key):
    """生成支付签名"""
    # 1. 过滤空值和签名参数
    filtered_params = {}
    for k, v in params.items():
        if v != '' and v is not None and k != 'sign' and k != 'sign_type':
            filtered_params[k] = v

    # 2. 按照参数名ASCII码从小到大排序
    sorted_params = sorted(filtered_params.items(), key=lambda x: x[0])

    # 3. 拼接成URL键值对
    stringA = '&'.join([f"{k}={v}" for k, v in sorted_params])

    # 4. 拼接商户密钥并进行MD5加密
    stringSignTemp = stringA + key
    sign = hashlib.md5(stringSignTemp.encode('utf-8')).hexdigest()

    return sign

# 初始化任务存储
from task_storage import get_print_tasks, set_print_tasks
print_tasks = get_print_tasks()

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """检查文件类型是否允许 - 现在允许所有文件类型"""
    # 如果ALLOWED_EXTENSIONS为空集合，则允许所有文件类型
    if not ALLOWED_EXTENSIONS:
        return True
    # 否则检查文件扩展名是否在允许列表中
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    # 获取系统打印机列表
    printers = get_printers()
    return render_template('index.html', printers=printers)

@app.route('/api/printers')
def get_printers_api():
    """获取系统打印机列表的API接口"""
    try:
        printers = get_printers()
        return jsonify({'success': True, 'printers': printers})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/task/<task_id>')
def task_status(task_id):
    """显示任务状态页面"""
    task = print_tasks.get(task_id)
    if not task:
        flash('任务不存在或已过期')
        return redirect(url_for('index'))

    return render_template('task_status.html', task=task)

@app.route('/api/task/<task_id>')
def get_task_status_api(task_id):
    """获取任务状态的API接口"""
    task = print_tasks.get(task_id)
    if not task:
        return jsonify({'success': False, 'error': '任务不存在或已过期'})

    return jsonify({'success': True, 'task': task})

@app.route('/tasks')
def list_tasks():
    """显示所有任务列表"""
    # 按创建时间倒序排序
    sorted_tasks = sorted(print_tasks.values(),
                         key=lambda x: x['created_at'],
                         reverse=True)
    return render_template('tasks.html', tasks=sorted_tasks)

def get_printers():
    """获取系统安装的打印机列表并检测其状态"""
    printers = []
    # 定义要过滤掉的虚拟打印机名称列表
    filtered_printers = ["导出为WPS PDF", "Microsoft Print to PDF"]

    try:
        # 使用win32print获取本地打印机列表
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        local_printers = win32print.EnumPrinters(flags, None, 2)
        print(f"找到 {len(local_printers)} 台本地打印机")

        for p in local_printers:
            printer_name = p['pPrinterName']

            # 跳过虚拟打印机
            if any(virtual_printer in printer_name for virtual_printer in filtered_printers):
                print(f"跳过虚拟打印机: {printer_name}")
                continue

            printer_comment = p['pComment'] if 'pComment' in p and p['pComment'] else printer_name

            # 简化状态检测 - 只检查是否可以打开打印机
            try:
                hprinter = win32print.OpenPrinter(printer_name)
                win32print.ClosePrinter(hprinter)
                status = "就绪"
            except:
                status = "未知"

            printers.append({
                'name': printer_name,
                'description': printer_comment,
                'status': status,
                'type': 'local'
            })

        # 尝试添加网络打印机 - 使用简单方法
        try:
            # 检查是否已经有HP打印机
            hp_printer_found = any('HP DeskJet 2800' in p['name'] or 'HP2C05A1' in p['name'] for p in printers)

            if not hp_printer_found:
                # 尝试简单的网络打印机名称
                network_printer_name = f"HP DeskJet 2800e"

                # 检查是否已经在列表中
                if not any(p['name'] == network_printer_name for p in printers):
                    # 简化状态检测
                    try:
                        hprinter = win32print.OpenPrinter(network_printer_name)
                        win32print.ClosePrinter(hprinter)
                        status = "就绪"

                        printers.append({
                            'name': network_printer_name,
                            'description': f"{HP_PRINTER_NAME} (IP: {HP_PRINTER_IP})",
                            'status': status,
                            'type': 'network'
                        })
                        print(f"添加网络打印机: {network_printer_name} (状态: {status})")
                    except:
                        print(f"无法连接网络打印机: {network_printer_name}")
        except Exception as e:
            print(f"添加网络打印机失败: {str(e)}")

        # 如果没有找到打印机，尝试获取默认打印机
        if not printers:
            try:
                default_printer = win32print.GetDefaultPrinter()
                # 简化状态检测
                try:
                    hprinter = win32print.OpenPrinter(default_printer)
                    win32print.ClosePrinter(hprinter)
                    status = "就绪"
                except:
                    status = "未知"

                printers.append({
                    'name': default_printer,
                    'description': default_printer,
                    'status': status,
                    'type': 'default'
                })
                print(f"使用默认打印机: {default_printer}")
            except Exception as e:
                print(f"获取默认打印机失败: {str(e)}")
                printers.append({
                    'name': PRINTER_NAME,
                    'description': PRINTER_NAME,
                    'status': '未知',
                    'type': 'fallback'
                })
                print(f"使用预设默认打印机: {PRINTER_NAME}")

        # 打印所有找到的打印机及其状态
        print("找到的打印机列表:")
        for p in printers:
            print(f" - {p['name']}: {p['description']} (状态: {p['status']}, 类型: {p['type']})")

    except Exception as e:
        print(f"获取打印机失败: {str(e)}")
        # 出错时返回默认打印机
        printers.append({
            'name': PRINTER_NAME,
            'description': PRINTER_NAME,
            'status': '未知',
            'type': 'fallback'
        })

    return printers

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('没有选择文件')
        return redirect(request.url)

    file = request.files['file']

    if file.filename == '':
        flash('没有选择文件')
        return redirect(request.url)

    if file and allowed_file(file.filename):
        # 生成唯一任务ID和文件名
        task_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)

        # 获取打印选项
        copies = request.form.get('copies', 1, type=int)
        color = request.form.get('color', 'false') == 'true'
        duplex = request.form.get('duplex', 'false') == 'true'
        printer_name = request.form.get('printer', PRINTER_NAME)
        paper_size = request.form.get('paper-size', 'A4')

        # 计算打印费用（使用配置的价格常量）
        pages = 1  # 实际应用中应该解析文件获取页数
        cost_per_page = PRICE_PER_PAGE_COLOR if color else PRICE_PER_PAGE_MONOCHROME
        total_cost = round(pages * copies * cost_per_page, 2)

        # 根据是否启用支付功能创建任务记录
        if ENABLE_PAYMENT:
            # 启用支付功能，初始状态为等待支付
            status = "pending_payment"
            message = f"任务已创建，请支付{total_cost}元"
        else:
            # 不启用支付功能，直接进入处理中状态
            status = "processing"
            message = "任务已创建，正在准备打印"

        # 创建任务记录
        print_tasks[task_id] = {
            "id": task_id,
            "filename": filename,
            "file_path": file_path,
            "copies": copies,
            "color": color,
            "duplex": duplex,
            "printer_name": printer_name,
            "paper_size": paper_size,
            "cost": total_cost,
            "status": status,
            "progress": 0,
            "message": message,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        # 根据是否启用支付功能决定后续流程
        if ENABLE_PAYMENT:
            # 启用支付功能，重定向到支付页面
            # 实际打印将在支付成功后在execute_print函数中执行
            return redirect(url_for('payment_page', task_id=task_id))
        else:
            # 不启用支付功能，直接执行打印任务
            # 在后台线程中执行打印任务，避免阻塞HTTP请求
            def print_task_background():
                try:
                    # 更新任务状态
                    print_tasks[task_id]["status"] = "processing"
                    print_tasks[task_id]["message"] = "正在准备打印任务"
                    print_tasks[task_id]["progress"] = 10
                    print_tasks[task_id]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # 更新任务状态
                    print_tasks[task_id]["message"] = "正在发送文件到打印机"
                    print_tasks[task_id]["progress"] = 30
                    print_tasks[task_id]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # 打印文件
                    print_file(
                        print_tasks[task_id]["file_path"],
                        print_tasks[task_id]["copies"],
                        print_tasks[task_id]["color"],
                        print_tasks[task_id]["duplex"],
                        print_tasks[task_id]["printer_name"],
                        print_tasks[task_id]["paper_size"]
                    )

                    # 更新任务状态为完成
                    print_tasks[task_id]["status"] = "completed"
                    print_tasks[task_id]["message"] = "打印任务已完成"
                    print_tasks[task_id]["progress"] = 100
                    print_tasks[task_id]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                except Exception as e:
                    # 更新任务状态为失败
                    print_tasks[task_id]["status"] = "failed"
                    print_tasks[task_id]["message"] = f"打印失败: {str(e)}"
                    print_tasks[task_id]["progress"] = 0
                    print_tasks[task_id]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    print(f"打印任务失败: {str(e)}")

            # 启动后台线程
            thread = threading.Thread(target=print_task_background)
            thread.daemon = True
            thread.start()

            # 重定向到任务状态页面
            return redirect(url_for('task_status', task_id=task_id))
    else:
        flash('不支持的文件类型')
        return redirect(request.url)

def print_file(file_path, copies=1, color=False, duplex=False, printer_name=None, paper_size='A4'):
    """使用win32com.client方式统一打印文件"""
    if printer_name is None:
        printer_name = PRINTER_NAME

    print(f"尝试使用打印机: {printer_name}")

    # 获取所有可用打印机及其状态
    available_printers = get_printers()
    print("系统可用打印机及其状态:")
    for p in available_printers:
        print(f"  - {p['name']}: {p['status']}")

    # 检查指定打印机是否在可用列表中
    selected_printer = None
    for p in available_printers:
        if p['name'] == printer_name:
            selected_printer = p
            break

    # 如果指定打印机不可用，尝试找到状态为"就绪"的打印机
    if not selected_printer or selected_printer['status'] != "就绪":
        print(f"打印机 '{printer_name}' 不可用或状态不为就绪")

        # 查找状态为"就绪"的打印机
        for p in available_printers:
            if p['status'] == "就绪":
                selected_printer = p
                printer_name = p['name']
                print(f"切换到就绪的打印机: {printer_name}")
                break

        # 如果没有就绪的打印机，尝试使用默认打印机
        if not selected_printer:
            try:
                default_printer = win32print.GetDefaultPrinter()
                print(f"尝试使用默认打印机: {default_printer}")

                # 检查默认打印机状态
                for p in available_printers:
                    if p['name'] == default_printer:
                        selected_printer = p
                        printer_name = default_printer
                        print(f"使用默认打印机: {default_printer} (状态: {p['status']})")
                        break

                if not selected_printer:
                    raise Exception("默认打印机不在可用列表中")

            except Exception as e:
                print(f"无法使用默认打印机: {e}")

                # 最后尝试使用列表中的第一台打印机
                if available_printers:
                    selected_printer = available_printers[0]
                    printer_name = selected_printer['name']
                    print(f"使用列表中的第一台打印机: {printer_name} (状态: {selected_printer['status']})")
                else:
                    raise Exception("没有可用的打印机")

    try:
        # 尝试打开选择的打印机
        print(f"尝试打开打印机: {printer_name}")
        hprinter = win32print.OpenPrinter(printer_name)
        print(f"成功打开打印机: {printer_name} (状态: {selected_printer['status']})")
        win32print.ClosePrinter(hprinter)

        # 获取文件扩展名
        file_ext = os.path.splitext(file_path)[1].lower()
        abs_file_path = os.path.abspath(file_path)

        print(f"开始打印文件: {file_path}")
        print(f"使用打印机: {printer_name}")
        print(f"打印份数: {copies}")
        print(f"打印模式: {'彩色' if color else '黑白'}")
        print(f"双面打印: {'是' if duplex else '否'}")
        print(f"纸张大小: {paper_size}")

        # 初始化COM库
        pythoncom.CoInitialize()

        # 根据文件类型选择不同的应用程序进行打印
        try:
            if file_ext in ('.doc', '.docx'):
                # 使用Word打印文档
                try:
                    # 设置打印机属性
                    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
                    pHandle = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
                    properties = win32print.GetPrinter(pHandle, 2)

                    # 设置打印属性
                    # 使用Windows API常量来设置颜色模式
                    # DM_COLOR = 32
                    # DMCOLOR_COLOR = 1
                    # DMCOLOR_MONOCHROME = 2
                    properties['pDevMode'].Color = win32con.DMCOLOR_MONOCHROME if not color else win32con.DMCOLOR_COLOR
                    properties['pDevMode'].Copies = copies
                    properties['pDevMode'].Orientation = win32con.DMORIENT_PORTRAIT  # 纵向
                    
                    # 设置颜色标志位
                    properties['pDevMode'].Fields = properties['pDevMode'].Fields | win32con.DM_COLOR
                    
                    # 尝试设置其他可能的颜色控制属性
                    if hasattr(properties['pDevMode'], 'PrintQuality'):
                        # 尝试设置打印质量，可能会影响颜色
                        properties['pDevMode'].PrintQuality = -1  # 默认质量
                    
                    print(f"设置打印机颜色模式: {'黑白' if not color else '彩色'} (值={properties['pDevMode'].Color})")

                    # 设置双面打印
                    if duplex:
                        properties['pDevMode'].Duplex = win32con.DMDUP_VERTICAL
                    else:
                        properties['pDevMode'].Duplex = win32con.DMDUP_SIMPLEX

                    # 应用打印机设置
                    win32print.SetPrinter(pHandle, 2, properties, 0)
                    win32print.ClosePrinter(pHandle)
                    print(f"打印机设置已应用: 彩色={'是' if color else '否'}, 份数={copies}, 双面={'是' if duplex else '否'}")

                    # 使用Word打印文档
                    word = EnsureDispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(abs_file_path)
                    word.ActivePrinter = printer_name

                    # 设置打印选项
                    # 设置打印颜色
                    word.Options.PrintBackground = True
                    word.Options.PrintDrawingObjects = True
                    
                    # 设置打印份数
                    for _ in range(copies):
                        doc.PrintOut()

                    doc.Close(constants.wdDoNotSaveChanges)
                    word.Quit()
                    print("Word文档已发送到打印机")
                except Exception as e:
                    print(f"Word文档打印失败: {e}")
                    raise

            elif file_ext == '.pdf':
                # 使用Adobe Acrobat或默认PDF阅读器打印
                try:
                    # 设置打印机属性
                    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
                    pHandle = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
                    properties = win32print.GetPrinter(pHandle, 2)

                    # 设置打印属性
                    # 使用Windows API常量来设置颜色模式
                    # DM_COLOR = 32
                    # DMCOLOR_COLOR = 1
                    # DMCOLOR_MONOCHROME = 2
                    properties['pDevMode'].Color = win32con.DMCOLOR_MONOCHROME if not color else win32con.DMCOLOR_COLOR
                    properties['pDevMode'].Copies = copies
                    properties['pDevMode'].Orientation = win32con.DMORIENT_PORTRAIT  # 纵向
                    
                    # 设置颜色标志位
                    properties['pDevMode'].Fields = properties['pDevMode'].Fields | win32con.DM_COLOR
                    
                    # 尝试设置其他可能的颜色控制属性
                    if hasattr(properties['pDevMode'], 'PrintQuality'):
                        # 尝试设置打印质量，可能会影响颜色
                        properties['pDevMode'].PrintQuality = -1  # 默认质量
                    
                    print(f"设置打印机颜色模式: {'黑白' if not color else '彩色'} (值={properties['pDevMode'].Color})")

                    # 设置双面打印
                    if duplex:
                        properties['pDevMode'].Duplex = win32con.DMDUP_VERTICAL
                    else:
                        properties['pDevMode'].Duplex = win32con.DMDUP_SIMPLEX

                    # 应用打印机设置
                    win32print.SetPrinter(pHandle, 2, properties, 0)
                    win32print.ClosePrinter(pHandle)
                    print(f"打印机设置已应用: 彩色={'是' if color else '否'}, 份数={copies}, 双面={'是' if duplex else '否'}")

                    # 尝试使用Adobe Acrobat
                    adobe = EnsureDispatch("AcroExch.App")
                    avDoc = EnsureDispatch("AcroExch.AVDoc")

                    if avDoc.Open(abs_file_path, ""):
                        pdDoc = avDoc.GetPDDoc()
                        # 获取打印设置对象
                        pddoc = avDoc.GetPDDoc()
                        
                        # 设置打印参数
                        # 0 = bUI, 1 = nFrom, 2 = nTo, 3 = nPSLevel, 4 = bSilent, 5 = bShrinkToFit, 6 = bPrintAsImage, 7 = bReverse, 8 = bAnnotations
                        # 设置为静默打印，不显示打印对话框
                        bSilent = 1
                        # 设置打印份数
                        for i in range(copies):
                            avDoc.PrintPagesSilent(0, pdDoc.GetNumPages() - 1, 1, bSilent, True, printer_name)
                        avDoc.Close(True)
                        print("PDF文档已发送到打印机")
                    else:
                        raise Exception("无法打开PDF文档")
                except Exception as e:
                    print(f"使用Adobe Acrobat失败: {e}")
                    # 使用默认程序打开PDF
                    os.startfile(abs_file_path, "print")
                    print("已使用默认程序打开PDF进行打印")

            elif file_ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'):
                # 处理图片文件
                try:
                    # 打开图片并检查分辨率
                    im = Image.open(abs_file_path)
                    print(f"图片分辨率: {im.size[0]} x {im.size[1]}")

                    # 设置打印机属性
                    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
                    pHandle = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
                    properties = win32print.GetPrinter(pHandle, 2)

                    # 设置打印属性
                    # 使用Windows API常量来设置颜色模式
                    # DM_COLOR = 32
                    # DMCOLOR_COLOR = 1
                    # DMCOLOR_MONOCHROME = 2
                    properties['pDevMode'].Color = win32con.DMCOLOR_MONOCHROME if not color else win32con.DMCOLOR_COLOR
                    properties['pDevMode'].Copies = copies
                    properties['pDevMode'].Orientation = win32con.DMORIENT_PORTRAIT  # 纵向
                    
                    # 设置颜色标志位
                    properties['pDevMode'].Fields = properties['pDevMode'].Fields | win32con.DM_COLOR
                    
                    # 尝试设置其他可能的颜色控制属性
                    if hasattr(properties['pDevMode'], 'PrintQuality'):
                        # 尝试设置打印质量，可能会影响颜色
                        properties['pDevMode'].PrintQuality = -1  # 默认质量
                    
                    print(f"设置打印机颜色模式: {'黑白' if not color else '彩色'} (值={properties['pDevMode'].Color})")

                    # 设置双面打印
                    if duplex:
                        properties['pDevMode'].Duplex = win32con.DMDUP_VERTICAL
                    else:
                        properties['pDevMode'].Duplex = win32con.DMDUP_SIMPLEX

                    # 应用打印机设置
                    win32print.SetPrinter(pHandle, 2, properties, 0)
                    print(f"打印机设置已应用: 彩色={'是' if color else '否'}, 份数={copies}, 双面={'是' if duplex else '否'}")

                    # 使用ShellExecute打印图片
                    win32api.ShellExecute(0, "print", abs_file_path, None, ".", 0)
                    print("图片已发送到打印机")

                    # 关闭打印机句柄
                    win32print.ClosePrinter(pHandle)
                except Exception as e:
                    print(f"图片处理和打印失败: {e}")
                    # 尝试使用printto作为备选方案
                    try:
                        win32api.ShellExecute(
                            0,  # 父窗口句柄
                            "printto",  # 操作
                            abs_file_path,  # 文件路径
                            f'"{printer_name}"',  # 参数
                            ".",  # 当前目录
                            0  # 显示方式
                        )
                        print("已使用ShellExecute printto发送图片到打印机")
                    except Exception as e2:
                        print(f"ShellExecute printto也失败: {e2}")
                        # 最后尝试使用默认程序打开图片
                        try:
                            os.startfile(abs_file_path, "print")
                            print("已使用默认程序打开图片进行打印")
                        except Exception as e3:
                            print(f"所有打印方法都失败: {e3}")
                            raise Exception(f"无法打印图片: {e}, {e2}, {e3}")

            elif file_ext == '.txt':
                # 使用默认程序打印文本文件
                try:
                    os.startfile(abs_file_path, "print")
                    print("已使用默认程序打开文本进行打印")
                except Exception as e:
                    print(f"使用默认程序打印失败: {e}")
                    raise Exception(f"无法打印文本: {e}")
                # 使用记事本打印文本文件
                try:
                    notepad = EnsureDispatch("Notepad.Application")
                    notepad.Visible = False
                    notepad.Documents.Open(abs_file_path)
                    notepad.ActivePrinter = printer_name

                    # 设置打印选项并打印
                    for _ in range(copies):
                        notepad.PrintOut()

                    notepad.Documents.Close(constants.wdDoNotSaveChanges)
                    notepad.Quit()
                    print("文本文件已发送到打印机")
                except Exception as e:
                    print(f"使用记事本失败: {e}")
                    # 尝试使用printto作为备选方案
                    try:
                        win32api.ShellExecute(
                            0,  # 父窗口句柄
                            "printto",  # 操作
                            abs_file_path,  # 文件路径
                            f'"{printer_name}"',  # 参数
                            ".",  # 当前目录
                            0  # 显示方式
                        )
                        print("已使用ShellExecute printto发送文本到打印机")
                    except Exception as e2:
                        print(f"ShellExecute printto也失败: {e2}")
                        # 最后尝试使用默认程序打开文本
                        try:
                            os.startfile(abs_file_path, "print")
                            print("已使用默认程序打开文本进行打印")
                        except Exception as e3:
                            print(f"所有打印方法都失败: {e3}")
                            raise Exception(f"无法打印文本: {e}, {e2}, {e3}")
            else:
                # 其他类型文件使用默认程序打印
                print(f"使用默认程序打印{file_ext}文件")
                try:
                    os.startfile(abs_file_path, "print")
                    print("已使用默认程序打开文件进行打印")
                except Exception as e:
                    print(f"使用默认程序打印失败: {e}")
                    # 尝试使用ShellExecute作为备选方案
                    try:
                        win32api.ShellExecute(
                            0,  # 父窗口句柄
                            "printto",  # 操作
                            abs_file_path,  # 文件路径
                            f'"{printer_name}"',  # 参数
                            ".",  # 当前目录
                            0  # 显示方式
                        )
                        print("已使用ShellExecute发送文件到打印机")
                    except Exception as e2:
                        print(f"ShellExecute打印也失败: {e2}")
                        raise Exception(f"无法打印文件: {e}, {e2}")

            print("打印任务已成功提交")

        except Exception as e:
            print(f"打印错误详情: {e}")
            raise Exception(f"打印过程中出错: {str(e)}")
        finally:
            # 释放COM资源
            pythoncom.CoUninitialize()

    except Exception as e:
        print(f"打印机错误详情: {e}")
        raise Exception(f"打印机连接出错: {str(e)}")

# 添加支付相关路由
@app.route('/payment/<task_id>')
def payment_page(task_id):
    """显示支付页面"""
    task = print_tasks.get(task_id)
    if not task:
        flash('任务不存在或已过期')
        return redirect(url_for('index'))

    # 如果任务已经支付，直接跳转到任务状态页面
    if task.get("status") != "pending_payment":
        return redirect(url_for('task_status', task_id=task_id))

    return render_template('payment.html', task=task)

@app.route('/process_payment', methods=['POST'])
def process_payment():
    """处理支付请求"""
    task_id = request.form.get('task_id')
    payment_type = request.form.get('payment_type', 'form')  # form 或 api
    task = print_tasks.get(task_id)

    if not task:
        return jsonify({'success': False, 'error': '任务不存在或已过期'})

    # 创建支付订单
    try:
        # 获取支付方式，默认为支付宝
        payment_method = request.form.get('payment_method', 'alipay')

        params = create_payment_order(
            money=task['cost'],
            name=f"打印服务-{task['filename']}",
            out_trade_no=f"PRINT_{task_id}",
            param=task_id,
            payment_type=payment_method
        )

        # 根据支付类型返回不同的参数
        if payment_type == 'api':
            # API接口支付
            return jsonify({
                'success': True,
                'params': params,
                'api_url': PAYMENT_API_URL,
                'payment_type': 'api'
            })
        else:
            # 页面跳转支付
            return jsonify({
                'success': True,
                'params': params,
                'submit_url': PAYMENT_SUBMIT_URL,
                'payment_type': 'form'
            })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/execute_print/<task_id>')
def execute_print(task_id):
    """执行打印任务（支付成功后调用）"""
    task = print_tasks.get(task_id)
    if not task:
        flash('任务不存在或已过期')
        return redirect(url_for('index'))

    # 更新任务状态为处理中
    task["status"] = "processing"
    task["message"] = "支付成功，正在准备打印任务"
    task["progress"] = 10
    task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 在后台线程中执行打印任务，避免阻塞HTTP请求
    def print_task_background():
        try:
            # 更新任务状态
            task["status"] = "processing"
            task["message"] = "正在准备打印任务"
            task["progress"] = 10
            task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 更新任务状态
            task["message"] = "正在发送文件到打印机"
            task["progress"] = 30
            task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # 打印文件
            print_file(
                task["file_path"],
                task["copies"],
                task["color"],
                task["duplex"],
                task["printer_name"],
                task["paper_size"]
            )

            # 更新任务状态为完成
            task["status"] = "completed"
            task["message"] = "打印任务已完成"
            task["progress"] = 100
            task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception as e:
            # 更新任务状态为失败
            task["status"] = "failed"
            task["message"] = f"打印失败: {str(e)}"
            task["progress"] = 0
            task["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"打印任务失败: {str(e)}")

    # 启动后台线程
    thread = threading.Thread(target=print_task_background)
    thread.daemon = True
    thread.start()

    # 重定向到任务状态页面
    return redirect(url_for('task_status', task_id=task_id))

# 添加易支付平台接口路由
@app.route('/submit.php', methods=['POST', 'GET'])
def payment_submit():
    """支付提交接口"""
    # 获取支付参数
    params = request.form.to_dict() if request.method == 'POST' else request.args.to_dict()

    # 验证必要参数
    required_params = ['pid', 'out_trade_no', 'name', 'money', 'notify_url', 'return_url']
    for param in required_params:
        if param not in params or not params[param]:
            return f'缺少必要参数: {param}', 400

    # 验证签名
    received_sign = params.get('sign')
    calculated_sign = generate_sign(params, PAYMENT_KEY)

    if received_sign != calculated_sign:
        return '签名验证失败', 400

    # 返回表单页面，让用户跳转到支付页面
    return render_template('payment_redirect.html', params=params)

@app.route('/mapi.php', methods=['POST'])
def payment_api():
    """API支付接口"""
    try:
        # 获取支付参数
        params = request.form.to_dict()

        # 验证必要参数
        required_params = ['pid', 'type', 'out_trade_no', 'name', 'money', 'notify_url']
        for param in required_params:
            if param not in params or not params[param]:
                return jsonify({'code': 0, 'msg': f'缺少必要参数: {param}'}), 400

        # 验证签名
        received_sign = params.get('sign')
        calculated_sign = generate_sign(params, PAYMENT_KEY)

        if received_sign != calculated_sign:
            return jsonify({'code': 0, 'msg': '签名验证失败'}), 400

        # 生成支付二维码或跳转URL
        # 根据支付方式生成不同的支付链接
        payment_type = params.get('type', 'alipay')
        trade_no = f"EASY_PAY_{uuid.uuid4().hex[:12]}"

        # 模拟生成支付链接
        if payment_type == 'alipay':
            payurl = f"https://example.web/pay/alipay/{trade_no}"
            qrcode = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data={payurl}"
        elif payment_type == 'wxpay':
            payurl = f"https://example.web/pay/wxpay/{trade_no}"
            qrcode = f"weixin://wxpay/bizpayurl?pr=EASY_PAY_{uuid.uuid4().hex[:8]}"
        elif payment_type == 'unionpay':
            payurl = f"https://example.web/pay/unionpay/{trade_no}"
            qrcode = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data={payurl}"
        else:
            payurl = f"https://example.web/pay/{payment_type}/{trade_no}"
            qrcode = f"https://api.qrserver.com/v1/create-qr-code/?size=200x200&data={payurl}"

        # 根据支付方式返回不同的参数
        if payment_type == 'wxpay':
            # 微信支付返回二维码链接
            return jsonify({
                'code': 1,
                'msg': '支付请求成功',
                'trade_no': trade_no,
                'qrcode': qrcode,
                'money': params['money']
            })
        else:
            # 其他支付方式返回跳转URL
            return jsonify({
                'code': 1,
                'msg': '支付请求成功',
                'trade_no': trade_no,
                'payurl': payurl,
                'money': params['money']
            })
    except Exception as e:
        return jsonify({'code': 0, 'msg': f'处理异常: {str(e)}'}), 500

@app.route('/api.php', methods=['GET'])
def payment_query():
    """查询接口"""
    try:
        # 获取查询参数
        act = request.args.get('act')
        pid = request.args.get('pid')
        key = request.args.get('key')
        out_trade_no = request.args.get('out_trade_no')

        if not act or act != 'order':
            return jsonify({'code': 0, 'msg': '操作类型错误'}), 400

        if not pid or not key:
            return jsonify({'code': 0, 'msg': '缺少商户ID或密钥'}), 400

        # 查询订单状态
        task_id = out_trade_no.replace('PRINT_', '')
        task = print_tasks.get(task_id)

        if not task:
            return jsonify({'code': 0, 'msg': '订单不存在'}), 400

        # 返回订单信息
        return jsonify({
            'code': 1,
            'msg': '查询订单号成功！',
            'trade_no': f"EASY_PAY_{uuid.uuid4().hex[:12]}",
            'out_trade_no': out_trade_no,
            'type': 'alipay',
            'pid': pid,
            'addtime': task.get('created_at', ''),
            'endtime': task.get('updated_at', ''),
            'name': f"打印服务-{task['filename']}",
            'money': str(task['cost']),
            'status': 1 if task.get('status') == 'paid' else 0,
            'param': task_id,
            'buyer': ''
        })
    except Exception as e:
        return jsonify({'code': 0, 'msg': f'查询异常: {str(e)}'}), 500

@app.route('/payment/redirect/<out_trade_no>')
def payment_redirect(out_trade_no):
    """支付跳转页面"""
    return render_template('payment_redirect_page.html', out_trade_no=out_trade_no)

# 初始化支付路由
init_payment_routes(app)

if __name__ == '__main__':
    # 输出项目信息到日志
    print(f"========================================")
    print(f"项目名称: {PROJECT_NAME}")
    print(f"版本: {VERSION}")
    print(f"作者: {AUTHOR}")
    print(f"联系方式: {CONTACT}")
    print(f"========================================")
    print(f"启动自助打印系统...")

    app.run(debug=True, host='0.0.0.0', port=5000)
