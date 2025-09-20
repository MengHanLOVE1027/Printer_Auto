import hashlib
import requests
import uuid
import json
from datetime import datetime
from flask import Flask, request, redirect, url_for, render_template, flash, jsonify
from task_storage import get_print_tasks, update_task

# ================== 支付配置 ==================

# 支付配置
PAYMENT_CONFIG = {
    'pid': '1000',  # 商户ID
    'key': '1145141919810',  # 商户密钥
    'notify_url': 'http://127.0.0.1:5000/payment/notify',  # 异步通知地址（不用改）
    'return_url': 'http://127.0.0.1:5000/payment/return',  # 跳转通知地址（不用改）
    'submit_url': 'https://example.web/submit.php',  # 支付提交地址
    'api_url': 'https://example.web/mapi.php',  # API支付地址
    'query_url': 'https://example.web/api.php'  # 查询地址
}

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

def create_payment_order(money, name, out_trade_no=None, param=None, payment_type='alipay'):
    """创建支付订单"""
    if not out_trade_no:
        out_trade_no = f"PRINT_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}"

    params = {
        'pid': PAYMENT_CONFIG['pid'],
        'type': payment_type,  # 支付方式: alipay, wxpay, qqpay等
        'out_trade_no': out_trade_no,
        'notify_url': PAYMENT_CONFIG['notify_url'],
        'return_url': PAYMENT_CONFIG['return_url'],
        'name': name,
        'money': str(money),
        'clientip': request.remote_addr if request else '127.0.0.1',
        'device': 'pc',
        'param': param or '',
        'sitename': '自助打印系统'
    }

    # 生成签名
    sign = generate_sign(params, PAYMENT_CONFIG['key'])
    params['sign'] = sign
    params['sign_type'] = 'MD5'

    return params

def check_payment_status(out_trade_no):
    """查询支付状态"""
    params = {
        'act': 'order',
        'pid': PAYMENT_CONFIG['pid'],
        'key': PAYMENT_CONFIG['key'],
        'out_trade_no': out_trade_no
    }

    try:
        response = requests.get(f"{PAYMENT_CONFIG['query_url']}", params=params)
        result = response.json()

        if result.get('code') == 1:
            return result.get('status') == 1  # 1表示支付成功
        return False
    except Exception as e:
        print(f"查询支付状态异常: {str(e)}")
        return False

def init_payment_routes(app):
    """初始化支付路由"""

    @app.route('/payment/create', methods=['POST'])
    def create_payment():
        """创建支付订单"""
        try:
            money = request.form.get('money', type=float)
            name = request.form.get('name', '打印服务')
            task_id = request.form.get('task_id')

            if not money or money <= 0:
                return jsonify({'success': False, 'error': '金额必须大于0'})

            # 获取支付方式，默认为支付宝
            payment_method = request.form.get('payment_method', 'alipay')
            
            # 创建支付订单
            params = create_payment_order(
                money=money,
                name=name,
                param=task_id,
                payment_type=payment_method
            )

            # 返回支付参数，前端将使用这些参数构建支付表单
            return jsonify({
                'success': True,
                'params': params,
                'submit_url': PAYMENT_CONFIG['submit_url']
            })
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)})

    @app.route('/payment/notify', methods=['GET'])
    def payment_notify():
        """支付异步通知"""
        try:
            # 获取支付平台返回的参数
            params = request.args.to_dict()

            # 验证必要参数
            required_params = ['pid', 'trade_no', 'out_trade_no', 'type', 'name', 'money', 'trade_status', 'sign', 'sign_type']
            for param in required_params:
                if param not in params or not params[param]:
                    print(f"缺少必要参数: {param}")
                    return '缺少必要参数', 400

            # 验证签名
            received_sign = params.get('sign')
            calculated_sign = generate_sign(params, PAYMENT_CONFIG['key'])

            if received_sign != calculated_sign:
                print(f"签名验证失败: received={received_sign}, calculated={calculated_sign}")
                return '签名验证失败', 400

            # 检查支付状态
            if params.get('trade_status') == 'TRADE_SUCCESS':
                # 支付成功，处理业务逻辑
                out_trade_no = params.get('out_trade_no')
                task_id = params.get('param')

                # 更新打印任务状态为已支付
                if task_id:
                    updates = {
                        "status": "paid",
                        "message": "支付成功，等待打印",
                        "progress": 50,
                        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    if update_task(task_id, updates):
                        print(f"支付成功: 订单号={out_trade_no}, 任务ID={task_id}")
                        
                        # 自动执行打印任务
                        try:
                            # 使用requests库调用execute_print函数
                            import requests
                            server_url = request.host_url
                            response = requests.get(f"{server_url}execute_print/{task_id}")
                            if response.status_code != 200:
                                print(f"执行打印任务失败: HTTP状态码={response.status_code}")
                        except Exception as e:
                            print(f"执行打印任务异常: {str(e)}")

                # 返回success表示接收到了通知
                return 'success'

            print(f"支付状态异常: {params.get('trade_status')}")
            return '支付状态异常', 400
        except Exception as e:
            print(f"支付通知处理异常: {str(e)}")
            return '处理异常', 500

    @app.route('/payment/return')
    def payment_return():
        """支付跳转通知"""
        try:
            # 获取支付平台返回的参数
            params = request.args.to_dict()

            # 验证必要参数
            required_params = ['pid', 'trade_no', 'out_trade_no', 'type', 'name', 'money', 'trade_status', 'sign', 'sign_type']
            for param in required_params:
                if param not in params or not params[param]:
                    print(f"缺少必要参数: {param}")
                    flash('支付参数错误')
                    return redirect(url_for('index'))

            # 验证签名
            received_sign = params.get('sign')
            calculated_sign = generate_sign(params, PAYMENT_CONFIG['key'])

            if received_sign != calculated_sign:
                print(f"签名验证失败: received={received_sign}, calculated={calculated_sign}")
                flash('签名验证失败')
                return redirect(url_for('index'))

            # 检查支付状态
            out_trade_no = params.get('out_trade_no')
            task_id = params.get('param')

            if params.get('trade_status') == 'TRADE_SUCCESS':
                flash('支付成功！')
                # 更新打印任务状态为已支付
                if task_id:
                    updates = {
                        "status": "paid",
                        "message": "支付成功，等待打印",
                        "progress": 50,
                        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    if update_task(task_id, updates):
                        print(f"支付成功: 订单号={out_trade_no}, 任务ID={task_id}")
                        
                        # 自动执行打印任务
                        try:
                            # 使用requests库调用execute_print函数
                            import requests
                            server_url = request.host_url
                            response = requests.get(f"{server_url}execute_print/{task_id}")
                            if response.status_code != 200:
                                print(f"执行打印任务失败: HTTP状态码={response.status_code}")
                        except Exception as e:
                            print(f"执行打印任务异常: {str(e)}")
                
                # 跳转到任务状态页面
                if task_id:
                    return redirect(url_for('task_status', task_id=task_id))
                return redirect(url_for('list_tasks'))
            else:
                print(f"支付状态异常: {params.get('trade_status')}")
                flash('支付未完成或失败')
                return redirect(url_for('index'))
        except Exception as e:
            print(f"支付跳转处理异常: {str(e)}")
            flash('支付处理异常')
            return redirect(url_for('index'))

print(f"========================================")
print(f"支付模块已加载")
