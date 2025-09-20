# 任务存储模块

# ================== 任务存储 ==================

# 全局任务字典
print_tasks = {}

def get_print_tasks():
    """获取打印任务字典"""
    global print_tasks
    return print_tasks

def set_print_tasks(tasks):
    """设置打印任务字典"""
    global print_tasks
    print_tasks = tasks

def get_task(task_id):
    """获取指定任务"""
    global print_tasks
    return print_tasks.get(task_id)

def update_task(task_id, updates):
    """更新任务"""
    global print_tasks
    if task_id in print_tasks:
        for key, value in updates.items():
            print_tasks[task_id][key] = value
        return True
    return False

print(f"========================================")
print(f"任务存储模块已加载")
