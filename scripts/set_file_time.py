# -*- coding: utf-8 -*-
"""
file-time-magic — 文件时间属性修改脚本 v2.0

功能：
1. 解析用户输入的编辑时长（支持多种格式）
2. 支持用户指定创建时间、修改时间
3. 计算随机化的时间（分钟有误差，秒数随机）
4. 修改 Office 文件内部 XML（TotalTime 属性）
5. 设置文件系统时间属性

使用：
  python set_file_time.py --file "文件路径" --edit-duration "2小时"
  python set_file_time.py --file "文件路径" --create-time "2024-01-15 09:30:00" --edit-duration "3小时"
  python set_file_time.py --file "文件路径" --modify-time "2026-04-18 14:00:00" --edit-duration "90分钟"
  python set_file_time.py --file "文件路径" --create-time "2024-01-15 09:00:00" --modify-time "2026-04-18 14:00:00"
"""

import argparse
import zipfile
import os
import shutil
import random
import re
import json
import subprocess
from datetime import datetime, timedelta
from xml.etree import ElementTree as ET
from pathlib import Path


def parse_duration(text: str) -> int:
    """
    解析用户输入的时长，返回分钟数（整数）
    
    支持格式：
    - "2小时"、"2小时30分"、"两小时半"
    - "120分钟"、"90"（纯数字默认分钟）
    - "3h"、"2h30m"、"150m"
    - "1.5小时"
    """
    text = text.strip().lower()
    
    # 替换中文
    text = text.replace("两", "2")
    text = text.replace("半", "30分")
    text = text.replace("小时", "h")
    text = text.replace("分钟", "m")
    text = text.replace("分", "m")
    text = text.replace("时", "h")
    
    total_minutes = 0
    
    # 匹配 "Xh Ym" 或 "Xh" 或 "Ym" 格式
    pattern = r'(\d+(?:\.\d+)?)\s*([hm]?)'
    matches = re.findall(pattern, text)
    
    for value, unit in matches:
        value = float(value)
        if unit == 'h' or (unit == '' and 'h' in text and 'm' not in text):
            total_minutes += int(value * 60)
        elif unit == 'm' or unit == '':
            total_minutes += int(value)
    
    # 如果没匹配到任何内容，尝试纯数字
    if total_minutes == 0:
        try:
            total_minutes = int(float(text))
        except:
            pass
    
    return total_minutes


def parse_time_str(text: str) -> datetime:
    """
    解析时间字符串，支持多种格式
    
    支持格式：
    - "2024-01-15 09:30:00"
    - "2024-01-15 09:30"
    - "2024/01/15 09:30"
    - "01-15 09:30"（默认今年）
    - "09:30"（默认今天）
    """
    text = text.strip()
    
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %H:%M',
        '%Y/%m/%d %H:%M:%S',
        '%Y/%m/%d %H:%M',
        '%m-%d %H:%M:%S',
        '%m-%d %H:%M',
        '%H:%M:%S',
        '%H:%M',
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(text, fmt)
            # 如果格式不包含年份，使用当前年
            if '%Y' not in fmt:
                dt = dt.replace(year=datetime.now().year)
            # 如果格式不包含日期，使用今天
            if '%d' not in fmt:
                dt = dt.replace(
                    year=datetime.now().year,
                    month=datetime.now().month,
                    day=datetime.now().day
                )
            return dt
        except ValueError:
            continue
    
    raise ValueError(f"无法解析时间: {text}")


def randomize_duration(minutes: int) -> tuple:
    """
    随机化时长
    返回：(实际分钟数, 秒数)
    """
    # 根据时长决定误差范围
    if minutes >= 60:
        # 大于1小时，误差 ±5 分钟
        actual_minutes = minutes + random.randint(-5, 5)
    else:
        # 小于1小时，误差 ±3 分钟
        actual_minutes = minutes + random.randint(-3, 3)
    
    # 确保不低于最小值
    actual_minutes = max(actual_minutes, 1)
    
    # 随机秒数
    actual_seconds = random.randint(0, 59)
    
    return actual_minutes, actual_seconds


def is_work_hour(hour: int) -> bool:
    """检查是否在工作时间（08:00 - 22:00）"""
    return 8 <= hour <= 22


def adjust_to_work_time(dt: datetime) -> datetime:
    """将时间调整到工作时间范围"""
    if dt.hour < 8:
        # 太早，调整到早上8-10点
        return dt.replace(hour=random.randint(8, 10), minute=random.randint(0, 59))
    elif dt.hour > 22:
        # 太晚，调整到前一天晚上8-10点
        return (dt - timedelta(days=1)).replace(hour=random.randint(20, 22), minute=random.randint(0, 59))
    return dt


def add_random_seconds(dt: datetime) -> datetime:
    """添加随机秒数"""
    return dt.replace(second=random.randint(0, 59))


def calculate_times_v2(
    edit_minutes: int = None,
    edit_seconds: int = None,
    create_time: datetime = None,
    modify_time: datetime = None,
    access_time: datetime = None,
    base_time: datetime = None
) -> dict:
    """
    计算所有时间属性（v2 版本，支持用户指定时间点）
    
    参数：
    - edit_minutes: 编辑分钟数（可选）
    - edit_seconds: 编辑秒数（可选）
    - create_time: 用户指定的创建时间（可选）
    - modify_time: 用户指定的修改时间（可选）
    - access_time: 用户指定的访问时间（可选）
    - base_time: 参考时间（默认当前时间）
    
    返回：
    - dict: {create, modify, access, edit_minutes, edit_seconds}
    """
    if base_time is None:
        base_time = datetime.now()
    
    result = {
        'create': None,
        'modify': None,
        'access': None,
        'edit_minutes': edit_minutes,
        'edit_seconds': edit_seconds if edit_seconds is not None else random.randint(0, 59)
    }
    
    # 情况1：用户指定了创建时间和修改时间
    if create_time and modify_time:
        # 用户明确指定的时间，不强制调整工作时间，但添加随机秒数
        result['create'] = add_random_seconds(create_time)
        result['modify'] = add_random_seconds(modify_time)
        
        # 编辑时长：如果用户指定了，就用用户的；否则自动计算
        if edit_minutes is None:
            duration = result['modify'] - result['create']
            result['edit_minutes'] = int(duration.total_seconds() / 60)
            # 秒数也随机，不用计算值
            result['edit_seconds'] = random.randint(0, 59)
        # 否则保持用户指定的 edit_minutes，秒数已随机初始化
        
    # 情况2：用户指定了创建时间和编辑时长
    elif create_time and edit_minutes:
        # 创建时间：用户指定，不调整工作时间
        result['create'] = add_random_seconds(create_time)
        # 修改时间：系统计算，添加随机缓冲
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=edit_minutes + random.randint(0, 2))
        )
        
    # 情况3：用户指定了修改时间和编辑时长
    elif modify_time and edit_minutes:
        # 修改时间：用户指定，不调整工作时间
        result['modify'] = add_random_seconds(modify_time)
        # 创建时间：系统计算，调整到工作时间
        calc_create = result['modify'] - timedelta(minutes=edit_minutes + random.randint(0, 2))
        result['create'] = add_random_seconds(adjust_to_work_time(calc_create))
        
    # 情况4：用户只指定了编辑时长（当前行为）
    elif edit_minutes:
        # 从基准时间往前推算
        buffer_minutes = random.randint(5, 30)
        result['create'] = base_time - timedelta(
            minutes=edit_minutes + buffer_minutes,
            seconds=random.randint(0, 59)
        )
        
        # 调整到工作时间
        result['create'] = adjust_to_work_time(result['create'])
        result['create'] = add_random_seconds(result['create'])
        
        # 修改时间
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=edit_minutes + random.randint(0, 2))
        )
        
    # 情况5：用户什么都没指定
    else:
        # 使用默认行为：编辑时长随机30-60分钟
        default_edit = random.randint(30, 60)
        result['edit_minutes'] = default_edit
        
        buffer_minutes = random.randint(5, 30)
        result['create'] = adjust_to_work_time(
            base_time - timedelta(minutes=default_edit + buffer_minutes)
        )
        result['create'] = add_random_seconds(result['create'])
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=default_edit)
        )
    
    # 处理访问时间
    if access_time:
        result['access'] = add_random_seconds(access_time)
    else:
        # 访问时间 = 修改时间 + 随机间隔
        result['access'] = add_random_seconds(
            result['modify'] + timedelta(minutes=random.randint(3, 15))
        )
    
    # 确保时间逻辑正确
    if result['create'] > result['modify']:
        result['modify'] = result['create'] + timedelta(minutes=result['edit_minutes'] or 30)
    if result['modify'] > result['access']:
        result['access'] = result['modify'] + timedelta(minutes=random.randint(3, 15))
    
    # 确保不超过基准时间（只调整访问时间，不调整用户明确指定的时间）
    # 如果用户指定了创建或修改时间，跳过这个检查
    user_specified_times = create_time is not None or modify_time is not None
    
    if not user_specified_times and result['access'] > base_time:
        # 按比例缩小（只在完全自动计算模式下）
        total_span = (result['access'] - result['create']).total_seconds()
        allowed_span = (base_time - result['create']).total_seconds()
        if total_span > 0 and allowed_span > 0:
            ratio = allowed_span / total_span
            result['modify'] = result['create'] + timedelta(seconds=total_span * ratio * 0.85)
            result['access'] = result['create'] + timedelta(seconds=total_span * ratio * 0.95)
    elif result['access'] > base_time:
        # 用户指定了时间，只调整访问时间
        # 如果修改时间已经早于基准时间，访问时间可以正常设置
        # 如果修改时间晚于基准时间（未来），访问时间等于修改时间
        if result['modify'] > base_time:
            result['access'] = result['modify']
        else:
            result['access'] = min(result['access'], base_time - timedelta(minutes=random.randint(1, 5)))
    
    # 验证：编辑时长不能超过文件存在时间
    file_exist_minutes = int((result['modify'] - result['create']).total_seconds() / 60)
    if result['edit_minutes'] and result['edit_minutes'] > file_exist_minutes:
        # 警告：编辑时长超过文件存在时间，自动调整
        result['edit_minutes'] = max(1, file_exist_minutes - random.randint(1, 5))
        result['edit_seconds'] = random.randint(0, 59)
    
    # 警告信息：创建时间在非工作时间
    if result['create'] and not is_work_hour(result['create'].hour):
        result['work_time_warning'] = f"创建时间 {result['create'].strftime('%H:%M')} 不在常规工作时间（08:00-22:00）"
    
    return result


def modify_office_internal(file_path: str, edit_minutes: int) -> bool:
    """
    修改 Office 文件（docx/pptx/xlsx）内部的编辑时长属性
    
    返回：是否成功修改
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext not in ['.docx', '.pptx', '.xlsx']:
        return False
    
    tmp_dir = os.path.join(os.environ.get('TEMP', '/tmp'), '_office_edit_time_tmp')
    
    # 清理旧临时目录
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    
    # 解压
    with zipfile.ZipFile(file_path, 'r') as z:
        z.extractall(tmp_dir)
    
    # 查找 app.xml 或对应的属性文件
    app_xml = os.path.join(tmp_dir, 'docProps', 'app.xml')
    
    if not os.path.exists(app_xml):
        shutil.rmtree(tmp_dir)
        return False
    
    # 修改 TotalTime
    tree = ET.parse(app_xml)
    root = tree.getroot()
    
    ns = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
    total_time = root.find('ep:TotalTime', ns)
    
    if total_time is not None:
        total_time.text = str(edit_minutes)
    else:
        # 创建 TotalTime 元素
        total_time = ET.SubElement(root, '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}TotalTime')
        total_time.text = str(edit_minutes)
    
    tree.write(app_xml, xml_declaration=True, encoding='UTF-8')
    
    # 重新打包
    new_file = file_path + '.new'
    with zipfile.ZipFile(new_file, 'w', zipfile.ZIP_DEFLATED) as z:
        for root_dir, dirs, files in os.walk(tmp_dir):
            for file in files:
                file_path_item = os.path.join(root_dir, file)
                arc_name = os.path.relpath(file_path_item, tmp_dir)
                z.write(file_path_item, arc_name)
    
    # 替换原文件
    os.remove(file_path)
    os.rename(new_file, file_path)
    
    # 清理
    shutil.rmtree(tmp_dir)
    
    return True


def set_file_system_times(file_path: str, create: datetime, modify: datetime, access: datetime):
    """
    设置文件系统时间属性
    - 文件：使用 PowerShell
    - 文件夹：使用 Python os.utime + Windows API
    """
    if os.path.isdir(file_path):
        # 文件夹：使用 Python + Windows API
        return set_folder_times(file_path, create, modify, access)
    else:
        # 文件：使用 PowerShell
        return set_file_times_powershell(file_path, create, modify, access)


def set_file_times_powershell(file_path: str, create: datetime, modify: datetime, access: datetime):
    """
    使用 PowerShell 设置文件系统时间属性
    """
    ps_script = f'''
$f = "{file_path.replace('\\', '\\\\')}"
(Get-Item $f).CreationTime = "{create.strftime('%Y-%m-%d %H:%M:%S')}"
(Get-Item $f).LastWriteTime = "{modify.strftime('%Y-%m-%d %H:%M:%S')}"
(Get-Item $f).LastAccessTime = "{access.strftime('%Y-%m-%d %H:%M:%S')}"
'''
    
    result = subprocess.run(
        ['powershell', '-Command', ps_script],
        capture_output=True,
        text=True,
        timeout=10
    )
    
    return result.returncode == 0


def set_folder_times(folder_path: str, create: datetime, modify: datetime, access: datetime):
    """
    设置文件夹的时间属性
    - 创建时间：使用 Windows API SetFileTime
    - 修改时间 + 访问时间：使用 os.utime
    """
    import ctypes
    from ctypes import wintypes
    
    try:
        # 1. 设置修改时间和访问时间（用 os.utime）
        import time
        modify_ts = time.mktime(modify.timetuple())
        access_ts = time.mktime(access.timetuple())
        os.utime(folder_path, (access_ts, modify_ts))
        
        # 2. 设置创建时间（用 Windows API）
        kernel32 = ctypes.windll.kernel32
        
        class FILETIME(ctypes.Structure):
            _fields_ = [
                ('dwLowDateTime', wintypes.DWORD),
                ('dwHighDateTime', wintypes.DWORD)
            ]
        
        # 将本地时间转换为 UTC（Windows API 需要 UTC 时间）
        # 获取时区偏移
        import time as time_module
        if time_module.daylight:
            # 夏令时
            offset_seconds = -time_module.altzone
        else:
            offset_seconds = -time_module.timezone
        
        # 转换为 UTC
        utc_create = create - timedelta(seconds=offset_seconds)
        
        # 转换为 FILETIME 格式
        epoch = datetime(1601, 1, 1)
        delta = utc_create - epoch
        ft_value = int(delta.total_seconds() * 10_000_000)
        ft = FILETIME(ft_value & 0xFFFFFFFF, ft_value >> 32)
        
        # 打开文件夹
        handle = kernel32.CreateFileW(
            folder_path,
            0x0100,  # FILE_WRITE_ATTRIBUTES
            0x07,    # FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
            None,
            3,       # OPEN_EXISTING
            0x02000000,  # FILE_FLAG_BACKUP_SEMANTICS (必需用于文件夹)
            None
        )
        
        if handle != -1:
            result = kernel32.SetFileTime(handle, ctypes.byref(ft), None, None)
            kernel32.CloseHandle(handle)
            return result == 1
        else:
            return False
        
    except Exception as e:
        return False


def close_office_processes():
    """关闭可能占用文件的 Office 进程"""
    processes = ['WINWORD', 'EXCEL', 'POWERPNT']
    for proc in processes:
        try:
            subprocess.run(['taskkill', '/F', '/IM', f'{proc}.EXE'],
                          capture_output=True, timeout=5)
        except:
            pass


def confirm_future_times(future_times: list, now: datetime) -> bool:
    """
    确认是否设置未来时间
    
    返回：True=继续执行，False=取消
    """
    print("\n" + "="*60)
    print("⚠️  警告：以下时间超过当前时间")
    print("="*60)
    for t in future_times:
        print(f"  • {t}")
    print(f"\n当前时间：{now.strftime('%Y-%m-%d %H:%M:%S')}")
    print("\n这可能是一个不合理的设置（文件来自未来？）")
    print("-"*60)
    
    try:
        response = input("是否继续执行？(y/N): ").strip().lower()
        return response in ['y', 'yes', '是']
    except:
        # 非交互模式，默认拒绝
        print("非交互模式，已取消操作。使用 --force 强制执行。")
        return False


def main():
    parser = argparse.ArgumentParser(description='文件时间属性修改 v2.0')
    parser.add_argument('--file', required=True, help='目标文件路径')
    parser.add_argument('--edit-duration', help='编辑时长，如 "2小时"、"120分钟"')
    parser.add_argument('--create-time', help='创建时间，如 "2024-01-15 09:30:00"')
    parser.add_argument('--modify-time', help='修改时间，如 "2026-04-18 14:00:00"')
    parser.add_argument('--access-time', help='访问时间（可选）')
    parser.add_argument('--base-time', help='参考时间，如 "2026-04-18 14:30:00"')
    parser.add_argument('--dry-run', action='store_true', help='只显示计算结果，不实际修改')
    parser.add_argument('--force', action='store_true', help='强制执行，跳过未来时间确认')
    
    args = parser.parse_args()
    
    # 检查文件是否存在
    if not os.path.exists(args.file):
        print(json.dumps({
            'status': 'error',
            'message': f'文件不存在: {args.file}'
        }, ensure_ascii=False))
        return
    
    # 解析编辑时长
    edit_minutes = None
    edit_seconds = None
    requested_duration = None
    
    if args.edit_duration:
        requested_duration = args.edit_duration
        parsed_minutes = parse_duration(args.edit_duration)
        if parsed_minutes > 0:
            edit_minutes, edit_seconds = randomize_duration(parsed_minutes)
    
    # 解析用户指定的时间
    create_time = None
    modify_time = None
    access_time = None
    
    if args.create_time:
        try:
            create_time = parse_time_str(args.create_time)
        except ValueError as e:
            print(json.dumps({
                'status': 'error',
                'message': f'创建时间格式错误: {args.create_time}'
            }, ensure_ascii=False))
            return
    
    if args.modify_time:
        try:
            modify_time = parse_time_str(args.modify_time)
        except ValueError as e:
            print(json.dumps({
                'status': 'error',
                'message': f'修改时间格式错误: {args.modify_time}'
            }, ensure_ascii=False))
            return
    
    if args.access_time:
        try:
            access_time = parse_time_str(args.access_time)
        except ValueError as e:
            print(json.dumps({
                'status': 'error',
                'message': f'访问时间格式错误: {args.access_time}'
            }, ensure_ascii=False))
            return
    
    # 解析基准时间
    base_time = None
    if args.base_time:
        try:
            base_time = parse_time_str(args.base_time)
        except:
            print(json.dumps({
                'status': 'error',
                'message': f'无效的基准时间格式: {args.base_time}'
            }, ensure_ascii=False))
            return
    
    # 计算所有时间
    times = calculate_times_v2(
        edit_minutes=edit_minutes,
        edit_seconds=edit_seconds,
        create_time=create_time,
        modify_time=modify_time,
        access_time=access_time,
        base_time=base_time
    )
    
    # 检查是否有未来时间（超过当前时间）
    now = datetime.now()
    future_times = []
    
    if times['create'] > now:
        future_times.append(f"创建时间: {times['create'].strftime('%Y-%m-%d %H:%M:%S')}")
    if times['modify'] > now:
        future_times.append(f"修改时间: {times['modify'].strftime('%Y-%m-%d %H:%M:%S')}")
    if times['access'] > now:
        future_times.append(f"访问时间: {times['access'].strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 如果有未来时间且不是dry-run模式，需要确认
    if future_times and not args.dry_run and not args.force:
        confirmed = confirm_future_times(future_times, now)
        if not confirmed:
            print(json.dumps({
                'status': 'cancelled',
                'message': '用户取消操作'
            }, ensure_ascii=False))
            return
    
    # 准备输出
    result = {
        'status': 'ok',
        'file': args.file,
        'times': {
            'create': times['create'].strftime('%Y-%m-%d %H:%M:%S'),
            'modify': times['modify'].strftime('%Y-%m-%d %H:%M:%S'),
            'access': times['access'].strftime('%Y-%m-%d %H:%M:%S')
        },
        'edit_duration': {
            'requested': requested_duration,
            'parsed_minutes': edit_minutes,
            'actual_minutes': times['edit_minutes'],
            'actual_seconds': times['edit_seconds']
        },
        'user_specified': {
            'create_time': args.create_time,
            'modify_time': args.modify_time,
            'access_time': args.access_time
        },
        'dry_run': args.dry_run
    }
    
    # 添加警告信息
    warnings = []
    if 'work_time_warning' in times:
        warnings.append(times['work_time_warning'])
    
    # 检查编辑时长是否被自动调整
    file_exist_minutes = int((times['modify'] - times['create']).total_seconds() / 60)
    if edit_minutes and edit_minutes > file_exist_minutes:
        warnings.append(f"编辑时长({edit_minutes}分钟)超过文件存在时间({file_exist_minutes}分钟)，已自动调整为{times['edit_minutes']}分钟")
    
    if warnings:
        result['warnings'] = warnings
    
    if args.dry_run:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    
    # 实际修改
    try:
        # 先关闭可能占用的 Office 进程
        close_office_processes()
        
        # 修改 Office 文件内部属性
        ext = os.path.splitext(args.file)[1].lower()
        if ext in ['.docx', '.pptx', '.xlsx'] and times['edit_minutes']:
            internal_modified = modify_office_internal(args.file, times['edit_minutes'])
            result['docx_internal'] = {
                'modified': internal_modified,
                'property': 'TotalTime',
                'value': times['edit_minutes']
            }
        
        # 设置文件系统时间
        success = set_file_system_times(
            args.file,
            times['create'],
            times['modify'],
            times['access']
        )
        
        if not success:
            result['status'] = 'warning'
            result['message'] = '文件系统时间设置可能未完全成功'
        
    except Exception as e:
        result['status'] = 'error'
        result['message'] = str(e)
    
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
