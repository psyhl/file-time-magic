# -*- coding: utf-8 -*-
"""
file-time-magic -- File Time Attribute Modifier v1.1.1.1

功能:
1. 解析用户输入的编辑时长(支持多种格式)
2. 支持用户指定创建时间、修改时间
3. 计算随机化的时间(分钟有误差,秒数随机)
4. 修改 Office 文件内部 XML(TotalTime 属性 + core.xml 创建/修改时间)
5. 设置文件系统时间属性

使用:
  python set_file_time.py --file "文件路径" --edit-duration "2小时"
  python set_file_time.py --file "文件路径" --create-time "2024-01-15 09:30:00" --edit-duration "3小时"
  python set_file_time.py --file "文件路径" --modify-time "2026-04-18 14:00:00" --edit-duration "90分钟"
  python set_file_time.py --file "文件路径" --create-time "2024-01-15 09:00:00" --modify-time "2026-04-18 14:00:00"
  python set_file_time.py --file "文件路径" --create-time "2026-04-17 09:00" --modify-time "2026-04-18 18:00" --total-edit-minutes 480
"""

import argparse
import zipfile
import os
import shutil
import random
import re
import json
import subprocess
import time
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
    text = text.replace("半", ".5")
    text = text.replace("小时", "h")
    text = text.replace("分钟", "m")
    text = text.replace("分", "m")
    text = text.replace("时", "h")

    total_minutes = 0

    # 先匹配 "Xh Ym" 或 "XhYm" 等复合格式
    # 例如 "2h30m", "1.5h", "30m"
    # 匹配所有数字+单位组合
    pattern = r'([\d.]+)\s*h'
    h_matches = re.findall(pattern, text)
    for val in h_matches:
        total_minutes += int(float(val) * 60)

    pattern = r'([\d.]+)\s*m(?!o)'  # m 不匹配 "mo" 开头
    m_matches = re.findall(pattern, text)
    for val in m_matches:
        total_minutes += int(float(val))

    # 如果没匹配到任何内容，尝试纯数字（默认分钟）
    if total_minutes == 0:
        try:
            total_minutes = int(float(text))
        except:
            pass

    return total_minutes


def parse_time_str(text: str) -> datetime:
    """
    解析时间字符串，支持多种格式
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
            if '%Y' not in fmt:
                dt = dt.replace(year=datetime.now().year)
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
    """随机化时长，返回 (实际分钟数, 秒数)"""
    if minutes >= 60:
        actual_minutes = minutes + random.randint(-5, 5)
    else:
        actual_minutes = minutes + random.randint(-3, 3)
    actual_minutes = max(actual_minutes, 1)
    actual_seconds = random.randint(0, 59)
    return actual_minutes, actual_seconds


def is_work_hour(hour: int) -> bool:
    """检查是否在工作时间（08:00 - 22:00）"""
    return 8 <= hour <= 22


def adjust_to_work_time(dt: datetime) -> datetime:
    """将时间调整到工作时间范围"""
    if dt.hour < 8:
        return dt.replace(hour=random.randint(8, 10), minute=random.randint(0, 59))
    elif dt.hour > 22:
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
    计算所有时间属性
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
        result['create'] = add_random_seconds(create_time)
        result['modify'] = add_random_seconds(modify_time)
        if edit_minutes is None:
            duration = result['modify'] - result['create']
            result['edit_minutes'] = int(duration.total_seconds() / 60)
            result['edit_seconds'] = random.randint(0, 59)

    # 情况2：用户指定了创建时间和编辑时长
    elif create_time and edit_minutes:
        result['create'] = add_random_seconds(create_time)
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=edit_minutes + random.randint(0, 2))
        )

    # 情况3：用户指定了修改时间和编辑时长
    elif modify_time and edit_minutes:
        result['modify'] = add_random_seconds(modify_time)
        calc_create = result['modify'] - timedelta(minutes=edit_minutes + random.randint(0, 2))
        result['create'] = add_random_seconds(adjust_to_work_time(calc_create))

    # 情况4：用户只指定了编辑时长
    elif edit_minutes:
        buffer_minutes = random.randint(5, 30)
        result['create'] = base_time - timedelta(
            minutes=edit_minutes + buffer_minutes,
            seconds=random.randint(0, 59)
        )
        result['create'] = adjust_to_work_time(result['create'])
        result['create'] = add_random_seconds(result['create'])
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=edit_minutes + random.randint(0, 2))
        )

    # 情况5：什么都没指定
    else:
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

    # 访问时间
    if access_time:
        result['access'] = add_random_seconds(access_time)
    else:
        result['access'] = add_random_seconds(
            result['modify'] + timedelta(minutes=random.randint(3, 15))
        )

    # 确保时间逻辑正确
    if result['create'] > result['modify']:
        result['modify'] = result['create'] + timedelta(minutes=result['edit_minutes'] or 30)
    if result['modify'] > result['access']:
        result['access'] = result['modify'] + timedelta(minutes=random.randint(3, 15))

    # 自动计算模式下：确保不超出基准时间
    user_specified_times = create_time is not None or modify_time is not None
    if not user_specified_times and result['access'] > base_time:
        total_span = (result['access'] - result['create']).total_seconds()
        allowed_span = (base_time - result['create']).total_seconds()
        if total_span > 0 and allowed_span > 0:
            ratio = allowed_span / total_span
            result['modify'] = result['create'] + timedelta(seconds=total_span * ratio * 0.85)
            result['access'] = result['create'] + timedelta(seconds=total_span * ratio * 0.95)
    elif result['access'] > base_time:
        if result['modify'] > base_time:
            result['access'] = result['modify']
        else:
            result['access'] = min(result['access'], base_time - timedelta(minutes=random.randint(1, 5)))

    # 验证：编辑时长不超过文件存在时间
    file_exist_minutes = int((result['modify'] - result['create']).total_seconds() / 60)
    if result['edit_minutes'] and result['edit_minutes'] > file_exist_minutes:
        result['edit_minutes'] = max(1, file_exist_minutes - random.randint(1, 5))
        result['edit_seconds'] = random.randint(0, 59)

    # 非工作时间警告
    if result['create'] and not is_work_hour(result['create'].hour):
        result['work_time_warning'] = (
            f"创建时间 {result['create'].strftime('%H:%M')} "
            f"不在常规工作时间（08:00-22:00），已自动调整"
        )

    return result


def modify_office_internal(file_path: str, edit_minutes: int,
                             create_dt: datetime = None,
                             modify_dt: datetime = None) -> bool:
    """
    修改 Office 文件（docx/pptx/xlsx）内部属性
    - TotalTime（编辑时长）：来自 docProps/app.xml
    - 创建时间、修改时间：来自 docProps/core.xml

    Args:
        file_path:  文件路径
        edit_minutes:  编辑时长（分钟）
        create_dt:     创建时间（datetime）
        modify_dt:     修改时间（datetime）
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ['.docx', '.pptx', '.xlsx']:
        return False

    # 使用 PID 区分临时目录，避免并发冲突
    tmp_dir = os.path.join(os.environ.get('TEMP', '/tmp'),
                           '_ftm_' + str(os.getpid()) + '_' + str(random.randint(1000, 9999)))

    # 清理旧残留
    if os.path.exists(tmp_dir):
        try:
            shutil.rmtree(tmp_dir)
        except Exception:
            pass
    os.makedirs(tmp_dir)

    try:
        # 解压
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tmp_dir)

        # === app.xml: TotalTime ===
        app_xml = os.path.join(tmp_dir, 'docProps', 'app.xml')
        if os.path.exists(app_xml):
            tree = ET.parse(app_xml)
            root = tree.getroot()
            ns = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
            total = root.find(f'{{{ns}}}TotalTime')
            if total is None:
                total = ET.SubElement(root, f'{{{ns}}}TotalTime')
            total.text = str(edit_minutes)
            tree.write(app_xml, xml_declaration=True, encoding='UTF-8')

        # === core.xml: 创建时间 + 修改时间 ===
        core_xml = os.path.join(tmp_dir, 'docProps', 'core.xml')
        if os.path.exists(core_xml) and create_dt and modify_dt:
            tree = ET.parse(core_xml)
            root = tree.getroot()

            xsi_ns = 'http://www.w3.org/2001/XMLSchema-instance'
            dct_ns = 'http://purl.org/dc/terms/'

            create_utc = create_dt.strftime('%Y-%m-%dT%H:%M:%SZ')
            modify_utc = modify_dt.strftime('%Y-%m-%dT%H:%M:%SZ')

            # 更新 dcterms:created（Word属性对话框中的"创建时间"）
            for el in root.iter(dct_ns + 'created'):
                el.set(f'{{{xsi_ns}}}type', 'dcterms:W3CDTF')
                el.text = create_utc

            # 更新 dcterms:modified（Word属性对话框中的"修改时间"）
            for el in root.iter(dct_ns + 'modified'):
                el.set(f'{{{xsi_ns}}}type', 'dcterms:W3CDTF')
                el.text = modify_utc

            tree.write(core_xml, xml_declaration=True, encoding='UTF-8')

        # 重新打包
        new_file = file_path + '.tmp'
        with zipfile.ZipFile(new_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for rd, dirs, files in os.walk(tmp_dir):
                for fn in files:
                    fp = os.path.join(rd, fn)
                    z.write(fp, os.path.relpath(fp, tmp_dir))

        # 替换原文件
        os.remove(file_path)
        os.rename(new_file, file_path)

    finally:
        # 清理临时目录
        if os.path.exists(tmp_dir):
            try:
                shutil.rmtree(tmp_dir)
            except Exception:
                pass

    return True


def set_file_system_times(file_path: str,
                           create: datetime,
                           modify: datetime,
                           access: datetime) -> bool:
    """
    设置文件系统时间属性
    - 文件夹：使用 Python os.utime + Windows API（SetFileTime）
    - 文件：使用 os.utime（先尝试，失败则用 PowerShell）
    """
    if os.path.isdir(file_path):
        return _set_folder_times(file_path, create, modify, access)
    else:
        return _set_file_times(file_path, create, modify, access)


def _utc_offset_hours() -> int:
    """获取当前时区与UTC的偏移量（小时），Windows专用"""
    # time.timezone 是从本地时间往UTC走的偏移（秒），west为正
    # time.altzone 是夏令时偏移（如果有）
    # UTC+8: time.timezone = -28800 (8*3600), time.altzone = -28800
    offset_sec = time.timezone if time.daylight == 0 else time.altzone
    return -offset_sec // 3600  # UTC+8 -> +8


def _set_file_times(file_path: str,
                     create: datetime,
                     modify: datetime,
                     access: datetime) -> bool:
    """
    用 Python os 设置文件时间（跨平台，不依赖 PowerShell）
    """
    try:
        import ctypes
        from ctypes import wintypes

        # os.utime 设置 modify 和 access（这两个用 datetime.timestamp() 正确转换）
        access_ts = access.timestamp()
        modify_ts = modify.timestamp()
        os.utime(file_path, (access_ts, modify_ts))

        # 创建时间用 Windows API SetFileTime
        # 注意：SetFileTime 接收 UTC 时间
        # 输入的 datetime 是本地时间（CST），需要转 UTC
        # Windows FILETIME = 100-ns ticks since 1601-01-01 UTC
        kernel32 = ctypes.windll.kernel32

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD),
                        ('dwHighDateTime', wintypes.DWORD)]

        def dt_to_filetime_utc(dt: datetime) -> FILETIME:
            # 将本地时间转为 UTC 再算 FILETIME
            utc_offset = _utc_offset_hours()
            utc_dt = dt - timedelta(hours=utc_offset)
            epoch = datetime(1601, 1, 1)
            ft_value = int((utc_dt - epoch).total_seconds() * 10_000_000)
            lo = wintypes.DWORD(ft_value & 0xFFFFFFFF)
            hi = wintypes.DWORD(ft_value >> 32)
            return FILETIME(lo, hi)

        handle = kernel32.CreateFileW(
            file_path, 0x0100, 0, None, 3, 0, None)
        if handle == -1:
            return False
        try:
            ft = dt_to_filetime_utc(create)
            ok = kernel32.SetFileTime(handle, ctypes.byref(ft), None, None)
            return ok == 1
        finally:
            kernel32.CloseHandle(handle)

    except Exception:
        return False


def _set_folder_times(folder_path: str,
                       create: datetime,
                       modify: datetime,
                       access: datetime) -> bool:
    """设置文件夹时间属性（Windows 专用）"""
    try:
        import ctypes
        from ctypes import wintypes

        # 修改时间和访问时间用 os.utime
        os.utime(folder_path, (access.timestamp(), modify.timestamp()))

        # 创建时间用 Windows API SetFileTime（需要 UTC）
        kernel32 = ctypes.windll.kernel32

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD),
                        ('dwHighDateTime', wintypes.DWORD)]

        def dt_to_filetime_utc(dt: datetime) -> FILETIME:
            utc_offset = _utc_offset_hours()
            utc_dt = dt - timedelta(hours=utc_offset)
            epoch = datetime(1601, 1, 1)
            ft_value = int((utc_dt - epoch).total_seconds() * 10_000_000)
            lo = wintypes.DWORD(ft_value & 0xFFFFFFFF)
            hi = wintypes.DWORD(ft_value >> 32)
            return FILETIME(lo, hi)

        handle = kernel32.CreateFileW(
            folder_path,
            0x0100, 0, None, 3, 0x02000000, None)  # FILE_FLAG_BACKUP_SEMANTICS
        if handle == -1:
            return False
        try:
            ft = dt_to_filetime_utc(create)
            ok = kernel32.SetFileTime(handle, ctypes.byref(ft), None, None)
            return ok == 1
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return False


def close_office_processes():
    """关闭可能占用文件的 Office 进程"""
    for proc in ['WINWORD', 'EXCEL', 'POWERPNT']:
        try:
            subprocess.run(['taskkill', '/F', '/IM', f'{proc}.EXE'],
                           capture_output=True, timeout=5)
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(description='文件时间属性修改 v1.1.1')
    parser.add_argument('--file', required=True, help='目标文件路径')
    parser.add_argument('--edit-duration', help='编辑时长，如 "2小时"、"120分钟"、"3h"')
    parser.add_argument(
        '--total-edit-minutes', type=int,
        help='总编辑时长（分钟），优先级高于 --edit-duration，用于 "分多天编辑共X小时" 场景'
    )
    parser.add_argument('--create-time', help='创建时间，如 "2024-01-15 09:30:00"')
    parser.add_argument('--modify-time', help='修改时间，如 "2026-04-18 14:00:00"')
    parser.add_argument('--access-time', help='访问时间（可选）')
    parser.add_argument('--base-time', help='参考时间（默认当前时间）')
    parser.add_argument('--dry-run', action='store_true', help='只显示计算结果，不实际修改')
    parser.add_argument('--force', action='store_true', help='强制执行，跳过未来时间确认')

    args = parser.parse_args()

    # -- 解析编辑时长 ---------------------------------------------
    # --total-edit-minutes > --edit-duration > 自动从 create+modify 差值算
    edit_minutes = None
    edit_seconds = None
    requested_duration = None

    if args.total_edit_minutes is not None:
        # 用户显式指定总编辑分钟数（独立于时间锚点）
        edit_minutes, edit_seconds = randomize_duration(args.total_edit_minutes)
        requested_duration = f"{args.total_edit_minutes}分钟（指定）"
    elif args.edit_duration:
        requested_duration = args.edit_duration
        parsed = parse_duration(args.edit_duration)
        if parsed > 0:
            edit_minutes, edit_seconds = randomize_duration(parsed)

    # -- 解析时间锚点 ----------------------------------------------
    create_time = modify_time = access_time = base_time = None

    def _parse(t):
        try:
            return parse_time_str(t)
        except ValueError:
            return None

    if args.create_time:
        create_time = _parse(args.create_time)
        if create_time is None:
            print(json.dumps({'status': 'error',
                               'message': f'创建时间格式错误: {args.create_time}'}, ensure_ascii=False))
            return
    if args.modify_time:
        modify_time = _parse(args.modify_time)
        if modify_time is None:
            print(json.dumps({'status': 'error',
                               'message': f'修改时间格式错误: {args.modify_time}'}, ensure_ascii=False))
            return
    if args.access_time:
        access_time = _parse(args.access_time)
    if args.base_time:
        base_time = _parse(args.base_time)

    # -- 组合逻辑 -------------------------------------------------
    # 优先：全部锚定（create + modify + duration）
    # 其次：create + modify -> 自动算 duration
    # 其次：create + duration -> 算 modify
    # 其次：modify + duration -> 算 create
    # 最后：只有 duration -> 从当前往回推

    warnings = []
    now = datetime.now()
    if base_time is None:
        base_time = now

    if create_time and modify_time:
        # Case C/D: 两端锚定
        if create_time >= modify_time:
            print(json.dumps({'status': 'error',
                               'message': '创建时间不能晚于修改时间'}, ensure_ascii=False))
            return
        if edit_minutes is None:
            # 自动从差值算
            diff_min = int((modify_time - create_time).total_seconds() / 60)
            edit_minutes, edit_seconds = randomize_duration(diff_min)
            requested_duration = f"{diff_min}分钟（自动）"
        # 注意：如果用户也给了 total-edit-minutes，保持用那个值
        access_time = modify_time + timedelta(minutes=random.randint(3, 15))
        # 创建时间可能在工作时间外 -> 警告
        if not (8 <= create_time.hour < 22):
            warnings.append(f"创建时间 {create_time.strftime('%H:%M')} 不在工作时间(08:00-22:00)内")
    elif create_time and edit_minutes is not None:
        # Case A: 从创建时间往后推
        modify_time = create_time + timedelta(
            minutes=edit_minutes, seconds=random.randint(0, 119))
        access_time = modify_time + timedelta(minutes=random.randint(3, 15))
        if not (8 <= create_time.hour < 22):
            warnings.append(f"创建时间 {create_time.hour}:00 不在工作时间(08:00-22:00)内")
    elif modify_time and edit_minutes is not None:
        # Case B: 从修改时间往前推
        delta = timedelta(minutes=edit_minutes, seconds=random.randint(0, 119))
        create_time = modify_time - delta
        # 调整到工作时间
        create_time = adjust_to_work_time(create_time)
        access_time = modify_time + timedelta(minutes=random.randint(3, 15))
    elif edit_minutes is not None:
        # Case E: 只有 duration，从当前往回推
        modify_time = base_time
        create_time = base_time - timedelta(
            minutes=edit_minutes, seconds=random.randint(0, 119))
        create_time = adjust_to_work_time(create_time)
        access_time = modify_time + timedelta(minutes=random.randint(3, 15))
    else:
        print(json.dumps({'status': 'error',
                           'message': '至少需要提供编辑时长或时间锚点之一'}, ensure_ascii=False))
        return

    # 秒数随机化（已在上面各分支处理）
    create_time = create_time.replace(second=random.randint(0, 59), microsecond=0)
    modify_time = modify_time.replace(second=random.randint(0, 59), microsecond=0)

    # -- 未来时间确认 ----------------------------------------------
    future_times = []
    for key, t in [('create', create_time), ('modify', modify_time),
                    ('access', access_time)]:
        if t and t > now:
            future_times.append(f"{key}: {t.strftime('%Y-%m-%d %H:%M:%S')}")

    if future_times and not args.dry_run and not args.force:
        if not confirm_future_times(future_times, now):
            print(json.dumps({'status': 'cancelled'}, ensure_ascii=False))
            return

    # -- 构建输出 -------------------------------------------------
    result = {
        'status': 'ok',
        'file': args.file,
        'times': {
            'create': create_time.strftime('%Y-%m-%d %H:%M:%S'),
            'modify': modify_time.strftime('%Y-%m-%d %H:%M:%S'),
            'access': access_time.strftime('%Y-%m-%d %H:%M:%S'),
        },
        'edit_duration': {
            'requested': requested_duration,
            'actual_minutes': edit_minutes,
            'actual_seconds': edit_seconds,
        },
        'dry_run': args.dry_run
    }

    # 计算时间跨度
    if create_time and modify_time:
        span_min = int((modify_time - create_time).total_seconds() / 60)
        result['time_span'] = {
            'minutes': span_min,
            'description': f'创建到修改共 {span_min} 分钟'
        }
        # 逻辑一致性检查
        if edit_minutes and edit_minutes > span_min:
            warnings.append(
                f"编辑时长({edit_minutes}min) > 时间跨度({span_min}min)，"
                f"已调整为{span_min}min"
            )
            edit_minutes = span_min  # 不能超

    if warnings:
        result['warnings'] = warnings

    if args.dry_run:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    # -- 实际修改 -------------------------------------------------
    # 文件存在检查（非 dry-run 时执行）
    if not os.path.exists(args.file):
        print(json.dumps({'status': 'error',
                           'message': f'文件不存在: {args.file}'}, ensure_ascii=False))
        return

    try:
        close_office_processes()

        ext = os.path.splitext(args.file)[1].lower()
        if ext in ['.docx', '.pptx', '.xlsx'] and edit_minutes:
            ok = modify_office_internal(
                args.file, edit_minutes,
                create_dt=create_time, modify_dt=modify_time
            )
            result['office_internal'] = {
                'app_total_time': edit_minutes,
                'core_created': create_time.strftime('%Y-%m-%dT%H:%M:%SZ'),
                'core_modified': modify_time.strftime('%Y-%m-%dT%H:%M:%SZ'),
                'success': ok
            }

        success = set_file_system_times(
            args.file, create_time, modify_time, access_time)
        result['file_system'] = {'success': success}

    except Exception as e:
        result['status'] = 'error'
        result['message'] = str(e)

    print(json.dumps(result, ensure_ascii=False, indent=2))


def adjust_to_work_time(dt: datetime) -> datetime:
    """如果时间不在工作时间(08:00-22:00)内，调整到最近的次日08:00"""
    if 8 <= dt.hour < 22:
        return dt
    # 往前找：如果是深夜，移到次日08:00
    next_day = dt.date() if dt.hour >= 0 else (dt - timedelta(days=1)).date()
    return datetime.combine(next_day, datetime.min.time()).replace(hour=8)


def confirm_future_times(future_times: list, now: datetime) -> bool:
    """显示未来时间警告，返回用户是否确认继续"""
    msg = [
        "=" * 60,
        "\u26a0  警告：以下时间超过当前时间",
        "=" * 60,
    ]
    for ft in future_times:
        msg.append(f"  \u2022 {ft}")
    msg.append(f"\n当前时间：{now.strftime('%Y-%m-%d %H:%M:%S')}")
    msg.append("\n这可能是一个不合理的设置（文件来自未来？）")
    msg.append("-" * 60)
    msg.append("是否继续执行？(y/N): ")
    print('\n'.join(msg), end='')
    try:
        return input().strip().lower() == 'y'
    except (EOFError, IOError):
        return False


if __name__ == '__main__':
    main()
