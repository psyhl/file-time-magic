# -*- coding: utf-8 -*-
"""
file-time-magic 鈥?鏂囦欢鏃堕棿灞炴€т慨鏀硅剼鏈?v1.1

鍔熻兘锛?1. 瑙ｆ瀽鐢ㄦ埛杈撳叆鐨勭紪杈戞椂闀匡紙鏀寔澶氱鏍煎紡锛?2. 鏀寔鐢ㄦ埛鎸囧畾鍒涘缓鏃堕棿銆佷慨鏀规椂闂?3. 璁＄畻闅忔満鍖栫殑鏃堕棿锛堝垎閽熸湁璇樊锛岀鏁伴殢鏈猴級
4. 淇敼 Office 鏂囦欢鍐呴儴 XML锛圱otalTime 灞炴€?+ core.xml 鍒涘缓/淇敼鏃堕棿锛?5. 璁剧疆鏂囦欢绯荤粺鏃堕棿灞炴€?
浣跨敤锛?  python set_file_time.py --file "鏂囦欢璺緞" --edit-duration "2灏忔椂"
  python set_file_time.py --file "鏂囦欢璺緞" --create-time "2024-01-15 09:30:00" --edit-duration "3灏忔椂"
  python set_file_time.py --file "鏂囦欢璺緞" --modify-time "2026-04-18 14:00:00" --edit-duration "90鍒嗛挓"
  python set_file_time.py --file "鏂囦欢璺緞" --create-time "2024-01-15 09:00:00" --modify-time "2026-04-18 14:00:00"
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
    瑙ｆ瀽鐢ㄦ埛杈撳叆鐨勬椂闀匡紝杩斿洖鍒嗛挓鏁帮紙鏁存暟锛?
    鏀寔鏍煎紡锛?    - "2灏忔椂"銆?2灏忔椂30鍒?銆?涓ゅ皬鏃跺崐"
    - "120鍒嗛挓"銆?90"锛堢函鏁板瓧榛樿鍒嗛挓锛?    - "3h"銆?2h30m"銆?150m"
    - "1.5灏忔椂"
    """
    text = text.strip().lower()

    # 鏇挎崲涓枃
    text = text.replace("涓?, "2")
    text = text.replace("鍗?, ".5")
    text = text.replace("灏忔椂", "h")
    text = text.replace("鍒嗛挓", "m")
    text = text.replace("鍒?, "m")
    text = text.replace("鏃?, "h")

    total_minutes = 0

    # 鍏堝尮閰?"Xh Ym" 鎴?"XhYm" 绛夊鍚堟牸寮?    # 渚嬪 "2h30m", "1.5h", "30m"
    # 鍖归厤鎵€鏈夋暟瀛?鍗曚綅缁勫悎
    pattern = r'([\d.]+)\s*h'
    h_matches = re.findall(pattern, text)
    for val in h_matches:
        total_minutes += int(float(val) * 60)

    pattern = r'([\d.]+)\s*m(?!o)'  # m 涓嶅尮閰?"mo" 寮€澶?    m_matches = re.findall(pattern, text)
    for val in m_matches:
        total_minutes += int(float(val))

    # 濡傛灉娌″尮閰嶅埌浠讳綍鍐呭锛屽皾璇曠函鏁板瓧锛堥粯璁ゅ垎閽燂級
    if total_minutes == 0:
        try:
            total_minutes = int(float(text))
        except:
            pass

    return total_minutes


def parse_time_str(text: str) -> datetime:
    """
    瑙ｆ瀽鏃堕棿瀛楃涓诧紝鏀寔澶氱鏍煎紡
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

    raise ValueError(f"鏃犳硶瑙ｆ瀽鏃堕棿: {text}")


def randomize_duration(minutes: int) -> tuple:
    """闅忔満鍖栨椂闀匡紝杩斿洖 (瀹為檯鍒嗛挓鏁? 绉掓暟)"""
    if minutes >= 60:
        actual_minutes = minutes + random.randint(-5, 5)
    else:
        actual_minutes = minutes + random.randint(-3, 3)
    actual_minutes = max(actual_minutes, 1)
    actual_seconds = random.randint(0, 59)
    return actual_minutes, actual_seconds


def is_work_hour(hour: int) -> bool:
    """妫€鏌ユ槸鍚﹀湪宸ヤ綔鏃堕棿锛?8:00 - 22:00锛?""
    return 8 <= hour <= 22


def adjust_to_work_time(dt: datetime) -> datetime:
    """灏嗘椂闂磋皟鏁村埌宸ヤ綔鏃堕棿鑼冨洿"""
    if dt.hour < 8:
        return dt.replace(hour=random.randint(8, 10), minute=random.randint(0, 59))
    elif dt.hour > 22:
        return (dt - timedelta(days=1)).replace(hour=random.randint(20, 22), minute=random.randint(0, 59))
    return dt


def add_random_seconds(dt: datetime) -> datetime:
    """娣诲姞闅忔満绉掓暟"""
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
    璁＄畻鎵€鏈夋椂闂村睘鎬?    """
    if base_time is None:
        base_time = datetime.now()

    result = {
        'create': None,
        'modify': None,
        'access': None,
        'edit_minutes': edit_minutes,
        'edit_seconds': edit_seconds if edit_seconds is not None else random.randint(0, 59)
    }

    # 鎯呭喌1锛氱敤鎴锋寚瀹氫簡鍒涘缓鏃堕棿鍜屼慨鏀规椂闂?    if create_time and modify_time:
        result['create'] = add_random_seconds(create_time)
        result['modify'] = add_random_seconds(modify_time)
        if edit_minutes is None:
            duration = result['modify'] - result['create']
            result['edit_minutes'] = int(duration.total_seconds() / 60)
            result['edit_seconds'] = random.randint(0, 59)

    # 鎯呭喌2锛氱敤鎴锋寚瀹氫簡鍒涘缓鏃堕棿鍜岀紪杈戞椂闀?    elif create_time and edit_minutes:
        result['create'] = add_random_seconds(create_time)
        result['modify'] = add_random_seconds(
            result['create'] + timedelta(minutes=edit_minutes + random.randint(0, 2))
        )

    # 鎯呭喌3锛氱敤鎴锋寚瀹氫簡淇敼鏃堕棿鍜岀紪杈戞椂闀?    elif modify_time and edit_minutes:
        result['modify'] = add_random_seconds(modify_time)
        calc_create = result['modify'] - timedelta(minutes=edit_minutes + random.randint(0, 2))
        result['create'] = add_random_seconds(adjust_to_work_time(calc_create))

    # 鎯呭喌4锛氱敤鎴峰彧鎸囧畾浜嗙紪杈戞椂闀?    elif edit_minutes:
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

    # 鎯呭喌5锛氫粈涔堥兘娌℃寚瀹?    else:
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

    # 璁块棶鏃堕棿
    if access_time:
        result['access'] = add_random_seconds(access_time)
    else:
        result['access'] = add_random_seconds(
            result['modify'] + timedelta(minutes=random.randint(3, 15))
        )

    # 纭繚鏃堕棿閫昏緫姝ｇ‘
    if result['create'] > result['modify']:
        result['modify'] = result['create'] + timedelta(minutes=result['edit_minutes'] or 30)
    if result['modify'] > result['access']:
        result['access'] = result['modify'] + timedelta(minutes=random.randint(3, 15))

    # 鑷姩璁＄畻妯″紡涓嬶細纭繚涓嶈秴鍑哄熀鍑嗘椂闂?    user_specified_times = create_time is not None or modify_time is not None
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

    # 楠岃瘉锛氱紪杈戞椂闀夸笉瓒呰繃鏂囦欢瀛樺湪鏃堕棿
    file_exist_minutes = int((result['modify'] - result['create']).total_seconds() / 60)
    if result['edit_minutes'] and result['edit_minutes'] > file_exist_minutes:
        result['edit_minutes'] = max(1, file_exist_minutes - random.randint(1, 5))
        result['edit_seconds'] = random.randint(0, 59)

    # 闈炲伐浣滄椂闂磋鍛?    if result['create'] and not is_work_hour(result['create'].hour):
        result['work_time_warning'] = (
            f"鍒涘缓鏃堕棿 {result['create'].strftime('%H:%M')} "
            f"涓嶅湪甯歌宸ヤ綔鏃堕棿锛?8:00-22:00锛夛紝宸茶嚜鍔ㄨ皟鏁?
        )

    return result


def modify_office_internal(file_path: str, edit_minutes: int,
                             create_dt: datetime = None,
                             modify_dt: datetime = None) -> bool:
    """
    淇敼 Office 鏂囦欢锛坉ocx/pptx/xlsx锛夊唴閮ㄥ睘鎬?    - TotalTime锛堢紪杈戞椂闀匡級锛氭潵鑷?docProps/app.xml
    - 鍒涘缓鏃堕棿銆佷慨鏀规椂闂达細鏉ヨ嚜 docProps/core.xml

    Args:
        file_path:  鏂囦欢璺緞
        edit_minutes:  缂栬緫鏃堕暱锛堝垎閽燂級
        create_dt:     鍒涘缓鏃堕棿锛坉atetime锛?        modify_dt:     淇敼鏃堕棿锛坉atetime锛?    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ['.docx', '.pptx', '.xlsx']:
        return False

    # 浣跨敤 PID 鍖哄垎涓存椂鐩綍锛岄伩鍏嶅苟鍙戝啿绐?    tmp_dir = os.path.join(os.environ.get('TEMP', '/tmp'),
                           '_ftm_' + str(os.getpid()) + '_' + str(random.randint(1000, 9999)))

    # 娓呯悊鏃ф畫鐣?    if os.path.exists(tmp_dir):
        try:
            shutil.rmtree(tmp_dir)
        except Exception:
            pass
    os.makedirs(tmp_dir)

    try:
        # 瑙ｅ帇
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

        # === core.xml: 鍒涘缓鏃堕棿 + 淇敼鏃堕棿 ===
        core_xml = os.path.join(tmp_dir, 'docProps', 'core.xml')
        if os.path.exists(core_xml) and create_dt and modify_dt:
            tree = ET.parse(core_xml)
            root = tree.getroot()

            xsi_ns = 'http://www.w3.org/2001/XMLSchema-instance'
            dct_ns = 'http://purl.org/dc/terms/'

            create_utc = create_dt.strftime('%Y-%m-%dT%H:%M:%SZ')
            modify_utc = modify_dt.strftime('%Y-%m-%dT%H:%M:%SZ')

            # 鏇存柊 dcterms:created锛圵ord灞炴€у璇濇涓殑"鍒涘缓鏃堕棿"锛?            for el in root.iter(dct_ns + 'created'):
                el.set(f'{{{xsi_ns}}}type', 'dcterms:W3CDTF')
                el.text = create_utc

            # 鏇存柊 dcterms:modified锛圵ord灞炴€у璇濇涓殑"淇敼鏃堕棿"锛?            for el in root.iter(dct_ns + 'modified'):
                el.set(f'{{{xsi_ns}}}type', 'dcterms:W3CDTF')
                el.text = modify_utc

            tree.write(core_xml, xml_declaration=True, encoding='UTF-8')

        # 閲嶆柊鎵撳寘
        new_file = file_path + '.tmp'
        with zipfile.ZipFile(new_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for rd, dirs, files in os.walk(tmp_dir):
                for fn in files:
                    fp = os.path.join(rd, fn)
                    z.write(fp, os.path.relpath(fp, tmp_dir))

        # 鏇挎崲鍘熸枃浠?        os.remove(file_path)
        os.rename(new_file, file_path)

    finally:
        # 娓呯悊涓存椂鐩綍
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
    璁剧疆鏂囦欢绯荤粺鏃堕棿灞炴€?    - 鏂囦欢澶癸細浣跨敤 Python os.utime + Windows API锛圫etFileTime锛?    - 鏂囦欢锛氫娇鐢?os.utime锛堝厛灏濊瘯锛屽け璐ュ垯鐢?PowerShell锛?    """
    if os.path.isdir(file_path):
        return _set_folder_times(file_path, create, modify, access)
    else:
        return _set_file_times(file_path, create, modify, access)


def _utc_offset_hours() -> int:
    """鑾峰彇褰撳墠鏃跺尯涓嶶TC鐨勫亸绉婚噺锛堝皬鏃讹級锛學indows涓撶敤"""
    # time.timezone 鏄粠鏈湴鏃堕棿寰€UTC璧扮殑鍋忕Щ锛堢锛夛紝west涓烘
    # time.altzone 鏄浠ゆ椂鍋忕Щ锛堝鏋滄湁锛?    # UTC+8: time.timezone = -28800 (8*3600), time.altzone = -28800
    offset_sec = time.timezone if time.daylight == 0 else time.altzone
    return -offset_sec // 3600  # UTC+8 鈫?+8


def _set_file_times(file_path: str,
                     create: datetime,
                     modify: datetime,
                     access: datetime) -> bool:
    """
    鐢?Python os 璁剧疆鏂囦欢鏃堕棿锛堣法骞冲彴锛屼笉渚濊禆 PowerShell锛?    """
    try:
        import ctypes
        from ctypes import wintypes

        # os.utime 璁剧疆 modify 鍜?access锛堣繖涓や釜鐢?datetime.timestamp() 姝ｇ‘杞崲锛?        access_ts = access.timestamp()
        modify_ts = modify.timestamp()
        os.utime(file_path, (access_ts, modify_ts))

        # 鍒涘缓鏃堕棿鐢?Windows API SetFileTime
        # 娉ㄦ剰锛歋etFileTime 鎺ユ敹 UTC 鏃堕棿
        # 杈撳叆鐨?datetime 鏄湰鍦版椂闂达紙CST锛夛紝闇€瑕佽浆 UTC
        # Windows FILETIME = 100-ns ticks since 1601-01-01 UTC
        kernel32 = ctypes.windll.kernel32

        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD),
                        ('dwHighDateTime', wintypes.DWORD)]

        def dt_to_filetime_utc(dt: datetime) -> FILETIME:
            # 灏嗘湰鍦版椂闂磋浆涓?UTC 鍐嶇畻 FILETIME
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
    """璁剧疆鏂囦欢澶规椂闂村睘鎬э紙Windows 涓撶敤锛?""
    try:
        import ctypes
        from ctypes import wintypes

        # 淇敼鏃堕棿鍜岃闂椂闂寸敤 os.utime
        os.utime(folder_path, (access.timestamp(), modify.timestamp()))

        # 鍒涘缓鏃堕棿鐢?Windows API SetFileTime锛堥渶瑕?UTC锛?        kernel32 = ctypes.windll.kernel32

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
    """鍏抽棴鍙兘鍗犵敤鏂囦欢鐨?Office 杩涚▼"""
    for proc in ['WINWORD', 'EXCEL', 'POWERPNT']:
        try:
            subprocess.run(['taskkill', '/F', '/IM', f'{proc}.EXE'],
                           capture_output=True, timeout=5)
        except Exception:
            pass


def confirm_future_times(future_times: list, now: datetime) -> bool:
    """纭鏄惁璁剧疆鏈潵鏃堕棿"""
    print("\n" + "=" * 60)
    print("鈿狅笍  璀﹀憡锛氫互涓嬫椂闂磋秴杩囧綋鍓嶆椂闂?)
    print("=" * 60)
    for t in future_times:
        print(f"  鈥?{t}")
    print(f"\n褰撳墠鏃堕棿锛歿now.strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 60)
    try:
        resp = input("鏄惁缁х画鎵ц锛?y/N): ").strip().lower()
        return resp in ['y', 'yes']
    except Exception:
        print("闈炰氦浜掓ā寮忥紝宸插彇娑堛€備娇鐢?--force 寮哄埗鎵ц銆?)
        return False


def main():
    parser = argparse.ArgumentParser(description='鏂囦欢鏃堕棿灞炴€т慨鏀?v1.1')
    parser.add_argument('--file', required=True, help='鐩爣鏂囦欢璺緞')
    parser.add_argument('--edit-duration', help='缂栬緫鏃堕暱锛屽 "2灏忔椂"銆?120鍒嗛挓"')
    parser.add_argument('--create-time', help='鍒涘缓鏃堕棿锛屽 "2024-01-15 09:30:00"')
    parser.add_argument('--modify-time', help='淇敼鏃堕棿锛屽 "2026-04-18 14:00:00"')
    parser.add_argument('--access-time', help='璁块棶鏃堕棿锛堝彲閫夛級')
    parser.add_argument('--base-time', help='鍙傝€冩椂闂达紝濡?"2026-04-18 14:30:00"')
    parser.add_argument('--dry-run', action='store_true', help='鍙樉绀鸿绠楃粨鏋滐紝涓嶅疄闄呬慨鏀?)
    parser.add_argument('--force', action='store_true', help='寮哄埗鎵ц锛岃烦杩囨湭鏉ユ椂闂寸‘璁?)

    args = parser.parse_args()

    if not os.path.exists(args.file):
        print(json.dumps({'status': 'error',
                           'message': f'鏂囦欢涓嶅瓨鍦? {args.file}'}, ensure_ascii=False))
        return

    # 瑙ｆ瀽缂栬緫鏃堕暱
    edit_minutes = None
    edit_seconds = None
    requested_duration = None

    if args.edit_duration:
        requested_duration = args.edit_duration
        parsed = parse_duration(args.edit_duration)
        if parsed > 0:
            edit_minutes, edit_seconds = randomize_duration(parsed)

    # 瑙ｆ瀽鏃堕棿
    create_time = modify_time = access_time = base_time = None
    if args.create_time:
        try:
            create_time = parse_time_str(args.create_time)
        except ValueError as e:
            print(json.dumps({'status': 'error',
                               'message': f'鍒涘缓鏃堕棿鏍煎紡閿欒: {args.create_time}'}, ensure_ascii=False))
            return
    if args.modify_time:
        try:
            modify_time = parse_time_str(args.modify_time)
        except ValueError as e:
            print(json.dumps({'status': 'error',
                               'message': f'淇敼鏃堕棿鏍煎紡閿欒: {args.modify_time}'}, ensure_ascii=False))
            return
    if args.access_time:
        try:
            access_time = parse_time_str(args.access_time)
        except ValueError as e:
            print(json.dumps({'status': 'error',
                               'message': f'璁块棶鏃堕棿鏍煎紡閿欒: {args.access_time}'}, ensure_ascii=False))
            return
    if args.base_time:
        try:
            base_time = parse_time_str(args.base_time)
        except ValueError as e:
            print(json.dumps({'status': 'error',
                               'message': f'鍩哄噯鏃堕棿鏍煎紡閿欒: {args.base_time}'}, ensure_ascii=False))
            return

    # 璁＄畻鏃堕棿
    times = calculate_times_v2(
        edit_minutes=edit_minutes,
        edit_seconds=edit_seconds,
        create_time=create_time,
        modify_time=modify_time,
        access_time=access_time,
        base_time=base_time
    )

    # 鏈潵鏃堕棿纭
    now = datetime.now()
    future_times = []
    for key in ['create', 'modify', 'access']:
        t = times[key]
        if t and t > now:
            future_times.append(f"{key}: {t.strftime('%Y-%m-%d %H:%M:%S')}")

    if future_times and not args.dry_run and not args.force:
        if not confirm_future_times(future_times, now):
            print(json.dumps({'status': 'cancelled',
                               'message': '鐢ㄦ埛鍙栨秷鎿嶄綔'}, ensure_ascii=False))
            return

    # 鏋勫缓杈撳嚭
    result = {
        'status': 'ok',
        'file': args.file,
        'times': {
            k: times[k].strftime('%Y-%m-%d %H:%M:%S')
            for k in ['create', 'modify', 'access']
        },
        'edit_duration': {
            'requested': requested_duration,
            'parsed_minutes': edit_minutes,
            'actual_minutes': times['edit_minutes'],
            'actual_seconds': times['edit_seconds']
        },
        'dry_run': args.dry_run
    }

    warnings = []
    if 'work_time_warning' in times:
        warnings.append(times['work_time_warning'])

    file_exist_min = int((times['modify'] - times['create']).total_seconds() / 60)
    if edit_minutes and edit_minutes > file_exist_min:
        warnings.append(
            f"缂栬緫鏃堕暱({edit_minutes}鍒嗛挓)瓒呰繃鏂囦欢瀛樺湪鏃堕棿({file_exist_min}鍒嗛挓)锛?
            f"宸茶皟鏁翠负{times['edit_minutes']}鍒嗛挓"
        )

    if warnings:
        result['warnings'] = warnings

    if args.dry_run:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    # 瀹為檯淇敼
    try:
        close_office_processes()

        ext = os.path.splitext(args.file)[1].lower()
        if ext in ['.docx', '.pptx', '.xlsx'] and times['edit_minutes']:
            ok = modify_office_internal(
                args.file, times['edit_minutes'],
                create_dt=times['create'], modify_dt=times['modify']
            )
            result['office_internal'] = {
                'app_total_time': times['edit_minutes'],
                'core_created': times['create'].strftime('%Y-%m-%dT%H:%M:%SZ'),
                'core_modified': times['modify'].strftime('%Y-%m-%dT%H:%M:%SZ'),
                'success': ok
            }

        success = set_file_system_times(
            args.file, times['create'], times['modify'], times['access'])
        result['file_system'] = {'success': success}
        if not success:
            result['status'] = 'warning'
            result['message'] = '鏂囦欢绯荤粺鏃堕棿璁剧疆鍙兘鏈畬鍏ㄦ垚鍔?

    except Exception as e:
        result['status'] = 'error'
        result['message'] = str(e)

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
