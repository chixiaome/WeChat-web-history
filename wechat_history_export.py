# -*- coding: utf-8 -*-
import os
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
import logging
import glob

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def convert_webkit_timestamp(webkit_timestamp):
    """
    转换 WebKit 时间戳为 datetime
    时间戳格式：从 1601-01-01 00:00:00 UTC 开始的微秒数
    """
    try:
        # 如果时间戳为空或0，返回None
        if not webkit_timestamp:
            return None
            
        # 处理负数时间戳
        if webkit_timestamp < 0:
            logging.warning(f"发现负数时间戳: {webkit_timestamp}")
            return None
            
        try:
            # 示例时间戳：13380105538768906
            # 1. 首先转换为秒
            microseconds = int(webkit_timestamp)
            seconds = microseconds // 1000000
            
            # 2. 计算从1601年到1970年的秒数
            seconds_from_1601_to_1970 = 11644473600
            
            # 3. 转换为Unix时间戳
            unix_timestamp = seconds - seconds_from_1601_to_1970
            
            # 4. 转换为datetime
            dt = datetime.fromtimestamp(unix_timestamp)
            
            # 5. 添加微秒部分
            remaining_microseconds = microseconds % 1000000
            dt = dt.replace(microsecond=remaining_microseconds)
            
            return dt
            
        except Exception as e:
            logging.warning(f"无法解析时间戳 {webkit_timestamp}: {str(e)}")
            return None
        
    except Exception as e:
        logging.warning(f"时间戳转换失败 ({webkit_timestamp}): {str(e)}")
        return None

def format_datetime(dt):
    """格式化日期时间为易读格式"""
    if pd.isna(dt):
        return ""
    try:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return ""

def get_possible_wechat_paths():
    """获取可能的微信浏览器数据路径"""
    user_profile = os.environ['USERPROFILE']
    
    # 基础路径
    base_path = os.path.join(user_profile, 'AppData', 'Roaming', 'Tencent', 'WeChat', 'radium', 'web', 'profiles')
    
    # 使用通配符查找所有可能的配置文件夹
    profile_paths = glob.glob(os.path.join(base_path, 'multitab_*'))
    
    if not profile_paths:
        logging.warning(f"在 {base_path} 中未找到配置文件夹")
        
    return profile_paths

def get_history_file_paths(profile_path):
    """获取历史记录文件路径"""
    history_path = os.path.join(profile_path, 'history')
    return [history_path] if os.path.exists(history_path) else []

def get_wechat_history() -> str:
    """
    导出微信浏览器历史记录到脚本所在目录的output文件夹
    
    Returns:
        str: 输出文件路径
    """
    try:
        # 获取脚本所在目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, 'output')
        
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 获取所有可能的配置文件夹
        profile_paths = get_possible_wechat_paths()
        if not profile_paths:
            raise FileNotFoundError("未找到微信浏览器配置目录")
        
        all_history = []
        
        # 定义列名映射
        column_names = {
            'url': '网址',
            'title': '标题',
            'last_visit_time': '最后访问时间',
            'visit_count': '访问次数',
            'visit_time': '访问时间',
            'from_visit': '来源访问',
            'transition': '跳转类型',
            'profile_id': '配置ID',
            'source_path': '源文件路径'
        }
        
        for profile_path in profile_paths:
            logging.info(f"正在检查配置目录: {profile_path}")
            profile_id = os.path.basename(profile_path)
            
            history_paths = get_history_file_paths(profile_path)
            
            if not history_paths:
                logging.warning(f"在配置 {profile_id} 中未找到历史记录文件")
                continue
            
            for db_path in history_paths:
                logging.info(f"尝试读取历史记录: {db_path}")
                
                try:
                    # 创建数据库副本以避免文件锁定
                    temp_db = os.path.join(output_dir, f'temp_history_{datetime.now().strftime("%H%M%S")}.db')
                    with open(db_path, 'rb') as f_in:
                        with open(temp_db, 'wb') as f_out:
                            f_out.write(f_in.read())
                    
                    # 读取历史记录
                    conn = sqlite3.connect(temp_db)
                    try:
                        # 尝试获取表结构
                        tables_query = "SELECT name FROM sqlite_master WHERE type='table'"
                        tables = pd.read_sql_query(tables_query, conn)
                        logging.info(f"数据库中的表: {tables['name'].tolist()}")
                        
                        # 尝试读取历史记录
                        query = """
                        SELECT 
                            u.url, 
                            u.title, 
                            u.last_visit_time,
                            u.visit_count,
                            v.visit_time,
                            v.from_visit,
                            v.transition
                        FROM urls u
                        LEFT JOIN visits v ON u.id = v.url
                        ORDER BY u.last_visit_time DESC
                        """
                        df = pd.read_sql_query(query, conn)
                        
                        # 转换时间戳前，先看看原始值
                        logging.info(f"时间戳示例: {df['last_visit_time'].head().tolist()}")
                        logging.info(f"访问时间示例: {df['visit_time'].head().tolist()}")
                        
                        # 转换时间戳
                        df['last_visit_time'] = df['last_visit_time'].apply(convert_webkit_timestamp)
                        df['visit_time'] = df['visit_time'].apply(convert_webkit_timestamp)
                        
                    except Exception as e:
                        logging.error(f"读取数据库时出错: {str(e)}")
                        # 如果出错，尝试获取表的具体结构
                        for table in tables['name']:
                            try:
                                schema_query = f"PRAGMA table_info({table})"
                                schema = pd.read_sql_query(schema_query, conn)
                                logging.info(f"表 {table} 的结构: {schema.to_dict()}")
                            except:
                                pass
                        raise
                    finally:
                        conn.close()
                    
                    # 删除临时文件
                    os.remove(temp_db)
                    
                    # 添加来源信息
                    df['profile_id'] = profile_id
                    df['source_path'] = db_path
                    all_history.append(df)
                    logging.info(f"成功读取历史记录: {len(df)} 条记录")
                    
                except Exception as e:
                    logging.error(f"处理历史记录文件 {db_path} 时出错: {str(e)}")
                    if os.path.exists(temp_db):
                        os.remove(temp_db)
                    continue
        
        if not all_history:
            raise ValueError("未能成功读取任何历史记录")
            
        # 合并所有历史记录
        result = pd.concat(all_history, ignore_index=True)
        
        # 格式化时间列
        result['last_visit_time'] = result['last_visit_time'].apply(format_datetime)
        result['visit_time'] = result['visit_time'].apply(format_datetime)
        
        # 重命名列名为中文
        result = result.rename(columns=column_names)
        
        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(
            output_dir,
            f'wechat_history_{timestamp}.xlsx'
        )
        
        # 使用 ExcelWriter 保存，以设置编码和格式
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            result.to_excel(writer, index=False)
            # 调整列宽
            worksheet = writer.sheets['Sheet1']
            for idx, col in enumerate(result.columns):
                max_length = max(
                    result[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
        
        logging.info(f"历史记录已导出至: {output_file}")
        return output_file
            
    except Exception as e:
        logging.error(f"导出历史记录时发生错误: {str(e)}")
        raise

if __name__ == '__main__':
    try:
        print("注意: 请确保已关闭微信,否则可能无法访问历史记录文件")
        output_file = get_wechat_history()
        print(f"历史记录已导出至: {output_file}")
    except Exception as e:
        print(f"程序执行出错: {str(e)}")