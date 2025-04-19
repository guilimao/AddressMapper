import pandas as pd
import requests
import hashlib
import time
from urllib.parse import urlencode
import pandas as pd
import simplekml
from tqdm import tqdm

def excel_to_kml(input_excel, output_kml):
    """
    将包含经纬度信息的Excel文件转换为KML格式
    :param input_excel: 输入Excel文件路径
    :param output_kml: 输出KML文件路径
    """
    try:
        # 读取Excel数据
        df = pd.read_excel(input_excel, engine='openpyxl')
        
        # 创建KML对象
        kml = simplekml.Kml()
        
        # 统计有效数据
        valid_count = 0
        skipped = []
        
        # 进度条
        with tqdm(total=len(df), desc="转换进度") as pbar:
            for index, row in df.iterrows():
                try:
                    # 检查坐标有效性
                    lng = float(row['经度'])
                    lat = float(row['纬度'])
                    
                    # 创建地标
                    point = kml.newpoint(
                        name=row['公司名'],
                        description=f"""
                        负责人：{row['负责人姓名']}
                        地址：{row['公司地址']}
                        解析结果：{row.get('解析结果', '')}
                        """,
                        coords=[(lng, lat)]
                    )
                    
                    # 设置样式
                    point.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pal4/icon28.png'
                    valid_count += 1
                except (ValueError, KeyError):
                    skipped.append(index + 2)  # 记录Excel行号（含标题）
                pbar.update(1)
        
        # 保存KML文件
        kml.save(output_kml)
        
        print(f"\n转换完成！有效标记 {valid_count} 个")
        if skipped:
            print(f"跳过无效数据行（Excel行号）：{skipped}")
    
    except Exception as e:
        print(f"转换失败：{str(e)}")
        
def read_excel_data(file_path):
    """读取新版三列Excel文件（处理空负责人）"""
    columns = [
        '公司名',
        '公司地址',
        '负责人姓名'
    ]
    
    try:
        df = pd.read_excel(
            file_path,
            engine='openpyxl',
            header=None,
            names=columns,
            skiprows=0,
            keep_default_na=False  # 新增：不将空值转为NaN
        )
        # 将空字符串转换为None
        df = df.replace(r'^\s*$', None, regex=True)
        return df.to_dict(orient='records')
    
    except Exception as e:
        print(f"读取文件时发生错误: {str(e)}")
        return []

def get_coordinates(company):
    """调用存储在文本文件中的腾讯地图API（适配新版地址字段）"""
    try:
        with open('tencent_api.txt', 'r') as file:
            API_KEY = file.readline().strip()
            API_SECRET = file.readline().strip()
    except FileNotFoundError:
        print("API密钥文件未找到，请确保存在'tencent_api.txt'文件。")
        return None, None
    
    # 直接使用完整地址字段
    full_address = company['公司地址']
    
    params = {
        "address": full_address,
        "key": API_KEY
    }

    try:
        sorted_params = sorted(params.items(), key=lambda x: x[0])
        param_pairs = [f"{k}={v}" for k, v in sorted_params]
        raw_param_str = "&".join(param_pairs)
        
        sign_str = f"/ws/geocoder/v1?{raw_param_str}{API_SECRET}"
        sig = hashlib.md5(sign_str.encode('utf-8')).hexdigest().lower()
        
        encoded_params = []
        for k, v in sorted_params:
            encoded_value = requests.utils.quote(v, safe='')
            encoded_params.append(f"{k}={encoded_value}")
        
        encoded_params.append(f"sig={sig}")
        url = f"https://apis.map.qq.com/ws/geocoder/v1?{'&'.join(encoded_params)}"
        
        headers = {'x-legacy-url-decode': 'no'}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        
        if data.get('status') == 0:
            location = data['result']['location']
            return {
                '经度': location['lng'],
                '纬度': location['lat'],
                '解析结果': data['result']['title']
            }
        else:
            return {
                '经度': '',
                '纬度': '',
                '解析结果': f"错误：{data.get('message', '未知错误')}"
            }
    
    except Exception as e:
        print(f"请求失败：{str(e)}")
        return {
            '经度': '',
            '纬度': '',
            '解析结果': f"异常：{str(e)}"
        }

def write_to_excel(data, output_path):
    """改进导出函数（处理空负责人显示）"""
    try:
        df = pd.DataFrame(data)
        # 将None转换为空字符串
        df['负责人姓名'] = df['负责人姓名'].fillna('')
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"数据已成功写入到 {output_path}")
    except Exception as e:
        print(f"写入失败：{str(e)}")


if __name__ == "__main__":
    input_file = "1.xlsx"
    output_file = "Address_With_GPS.xlsx"
    batch_size = 20

    companies = read_excel_data(input_file)
    
    if not companies:
        print("没有读取到有效数据，请检查输入文件")
        exit()
    
    print(f"开始处理 {len(companies)} 条数据...")
    
    try:
        for idx, company in enumerate(companies, 1):
            print(f"正在处理第 {idx} 条 ({idx/len(companies):.1%})", end='\r')
            
            coordinates = get_coordinates(company)
            company.update(coordinates)
            
            if idx % batch_size == 0:
                write_to_excel(companies, output_file)
                print(f"\n✅ 已安全保存前 {idx} 条数据")
            
            time.sleep(0.2)
        
        if len(companies) % batch_size != 0:
            write_to_excel(companies, output_file)
            print(f"\n✅ 最终保存完成（共 {len(companies)} 条）")
        

        input_excel = output_file
        output_kml = "Address_With_GPS.kml"
        excel_to_kml(input_excel, output_kml)
        print("\n处理完成！结果文件已保存，请在Google Earth中导入KML文件")

    except KeyboardInterrupt:
        write_to_excel(companies, output_file)
        print(f"\n⚠️ 用户中断！已保存已处理的 {len(companies)} 条数据")
    except Exception as e:
        write_to_excel(companies, output_file)
        print(f"\n⚠️ 发生异常 {str(e)}！已保存已处理的 {len(companies)} 条数据")
        raise