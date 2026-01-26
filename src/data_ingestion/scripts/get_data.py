from data_ingestion.config.schemas import EQUIPMENT_CONFIG_ID

def get_data(build, unitName='all', ):        # start_date, end_date
    """
    调取数据、设备的信息
    
    :params build: int 栋
    :params unitName: str 单元，默认为'all'，表示所有单元
    """

    if unitName == 'all':
        print(f'正在下载{build}栋所有单元的数据')
        return EQUIPMENT_CONFIG_ID.loc[EQUIPMENT_CONFIG_ID['build'] == build, :]
    else:
        result = EQUIPMENT_CONFIG_ID.loc[
            (EQUIPMENT_CONFIG_ID['build'] == build) & 
            (EQUIPMENT_CONFIG_ID['unitName'] == unitName), 
            :
            ]
        if result.empty:
            print('未找到对应的设备！请检查楼宇单元是否正确')
            return None
        else:
            return result


from data_ingestion.config.api import HEADERS, EXP_FEED_DATA_URL
from data_ingestion.config.path import PATH_FEED_ORI
from tqdm import tqdm
import requests
def download_data(df_info, start_date, end_date):
    """
    下载数据，需要先获取到设备编号，然后根据设备编号下载数据

    :param df_info: 选择的设备信息表
    :param start_date: 开始日期
    :param end_date: 结束日期
    :return: 下载文件至指定目录 PATH_FEED_ORI
    """

    unit_nums = len(df_info)

    if unit_nums == 0:
        print('楼宇单元不存在，请检查楼宇单元是否正确！')
        return None
    else:
        for i in tqdm(range(unit_nums), desc='下载进度'):
            # print(f'正在下载第{i+1}个单元的数据...')

            # 设置文件名称，保证与原文件对齐
            unit_name = df_info.iloc[i]['unitName']
            file_name = f'育肥{unit_name}单元饲喂记录-{start_date}--{end_date}.xlsx'
            # print(file_name)

            # 获取参数表 area_id
            area_id = df_info.iloc[i]['areaId']

            # 发起API请求
            req_params = {
                "areaId": area_id,    
                "startDate": start_date,  
                "endDate": end_date,  
                "earTag": "",  
                "flag": 1
            }
            resp = requests.post(EXP_FEED_DATA_URL, headers=HEADERS, json=req_params)

            # 保存文件
            if resp.status_code != 200:
                raise ValueError(f"Response content is not successful: {file_name}")
            with open(PATH_FEED_ORI / file_name, 'wb') as f:
                f.write(resp.content)
            
        print(f'所有数据下载完成！')
    
if __name__ == '__main__':
    # print('请输入楼宇编号：')
    # build = int(input())        # type: int

    # print('请输入单元编号：')
    # unit = input()              # type: str

    ## 调试下载数据
    build = 6
    unit = 'all'

    # 选取指定单元的数据配置表
    select_df_info = get_data(build=build, unitName=unit)

    # 下载数据
    download_data(select_df_info, start_date='2025-08-01', end_date='2026-01-24')