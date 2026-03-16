""""
平台字段定义

"""

from data_ingestion.config.path import PATH_ROOT, PATH_SRC_INGEST
import pandas as pd

EQUIPMENT_CONFIG_ID = pd.read_excel(PATH_SRC_INGEST / "config" /"gateway_config.xlsx", header=1).dropna()    # 删除中间空行


## ------ 调试 ---------- 打印当前保存设备信息
# print('当前配置设备信息：')
# print(EQUIPMENT_CONFIG_ID)



