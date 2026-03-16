import re
import pandas as pd
from dataclasses import dataclass

@dataclass
class FeedFileMeta:
    """
    解析生猪饲喂记录文件名的元数据类。

    功能：
        - 从文件名中提取饲喂阶段、单元编号、饲喂起始与结束日期
        - 验证日期顺序合法性
        - 提供饲喂周期（天数）计算
        - 支持导出字典形式便于 DataFrame 或数据库使用

    文件名规范示例：
        育肥3-4单元饲喂记录-2025-12-01--2026-01-12

    属性：
        filename (str): 原始文件名
        stage (str): 阶段名称，如 "育肥"
        unit (str): 单元编号，如 "3-4"
        start_date (datetime.date): 饲喂开始日期
        end_date (datetime.date): 饲喂结束日期

    方法：
        duration_days: 属性方法，返回饲喂周期天数（包含首尾日期）
        to_dict(): 返回类属性字典形式，便于存储或记录
    """
    
    filename: str

    stage: str = None
    unit: str = None
    start_date: pd.Timestamp = None
    end_date: pd.Timestamp = None

    def __post_init__(self):
        self._parse_filename()
        self.validate()

    def _parse_filename(self):
        pattern = (
            r"(?P<stage>育肥)"
            r"(?P<unit>\d+-\d+)单元饲喂记录-"
            r"(?P<start_date>\d{4}-\d{2}-\d{2})--"
            r"(?P<end_date>\d{4}-\d{2}-\d{2})"
        )
        
        m = re.search(pattern, self.filename)
        if not m:
            raise ValueError(f'文件名格式不符合规范：{self.filename}')
        
        self.stage = m.group("stage")
        self.unit = m.group("unit")
        self.start_date = pd.to_datetime(m.group("start_date")).date()
        self.end_date = pd.to_datetime(m.group("end_date")).date()

    def validate(self):
        if self.start_date > self.end_date:
            raise ValueError("开始日期不能晚于结束日期")
        
    @property
    def duration_days(self):
        return (self.end_date - self.start_date).days + 1
    
    def to_dict(self):
        return {
            "stage": self.stage,
            "unit": self.unit,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "duration_days": self.duration_days,
        }
        
