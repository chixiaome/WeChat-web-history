# WeChat Web History Export
一个用于导出微信 PC 版网页浏览历史记录的 Python 工具。
## 功能特点
- 自动查找并导出微信网页浏览历史
 支持多用户配置
 导出格式为 Excel，包含详细信息
 自动处理时间戳转换
 使用中文列名，方便查看
## 安装步骤
1. 确保已安装 Python 3.7 或更高版本
2. 安装所需依赖：pip install pandas openpyxl

## 使用方法
1. 关闭微信 PC 版
2. 运行脚本：python wechat_history_export.py
3. 导出的历史记录将保存在脚本所在目录的 `output` 文件夹中
4. 文件名格式：`wechat_history_YYYYMMDD_HHMMSS.xlsx`

## 输出内容
导出的 Excel 文件包含以下信息：
- 网址
- 标题
- 最后访问时间
- 访问次数
- 访问时间
- 来源访问
- 跳转类型
- 配置ID
- 源文件路径

