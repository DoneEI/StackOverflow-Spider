# Stack Overflow 网页数据爬取 1.0

## 项目说明
该项目支持绝大多数Stack Overflow网页的数据爬取，包括标题，问题内容，问题评论，问题回答以及问题回答相应评论。

需要注意的是1.0版本只支持爬取网页纯文本内容，即无法爬取包括a标签内链接、img标签对应图片等信息。未来可能会对此进行优化，所以项目目前只能针对这种情形对用户进行提醒说明（见后续配置说明）。

## 使用说明
目录中requirement.txt文件提供了运行项目所需要的第三方库以及对应版本，使用"pip install - r requirement.txt"命令即可快速配置运行环境。进行简单的配置后运行spy.py方法即可开始爬取。
爬取结果存放在'当前运行环境路径/Extracted Documents'下。

## 配置说明
项目可以进行一些简单的配置来满足用户的不同需求(见spy.py文件中baseConfig方法),目前支持的配置有：
    
    # url文档路径, 只支持一个excel文件
    config['COLLECTED_URL_EXCEL_FILE_PATH']
    
    # 爬取一个excel文件的文档数量,即需要爬取的该excel文件工作表个数
    config['NUMBER_OF_SPY_WORK'] = 1

    # url所在文档工作表索引, 若其长度与config['NUMBER_OF_SPY_WORK']不符，则报错
    config['COLLECTED_URL_EXCEL_FILE_SHEET']

    # url所在文档工作表的列, 若其长度与config['NUMBER_OF_SPY_WORK']不符，则会在末尾填充默认值0
    config['COLLECTED_URL_EXCEL_FILE_COL'] = [0]

    # 爬取url起始行, 若其长度与config['NUMBER_OF_SPY_WORK']不符，则会在末尾填充默认值0
    config['COLLECTED_URL_EXCEL_FILE_START_ROW'] = [0]

    # 爬取url结束行, 若其长度与config['NUMBER_OF_SPY_WORK']不符，则会在末尾填充默认值-1（表示文档末尾)
    config['COLLECTED_URL_EXCEL_FILE_END_ROW'] = [-1]

    # 保存爬取文件的文件名(没有则按工作表名命名)
    config['COLLECTED_DATA_WORD_FILE_NAME'] = ['demo.docx']

    # 保存爬取文件的文件名
    config['COLLECTED_DATA_WORD_FILE_NAME']

    # 对于无法处理且需要警告的标签，即当某段落中出现这类标签时虽无法正确爬取但会出现警告信息
    config['UNPROCESSABLE_TAGS_WARN_MSG']