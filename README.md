Douyin Comment Scraper
抖音评论爬取工具，支持关键词搜索和关注页视频评论采集，具备智能去重功能。
功能特点
多模式采集
关键词模式：按地区 + 关键词搜索视频并采集评论
关注模式：采集关注页视频的评论
智能去重机制
基于 视频URL + 用户ID + 评论内容 三元组去重
自动跳过已存在于历史文档中的评论
丰富的输出内容
视频标题 / 链接 / 发布时间
用户昵称 / 抖音号 / 地区
评论内容 / 点赞数 / 发布时间
自动截图评论区
多线程处理
异步下载图片
多标签页并行操作
安装依赖
bash
pip install DrissionPage docx requests Pillow
使用说明
1. 初始化配置
python
# 配置项（代码中修改）
CHROMIUM_PORT = 6666  # 调试端口
MAX_COMMENTS = 500    # 最大采集数量
2. 运行流程
bash
python main.py
3. 交互流程
扫码登录抖音
选择采集模式：
plaintext
|请输入要进行的模式(1.关键词 2.关注)->

关键词模式需输入：
plaintext
|请输入要查询的关键词->
|请输入要查询的地区关键词->
|请选择排序方式(1.最新发布 2.综合排序)->

4. 输出文件
关键词模式：{地区关键词}{关键词}(时间戳)-关键词评论.docx
关注模式：{时间戳}-关注评论.docx
代码结构
plaintext
├── main.py          # 主程序
├── requirements.txt # 依赖清单
└── README.md        # 文档说明
注意事项
建议使用 Chrome 浏览器调试模式
首次运行需手动扫码登录
评论采集速度受网络环境影响
超过 500 条评论自动停止
历史文档需与程序在同一目录
贡献指南
Fork 本仓库
创建功能分支
提交 Pull Request
请遵循 PEP8 代码规范
许可证
MIT License
Copyright (c) 2023-present Your Name
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
