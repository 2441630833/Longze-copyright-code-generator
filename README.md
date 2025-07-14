# 软著源代码生成器

这是一个用于生成软件著作权申请文档的工具，可以自动从项目中提取源代码并生成符合要求的Word文档。

## 功能特点

- 图形用户界面，简单易用
- 支持指定多种文件后缀
- 自动提取源代码并格式化
- 自动添加页眉（软件名称和版本号）
- 自动添加页脚（作者信息）
- 生成符合软著申请要求的文档（每页50行）
- 进度条显示处理进度

## 安装使用

### 方法一：直接运行可执行文件

1. 在dist目录下下载 `code_generator.exe`
2. 双击运行即可使用

### 方法二：通过Python运行

1. 确保已安装Python 3.6或更高版本
2. 安装依赖库：
   ```
   pip install -r requirements.txt
   ```
3. 运行程序：
   ```
   python code_generator.py
   ```
   或者运行批处理文件：
   ```
   run.bat
   ```

## 使用说明

1. 填写软件名称、版本号和作者名
2. 选择项目路径（包含源代码的文件夹）
3. 输入需要提取的文件后缀，多个后缀用英文逗号分隔（如：java,xml,yml）
4. 点击"生成"按钮开始处理
5. 等待处理完成，生成的Word文档将保存在程序运行目录下

## 注意事项

- 生成的文档名称格式为：[软件名称]源代码(前后30页).docx
- 文档最多包含60页，每页约50行代码
- 支持UTF-8和GBK编码的源代码文件 

## 许可证

本项目采用 Apache License 2.0 许可证进行授权。详情请参阅 [LICENSE](LICENSE) 文件。

```
Copyright 2024 Your Name or Organization

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
``` 