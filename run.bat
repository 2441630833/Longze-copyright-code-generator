@echo off
echo 正在启动软著源代码生成器...
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
python code_generator.py
pause 