"""
使用 Python 调试器 (pdb) 调试 docx2markdown
使用方法：
    python debug_with_pdb.py <docx文件路径> <输出md文件路径>
    
调试命令：
    n (next): 执行下一行
    s (step): 进入函数内部
    c (continue): 继续执行直到下一个断点
    l (list): 显示当前代码
    p <变量名>: 打印变量值
    pp <变量名>: 美化打印变量值
    b <行号>: 设置断点
    q (quit): 退出调试器
"""

import sys
import pdb
from pathlib import Path

# 导入要调试的模块
from src.docx2markdown._docx_to_markdown import docx_to_markdown


def main():
    """使用 pdb 调试主函数"""
    # 测试文件路径
    if len(sys.argv) >= 3:
        docx_file = sys.argv[1]
        output_md = sys.argv[2]
    else:
        docx_file = "demo/test-text.docx"
        output_md = "debug_output.md"
        print(f"使用默认测试文件: {docx_file}")
    
    if not Path(docx_file).exists():
        print(f"错误: 文件不存在: {docx_file}")
        return
    
    print("=" * 60)
    print("开始使用 pdb 调试")
    print("=" * 60)
    print(f"输入文件: {docx_file}")
    print(f"输出文件: {output_md}")
    print("\n提示: 程序将在 docx_to_markdown 函数入口处暂停")
    print("使用 'c' 继续执行，或使用 's' 进入函数内部")
    print("=" * 60)
    
    # 设置断点并开始调试
    pdb.set_trace()
    
    # 调用要调试的函数
    docx_to_markdown(docx_file, output_md)
    
    print("转换完成！")


if __name__ == "__main__":
    main()
