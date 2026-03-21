# 这是一个高性能的 AutoCAD 插件脚本：利用 COM 接口深度捕获当前图纸及外部参照 (XRef) 中所有含 "FM" 标识的防火门规格，并一键去重导出至 Excel。

# 1. 导入 AutoCAD 操控库和正则模块，准备潜入图纸底层
import win32com.client
import re
import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import messagebox

# 专门给妈准备的提示弹窗
def show_popup(title, message, is_error=False):
    """显示友好的提示弹窗"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口，只留弹窗
    if is_error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()

# 2. 编写一个无敌的正则，抓取所有含 FM 且带长后缀的字符串（支持中文和括号）
# 匹配逻辑：识别以 "FM" 开头的完整字符串，包含后续的所有数字、字母、连字符及括号参数（例如 FM1021-甲）
FM_PATTERN = r'FM[\d\w\-\u4e00-\u9fff()（）]*'

# 3. 建立与 AutoCAD 的 COM 通信，如果没打开就报错提醒
try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
except:
    print("错误：请先打开 AutoCAD 应用！")
    exit(1)

doc = acad.ActiveDocument
mspace = doc.ModelSpace

# 初始化 FM 规格集合（用于自动去重）
fm_specs = set()
xref_paths = set()

# 定义扫描函数，支持在给定的空间中提取 FM 规格
def scan_space(space, space_name=""):
    """扫描指定的空间（ModelSpace 或 PaperSpace）中的对象"""
    space_prefix = f"[{space_name}]" if space_name else ""
    
    for obj in space:
        try:
            obj_text = ""
            
            # 处理 BlockReference 对象
            if obj.ObjectName == "AcDbBlockReference":
                # 检查是否为外部参照
                if hasattr(obj, 'IsXRef') and obj.IsXRef:
                    xref_path = obj.Name
                    xref_paths.add(xref_path)
                    print(f"发现外部参照 {space_prefix}: {xref_path}")
                    
                    # 扫描 XRef 中的属性
                    if hasattr(obj, 'Attributes'):
                        for attr in obj.Attributes:
                            obj_text += attr.TextString + " "
            
            # 处理 Text 对象
            elif obj.ObjectName == "AcDbText":
                obj_text = obj.TextString
            
            # 处理 MText 对象
            elif obj.ObjectName == "AcDbMText":
                obj_text = obj.TextString
            
            # 用正则提取 FM 规格
            if obj_text:
                matches = re.findall(FM_PATTERN, obj_text)
                for match in matches:
                    fm_specs.add(match)
        
        except Exception as e:
            continue

# 4. 遍历当前图纸里所有的 BlockReference 和 Text 对象
# 5. 扫描过程中，如果发现 IsXRef 为 True，立即打印出那个外部引用的路径
# 6. 把所有抓到的 FM 规格塞进一个列表，顺手做一个去重处理

# 先扫描模型空间（ModelSpace）
print("正在扫描模型空间（ModelSpace）...")
scan_space(mspace, "ModelSpace")

# 再扫描所有布局中的纸张空间（PaperSpace/Layout）
print("正在扫描布局空间（Layouts）...")
layouts = doc.Layouts
for layout in layouts:
    # 跳过模型布局（Model）
    if layout.Name.lower() != "model":
        try:
            pspace = layout.Block
            print(f"  -> 扫描布局: {layout.Name}")
            scan_space(pspace, f"Layout:{layout.Name}")
        except Exception as e:
            print(f"  -> 扫描布局 {layout.Name} 时出错: {e}")
            continue

# 转换为有序列表
fm_list = sorted(list(fm_specs))

# 7. 最后用 pandas 把这个列表吐成一个整齐的 Excel，文件名要带上今天的日期
today = datetime.now().strftime("%Y%m%d")
output_dir = os.path.dirname(__file__)
excel_file = os.path.join(output_dir, f"FM规格_去重导出_{today}.xlsx")

# 创建 DataFrame
df = pd.DataFrame({
    "FM规格": fm_list,
    "统计": [1] * len(fm_list)
})

# 导出 Excel
df.to_excel(excel_file, index=False, sheet_name="防火门规格")

# 根据结果显示相应的提示弹窗
if not fm_list:
    show_popup("⚠️ 警告", "在图纸里【一扇防火门】都没找到！\n请检查：\n1. 图层是否被冻结？\n2. 标注是不是被炸开了？", True)
elif len(xref_paths) > 0:
    show_popup("💡 提示", f"已成功导出 {len(fm_list)} 类规格。\n注意：图纸包含 {len(xref_paths)} 个外部引用，\n请确认附件是否完整，否则数量会少！")
else:
    show_popup("✅ 成功", f"防火门数好啦！\n一共抓到 {len(fm_list)} 种规格，Excel 已生成。")

# 输出到控制台用于调试
print(f"\n✓ 成功导出 {len(fm_list)} 个不同的 FM 规格至: {excel_file}")
print(f"✓ 识别到 {len(xref_paths)} 个外部参照")
print("\n所有 FM 规格:")
for spec in fm_list:
    print(f"  - {spec}")
