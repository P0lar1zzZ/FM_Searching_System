"""
防火门图标生成脚本
生成：外围白色背景，中间黑色圆形的图标
"""
from PIL import Image, ImageDraw
import os

def create_icon(output_path="fm_icon.ico", size=256):
    """
    生成防火门图标
    
    Args:
        output_path: 输出文件路径
        size: 图标大小（像素）
    """
    # 创建新图像 - 白色背景
    img = Image.new('RGB', (size, size), color='white')
    draw = ImageDraw.Draw(img)
    
    # 绘制外圆边框（深灰色）
    border_width = int(size * 0.05)
    outer_bbox = [border_width, border_width, size - border_width, size - border_width]
    draw.ellipse(outer_bbox, outline='#333333', width=int(size * 0.02))
    
    # 绘制中间黑色圆形
    circle_radius = int(size * 0.25)
    center_x, center_y = size // 2, size // 2
    circle_bbox = [
        center_x - circle_radius,
        center_y - circle_radius,
        center_x + circle_radius,
        center_y + circle_radius
    ]
    draw.ellipse(circle_bbox, fill='#000000', outline='#000000')
    
    # 绘制中心白色点
    dot_radius = int(size * 0.08)
    dot_bbox = [
        center_x - dot_radius,
        center_y - dot_radius,
        center_x + dot_radius,
        center_y + dot_radius
    ]
    draw.ellipse(dot_bbox, fill='white', outline='white')
    
    # 转换为 RGBA 模式并保存为 ICO
    img = img.convert('RGBA')
    img.save(output_path, 'ICO')
    
    print(f"✓ 图标已生成: {output_path}")
    print(f"✓ 图标尺寸: {size}x{size} 像素")

if __name__ == "__main__":
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, "fm_icon.ico")
    
    try:
        create_icon(icon_path)
        print("\n✅ 图标创建完成！")
    except ImportError:
        print("❌ 错误：需要安装 Pillow 库")
        print("请运行: pip install Pillow")
    except Exception as e:
        print(f"❌ 生成图标时出错: {e}")
