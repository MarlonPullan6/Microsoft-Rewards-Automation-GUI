import os
import sys
import subprocess
import shutil
from pathlib import Path


def get_project_root() -> Path:
    """获取项目根目录"""
    return Path(__file__).parent.absolute()


def clean_build_dirs():
    """清理之前的构建文件"""
    root = get_project_root()
    dirs_to_clean = ['build', 'dist', '__pycache__', 'x64', 'Temp']
    
    for dir_name in dirs_to_clean:
        dir_path = root / dir_name
        if dir_path.exists():
            print(f"清理目录: {dir_path}")
            shutil.rmtree(dir_path)
    
    # 清理 .spec 文件
    spec_file = root / f"{root.name}.spec"
    if spec_file.exists():
        print(f"删除文件: {spec_file}")
        spec_file.unlink()


def build_executable():
    """使用 PyInstaller 构建可执行文件"""
    root = get_project_root()
    
    # PyInstaller 命令参数
    pyinstaller_cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--name=MicrosoftRewardsAutomation",  # 输出名称
        "--windowed",  # 使用 GUI 模式，不显示命令窗口
        "--icon=" + str(root / "Assets" / "app.ico"),  # 应用图标
        "--add-data=" + str(root / "Assets") + os.pathsep + "Assets",  # 包含 Assets 目录
        "--collect-all=PyQt6",  # 收集所有 PyQt6 相关文件
        "--collect-all=playwright",  # 收集 Playwright 相关文件
        "--hidden-import=PyQt6",  # 隐式导入
        "--hidden-import=playwright",
        "--hidden-import=playwright.async_api",
        "--distpath=" + str(root / "x64"),  # 输出目录
        "--workpath=" + str(root / "Temp"),  # 临时构建目录
        "--specpath=" + str(root),  # spec 文件位置
        str(root / "gui.py"),  # 主入口文件（GUI）
    ]
    
    print("=" * 60)
    print("开始构建 Microsoft Rewards Automation...")
    print("=" * 60)
    print("\nPyInstaller 命令:")
    print(" ".join(pyinstaller_cmd))
    print("\n" + "=" * 60 + "\n")
    
    # 执行 PyInstaller
    try:
        result = subprocess.run(pyinstaller_cmd, check=True)
        return result.returncode == 0
    except subprocess.CalledProcessError as e:
        print(f"错误: PyInstaller 构建失败 (返回码: {e.returncode})")
        return False
    except FileNotFoundError:
        print("错误: PyInstaller 未安装，请运行: pip install pyinstaller")
        return False


def clean_temp_dir():
    """清理临时构建目录"""
    root = get_project_root()
    temp_dir = root / "Temp"
    
    if temp_dir.exists():
        print("清理临时文件...")
        try:
            shutil.rmtree(temp_dir)
            print(f"✓ 已清理: {temp_dir}\n")
        except Exception as e:
            print(f"警告: 无法清理 {temp_dir}: {e}\n")


def verify_build():
    """验证构建结果"""
    root = get_project_root()
    app_dir = root / "x64" / "MicrosoftRewardsAutomation"
    exe_path = app_dir / "MicrosoftRewardsAutomation.exe"
    
    if app_dir.exists() and exe_path.exists():
        print("\n" + "=" * 60)
        print("✓ 构建成功!")
        print("=" * 60)
        print(f"应用目录: {app_dir}")
        print(f"可执行文件: {exe_path}")
        print("=" * 60 + "\n")
        return True
    else:
        print("\n" + "=" * 60)
        print("✗ 构建失败: 未找到输出文件")
        print(f"预期目录: {app_dir}")
        print(f"预期可执行文件: {exe_path}")
        print("=" * 60 + "\n")
        return False


def main():
    """主构建流程"""
    root = get_project_root()
    
    print(f"项目根目录: {root}\n")
    
    # 清理旧构建
    clean_build_dirs()
    
    # 构建可执行文件
    if not build_executable():
        return 1
    
    # 验证构建结果
    if not verify_build():
        return 1
    
    # 清理临时文件
    clean_temp_dir()
    
    print("构建流程完成！")
    return 0


if __name__ == "__main__":
    sys.exit(main())
