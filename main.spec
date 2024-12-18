# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],  # 主入口文件
    pathex=[],
    binaries=[],
    datas=[
        ('products.xlsx', '.'),  # 数据文件
        ('icon.ico', '.'),  # 图标文件
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'selenium',
        'webdriver_manager',
        'openpyxl',
        'pyperclip',
        'product_management_ui',  # 添加所有项目模块
        'stock_management_ui',
        'shopping_assistant_ui',
        'file_operations',
        'pasteboardmonitor'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 添加所有Python源文件
a.datas += [
    ('main.py', 'main.py', 'DATA'),
    ('product_management_ui.py', 'product_management_ui.py', 'DATA'),
    ('stock_management_ui.py', 'stock_management_ui.py', 'DATA'),
    ('shopping_assistant_ui.py', 'shopping_assistant_ui.py', 'DATA'),
    ('file_operations.py', 'file_operations.py', 'DATA'),
    ('pasteboardmonitor.py', 'pasteboardmonitor.py', 'DATA'),
]

# 确保添加所有必要的目录
import os
def add_dir_files(dir_path, data_list):
    if os.path.exists(dir_path):
        for root, dirs, files in os.walk(dir_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, os.path.dirname(dir_path))
                data_list.append((arcname, file_path, 'DATA'))

# 添加产品图片目录
if os.path.exists('product_images'):
    add_dir_files('product_images', a.datas)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='库存管理系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设置为True可以看到错误输出
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
    version='file_version_info.txt',
)

# 创建目录结构
import shutil
dist_dir = os.path.join('dist', '库存管理系统')
if not os.path.exists(dist_dir):
    os.makedirs(dist_dir)

# 复制必要的目录结构
dirs_to_create = ['product_images', 'temp_images', 'chrome_user_data']
for dir_name in dirs_to_create:
    dir_path = os.path.join(dist_dir, dir_name)
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

# 复制数据文件
if os.path.exists('products.xlsx'):
    shutil.copy2('products.xlsx', os.path.join(dist_dir, 'products.xlsx'))
