#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
EPUB目次解析ツール v2.0 - セットアップスクリプト
Calibre互換版 - 必要なライブラリを自動インストールするスクリプト
"""

import os
import sys
import subprocess
import importlib
from pathlib import Path

def install_package(package_name):
    """パッケージをインストール"""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        return True
    except subprocess.CalledProcessError:
        return False

def check_and_install_dependencies():
    """依存ライブラリのチェックとインストール"""
    print("📦 EPUB目次解析ツール v2.0 セットアップ (Calibre互換版)")
    print("=" * 60)
    
    # 必要なライブラリのリスト
    required_packages = [
        ("beautifulsoup4", "bs4", "HTML/XML解析用"),
        ("lxml", "lxml", "高速XML処理用"),
        ("six", "six", "Python2/3互換性用"),
        ("python-docx", "docx", "Word文書生成用"),
        ("defusedxml", "defusedxml", "安全なXML処理用"),
        ("tqdm", "tqdm", "プログレスバー表示用"),
        ("chardet", "chardet", "文字エンコーディング検出用")
    ]
    
    optional_packages = [
        ("jaconv", "jaconv", "日本語テキスト処理用（オプション）"),
        ("tkinterdnd2", "tkinterdnd2", "GUI版ドラッグ&ドロップ機能用（オプション）")
    ]
    
    print("🔍 必須ライブラリのチェック")
    print("-" * 40)
    
    missing_packages = []
    
    # 必須パッケージのチェック
    for package_name, import_name, description in required_packages:
        try:
            importlib.import_module(import_name)
            print(f"✅ {package_name}: インストール済み ({description})")
        except ImportError:
            print(f"❌ {package_name}: 未インストール ({description})")
            missing_packages.append(package_name)
    
    # オプションパッケージのチェック
    print("\\n🔍 オプションライブラリのチェック")
    print("-" * 40)
    
    optional_missing = []
    for package_name, import_name, description in optional_packages:
        try:
            importlib.import_module(import_name)
            print(f"✅ {package_name}: インストール済み ({description})")
        except ImportError:
            print(f"⚠️  {package_name}: 未インストール ({description})")
            optional_missing.append(package_name)
    
    # 必須パッケージのインストール
    if missing_packages:
        print(f"\\n📥 必須ライブラリをインストールします...")
        print(f"対象: {', '.join(missing_packages)}")
        
        user_input = input("\\n続行しますか？ [Y/n]: ").strip().lower()
        if user_input in ['', 'y', 'yes']:
            success_count = 0
            for package in missing_packages:
                print(f"\\n📦 {package} をインストール中...")
                if install_package(package):
                    print(f"✅ {package} インストール完了")
                    success_count += 1
                else:
                    print(f"❌ {package} インストール失敗")
            
            print(f"\\n📊 インストール結果: {success_count}/{len(missing_packages)} 成功")
        else:
            print("⏸️  インストールをキャンセルしました")
    
    # オプションパッケージのインストール
    if optional_missing:
        print(f"\\n🤔 オプションライブラリもインストールしますか？")
        print(f"対象: {', '.join(optional_missing)}")
        print("  - jaconv: 日本語テキストの高度な処理")
        print("  - tkinterdnd2: GUI版でのドラッグ&ドロップ機能")
        
        user_input = input("インストールしますか？ [y/N]: ").strip().lower()
        if user_input in ['y', 'yes']:
            for package in optional_missing:
                print(f"\\n📦 {package} をインストール中...")
                if install_package(package):
                    print(f"✅ {package} インストール完了")
                else:
                    print(f"❌ {package} インストール失敗")
    
    return len(missing_packages) == 0

def create_launcher_files():
    """ランチャーファイルを作成"""
    launchers = {
        "launch_gui.bat": '''@echo off
chcp 65001 > nul
echo 📚 EPUB目次解析ツール v2.0 GUI版起動
echo ========================================
python epub_toc_gui_v2.py
pause
''',
        "launch_cli.bat": '''@echo off
chcp 65001 > nul
echo 📚 EPUB目次解析ツール v2.0 CLI版
echo ========================================
echo.
echo 使用方法:
echo   単一ファイル: python epubsplit_word_toc_v2.py sample.epub
echo   バッチ処理:   python epubsplit_word_toc_v2.py -b folder_path
echo.
echo EPUBファイルパスを入力してください（Enterのみで終了）:
set /p epub_file="> "
if "%epub_file%"=="" (
    echo 終了します
    pause
    exit
)
python epubsplit_word_toc_v2.py "%epub_file%"
pause
''',
        "launcher.py": '''#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
EPUB目次解析ツール v2.0 - 統合ランチャー
GUI版とCLI版の選択実行
"""

import sys
import subprocess
from pathlib import Path

def main():
    print("📚 EPUB目次解析ツール v2.0 - Calibre互換版")
    print("=" * 50)
    print("1. GUI版を起動 (推奨)")
    print("2. CLI版を使用")
    print("3. セットアップ実行")
    print("4. 終了")
    print()
    
    while True:
        choice = input("選択してください [1-4]: ").strip()
        
        if choice == "1":
            print("🚀 GUI版を起動中...")
            try:
                subprocess.run([sys.executable, "epub_toc_gui_v2.py"], check=True)
            except subprocess.CalledProcessError as e:
                print(f"❌ GUI版の起動に失敗しました: {e}")
            except FileNotFoundError:
                print("❌ epub_toc_gui_v2.py が見つかりません")
            break
            
        elif choice == "2":
            print("📟 CLI版を起動中...")
            epub_file = input("EPUBファイルのパスを入力: ").strip()
            if epub_file and Path(epub_file).exists():
                try:
                    subprocess.run([sys.executable, "epubsplit_word_toc_v2.py", epub_file], check=True)
                except subprocess.CalledProcessError as e:
                    print(f"❌ 処理に失敗しました: {e}")
            else:
                print("❌ 無効なファイルパスです")
            break
            
        elif choice == "3":
            print("⚙️ セットアップを実行中...")
            try:
                subprocess.run([sys.executable, "setup.py"], check=True)
            except subprocess.CalledProcessError as e:
                print(f"❌ セットアップに失敗しました: {e}")
            break
            
        elif choice == "4":
            print("👋 終了します")
            break
            
        else:
            print("⚠️ 1-4の数字を入力してください")

if __name__ == "__main__":
    main()
'''
    }
    
    if os.name == 'nt':  # Windows
        for filename, content in launchers.items():
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f"✅ {filename} を作成しました")
            except Exception as e:
                print(f"⚠️ {filename} の作成に失敗: {e}")

def check_python_version():
    """Pythonバージョンのチェック"""
    print("🐍 Pythonバージョンチェック")
    print("-" * 40)
    
    version = sys.version_info
    print(f"Python {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 6):
        print("❌ Python 3.6以上が必要です")
        print("   最新のPythonをインストールしてください: https://python.org")
        return False
    else:
        print("✅ Pythonバージョンは要件を満たしています")
        return True

def show_usage_guide():
    """使用ガイドの表示"""
    print("\\n💡 使用方法ガイド")
    print("=" * 60)
    print("\\n【GUI版（推奨）】")
    print("  python epub_toc_gui_v2.py")
    print("  または launcher.py を実行")
    print("\\n【CLI版】")
    print("  単一ファイル:")
    print("    python epubsplit_word_toc_v2.py sample.epub")
    print("\\n  バッチ処理:")
    print("    python epubsplit_word_toc_v2.py -b /path/to/epub/folder")
    print("\\n  オプション:")
    print("    -o DIR    出力ディレクトリ指定")
    print("    -f FORMAT 出力形式 (text/word/both)")
    print("    --workers NUM 並列実行数（バッチ処理時）")
    print("\\n【新機能 v2.0】")
    print("  ✨ Calibre互換の高精度目次検出")
    print("  ✨ バッチ処理（複数ファイル一括処理）")
    print("  ✨ GUI版でのドラッグ&ドロップ対応")
    print("  ✨ プログレスバー表示")
    print("  ✨ 並列処理による高速化")

def main():
    """メイン関数"""
    print("🚀 EPUB目次解析ツール v2.0 セットアップ (Calibre互換版)")
    print("="*60)
    print("このスクリプトは必要なライブラリを自動インストールします")
    print("v2.0の新機能：バッチ処理、Calibre互換検出、高速化")
    print("")
    
    # Pythonバージョンチェック
    if not check_python_version():
        input("\\nEnterキーを押して終了...")
        return
    
    print("")
    
    # 依存ライブラリのチェックとインストール
    if check_and_install_dependencies():
        print("\\n🎉 セットアップ完了!")
        
        # ランチャーファイルの作成
        print("\\n📁 ランチャーファイルを作成中...")
        create_launcher_files()
        
        # 使用ガイドの表示
        show_usage_guide()
        
        print("\\n📚 詳細な使用方法はREADME.mdをご確認ください")
        print("🚀 launcher.py を実行してツールを起動できます")
    else:
        print("\\n😞 セットアップに失敗しました")
        print("手動でライブラリをインストールしてください:")
        print("pip install -r requirements.txt")
    
    print("")
    input("Enterキーを押して終了...")

if __name__ == "__main__":
    main()
