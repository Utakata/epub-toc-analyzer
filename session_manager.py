#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
セッション管理ツール - 開発プロジェクトの継続性を保つ
Claude Code Actionでの開発継続用
"""

import json
import os
import time
from datetime import datetime
from pathlib import Path
import uuid

class ProjectSession:
    """プロジェクトセッション管理クラス"""
    
    def __init__(self, project_name="epub-toc-analyzer"):
        self.project_name = project_name
        self.session_id = str(uuid.uuid4())[:8]
        self.session_file = f".session_{self.session_id}.json"
        self.created_at = datetime.now().isoformat()
        self.current_dir = os.getcwd()
        
        # セッション情報
        self.session_data = {
            "session_id": self.session_id,
            "project_name": self.project_name,
            "created_at": self.created_at,
            "current_dir": self.current_dir,
            "git_branch": self.get_git_branch(),
            "last_commit": self.get_last_commit(),
            "files_status": self.get_files_status(),
            "development_stage": "initial_setup",
            "next_tasks": [],
            "completed_tasks": [],
            "notes": []
        }
    
    def get_git_branch(self):
        """現在のGitブランチを取得"""
        try:
            import subprocess
            result = subprocess.run(['git', 'branch', '--show-current'], 
                                  capture_output=True, text=True)
            return result.stdout.strip() if result.returncode == 0 else "unknown"
        except:
            return "unknown"
    
    def get_last_commit(self):
        """最新のコミット情報を取得"""
        try:
            import subprocess
            result = subprocess.run(['git', 'log', '-1', '--oneline'], 
                                  capture_output=True, text=True)
            return result.stdout.strip() if result.returncode == 0 else "unknown"
        except:
            return "unknown"
    
    def get_files_status(self):
        """プロジェクトファイルの状況を取得"""
        files = {}
        project_files = [
            "README.md",
            "requirements.txt", 
            "setup.py",
            "epubsplit_word_toc_v2.py",
            "epub_toc_gui_v2.py",
            ".gitignore"
        ]
        
        for file_name in project_files:
            if os.path.exists(file_name):
                stat = os.stat(file_name)
                files[file_name] = {
                    "exists": True,
                    "size": stat.st_size,
                    "modified": datetime.fromtimestamp(stat.st_mtime).isoformat()
                }
            else:
                files[file_name] = {"exists": False}
        
        return files
    
    def add_task(self, task, priority="medium"):
        """タスクを追加"""
        task_item = {
            "task": task,
            "priority": priority,
            "added_at": datetime.now().isoformat(),
            "status": "pending"
        }
        self.session_data["next_tasks"].append(task_item)
    
    def complete_task(self, task_description):
        """タスクを完了としてマーク"""
        completed_task = {
            "task": task_description,
            "completed_at": datetime.now().isoformat()
        }
        self.session_data["completed_tasks"].append(completed_task)
    
    def add_note(self, note):
        """メモを追加"""
        note_item = {
            "note": note,
            "timestamp": datetime.now().isoformat()
        }
        self.session_data["notes"].append(note_item)
    
    def set_stage(self, stage):
        """開発段階を設定"""
        self.session_data["development_stage"] = stage
        self.session_data["stage_updated_at"] = datetime.now().isoformat()
    
    def save_session(self):
        """セッション情報を保存"""
        with open(self.session_file, 'w', encoding='utf-8') as f:
            json.dump(self.session_data, f, indent=2, ensure_ascii=False)
        return self.session_file
    
    def generate_resume_info(self):
        """再開用情報を生成"""
        resume_info = f"""
# 📋 プロジェクト継続情報

## 🆔 セッション情報
- **Session ID**: `{self.session_id}`
- **プロジェクト**: {self.project_name}
- **作成日時**: {self.created_at}
- **作業ディレクトリ**: `{self.current_dir}`

## 📂 プロジェクト状況
- **Gitブランチ**: {self.session_data['git_branch']}
- **最新コミット**: {self.session_data['last_commit']}
- **開発段階**: {self.session_data['development_stage']}

## ✅ 完了済みタスク
"""
        for task in self.session_data['completed_tasks']:
            resume_info += f"- ✅ {task['task']} ({task['completed_at']})\n"
        
        resume_info += "\n## 📋 次のタスク\n"
        for task in self.session_data['next_tasks']:
            priority_emoji = {"high": "🔴", "medium": "🟡", "low": "🟢"}.get(task['priority'], "⚪")
            resume_info += f"- {priority_emoji} {task['task']} (優先度: {task['priority']})\n"
        
        resume_info += "\n## 📝 メモ\n"
        for note in self.session_data['notes']:
            resume_info += f"- 📝 {note['note']} ({note['timestamp']})\n"
        
        resume_info += f"""
## 🔄 再開方法

### Claude Code Action再開コマンド:
```bash
cd {self.current_dir}
# セッション情報確認
cat {self.session_file}
```

### プロジェクト再開時の確認項目:
1. **作業ディレクトリ**: `{self.current_dir}`
2. **プロジェクトファイル**: epubsplit_word_toc_v2.py, epub_toc_gui_v2.py
3. **Gitリポジトリ**: https://github.com/Utakata/epub-toc-analyzer
4. **現在のブランチ**: {self.session_data['git_branch']}

### 次回セッション開始時の指示:
```
プロジェクト継続: Session ID `{self.session_id}`
作業ディレクトリ: {self.current_dir}
現在の開発段階: {self.session_data['development_stage']}
```
"""
        
        return resume_info

def create_current_session():
    """現在のセッション情報を作成"""
    session = ProjectSession("epub-toc-analyzer")
    
    # 完了済みタスクを記録
    completed_tasks = [
        "GitHubリポジトリ作成 (epub-toc-analyzer)",
        "Calibre互換TOC検出クラス実装",
        "バッチ処理機能実装",
        "GUI版v2.0作成 (ドラッグ&ドロップ対応)",
        "並列処理による高速化実装",
        "エラー処理とエンコーディング検出強化",
        "セットアップスクリプト作成",
        "README.md作成 (v2.0対応)",
        "requirements.txt更新",
        ".gitignore設定",
        "初期コミット作成とGitHubプッシュ"
    ]
    
    for task in completed_tasks:
        session.complete_task(task)
    
    # 次のタスクを設定
    next_tasks = [
        ("オリジナルEpubSplitファイルのマイグレーション", "high"),
        ("テスト用EPUBサンプルファイル作成", "medium"), 
        ("バッチ処理のパフォーマンステスト", "medium"),
        ("GUI版の詳細テスト", "medium"),
        ("ドキュメント充実 (使用例追加)", "low"),
        ("CI/CD環境構築 (GitHub Actions)", "low"),
        ("パッケージ化 (PyPI対応)", "low")
    ]
    
    for task, priority in next_tasks:
        session.add_task(task, priority)
    
    # 重要なメモを追加
    important_notes = [
        "CalibreのDeepWikiから学んだXPath式ベース検出を実装済み",
        "ユーザーの既存NVCファイル処理実績あり",
        "v2.0ではtqdm, chardet等の新依存関係追加",
        "GUI版でtkinterdnd2によるドラッグ&ドロップ実装",
        "バッチ処理でThreadPoolExecutor使用",
        "メモリサーバーで情報記憶済み (default_user, EpubSplit関係)"
    ]
    
    for note in important_notes:
        session.add_note(note)
    
    # 現在の開発段階を設定
    session.set_stage("core_implementation_complete")
    
    return session

def main():
    """メイン実行"""
    print("🔄 セッション情報を生成中...")
    
    # セッション作成
    session = create_current_session()
    
    # セッションファイル保存
    session_file = session.save_session()
    print(f"💾 セッション保存: {session_file}")
    
    # 再開情報を生成
    resume_info = session.generate_resume_info()
    
    # 再開情報をファイルに保存
    resume_file = f"PROJECT_RESUME_{session.session_id}.md"
    with open(resume_file, 'w', encoding='utf-8') as f:
        f.write(resume_info)
    
    print(f"📋 再開情報保存: {resume_file}")
    print(f"🆔 **Session ID: {session.session_id}**")
    print("\n" + "="*60)
    print(resume_info)
    print("="*60)
    
    return session.session_id

if __name__ == "__main__":
    session_id = main()
