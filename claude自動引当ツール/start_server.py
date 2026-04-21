"""在庫引当ツール起動スクリプト"""
import os
import sys

# このスクリプトのディレクトリに移動
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.getcwd())

from app import app

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5001, debug=False)
