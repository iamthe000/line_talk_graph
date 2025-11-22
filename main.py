import sys
import os
import glob
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QLabel, QMessageBox, QLineEdit)  # QLineEdit追加
from PyQt6.QtCore import Qt
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# 日本語フォントの設定（Windows/Mac/Linuxで一般的なフォントを探す）
# グラフの文字化けを防ぐため
def configure_japanese_font():
    import matplotlib.font_manager as fm
    fonts = [f.name for f in fm.fontManager.ttflist]
    target_fonts = ['Meiryo', 'Yu Gothic', 'MS Gothic', 'Hiragino Sans', 'TakaoGothic', 'IPAGothic']
    for font in target_fonts:
        if font in fonts:
            plt.rcParams['font.family'] = font
            return
    # 見つからない場合はデフォルトのまま（豆腐になる可能性あり）

class MplCanvas(FigureCanvas):
    """MatplotlibのグラフをPyQtに埋め込むためのキャンバス"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        configure_japanese_font()
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)

class LineTalkAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("LINEトーク履歴 ランキング＆グラフ化ツール")
        self.resize(1000, 600)
        self.target_dir = "talk_upload"

        # メインウィジェットとレイアウト
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QHBoxLayout(main_widget) # 左右に分割

        # 左側：操作パネルとランキング表
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        # 説明ラベル
        self.label_info = QLabel(f"フォルダ '{self.target_dir}' 内のファイルを解析します")
        left_layout.addWidget(self.label_info)

        # 解析ボタン
        self.btn_analyze = QPushButton("読み込み / 更新")
        self.btn_analyze.setStyleSheet("padding: 10px; font-weight: bold; font-size: 14px;")
        self.btn_analyze.clicked.connect(self.load_and_analyze)
        left_layout.addWidget(self.btn_analyze)

        # ランキング表
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["順位", "ニックネーム", "行数(トーク量)"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch) # 名前列を伸縮
        left_layout.addWidget(self.table)

        # 除外ニックネーム入力欄
        self.exclude_label = QLabel("グラフから除外したいニックネーム（カンマ区切り）:")
        left_layout.addWidget(self.exclude_label)
        self.exclude_input = QLineEdit()
        self.exclude_input.setPlaceholderText("例: 山田,田中")
        left_layout.addWidget(self.exclude_input)

        layout.addWidget(left_panel, 1) # 左側 (比率1)

        # 右側：グラフ
        self.canvas = MplCanvas(self, width=5, height=4, dpi=100)
        layout.addWidget(self.canvas, 2) # 右側 (比率2)

        # 起動時にフォルダチェック
        self.check_dir()

    def check_dir(self):
        if not os.path.exists(self.target_dir):
            os.makedirs(self.target_dir)
            QMessageBox.information(self, "フォルダ作成", 
                                    f"'{self.target_dir}' フォルダを作成しました。\n"
                                    "LINEのトーク履歴txtファイルをそこに入れて、\n"
                                    "「読み込み / 更新」ボタンを押してください。")

    def load_and_analyze(self):
        files = glob.glob(os.path.join(self.target_dir, "*.txt"))
        
        if not files:
            QMessageBox.warning(self, "ファイルなし", f"'{self.target_dir}' にtxtファイルが見つかりません。")
            return

        data = []

        for file_path in files:
            # ファイル名から拡張子を除去してニックネームにする
            filename = os.path.basename(file_path)
            nickname = os.path.splitext(filename)[0]
            
            line_count = 0
            try:
                # エンコーディング対応: UTF-8を試し、ダメならShift-JIS等を試す簡易ロジック
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        line_count = len(f.readlines())
                except UnicodeDecodeError:
                    with open(file_path, 'r', encoding='cp932') as f: # Windows Shift-JIS
                        line_count = len(f.readlines())
            except Exception as e:
                print(f"Error reading {filename}: {e}")
                continue

            data.append({"name": nickname, "count": line_count})

        # 行数が多い順にソート
        data.sort(key=lambda x: x["count"], reverse=True)

        # データの更新
        self.update_table(data)
        self.update_graph(data)

    def update_table(self, data):
        self.table.setRowCount(len(data))
        for i, item in enumerate(data):
            # 順位
            rank_item = QTableWidgetItem(str(i + 1))
            rank_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 0, rank_item)
            
            # ニックネーム
            self.table.setItem(i, 1, QTableWidgetItem(item["name"]))
            
            # 行数
            count_item = QTableWidgetItem(f"{item['count']:,}") # カンマ区切り
            count_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(i, 2, count_item)

    def update_graph(self, data):
        self.canvas.axes.cla() # 前のグラフをクリア

        # 除外ニックネーム取得
        exclude_text = self.exclude_input.text().strip()
        exclude_names = [name.strip() for name in exclude_text.split(",") if name.strip()] if exclude_text else []

        # 除外処理
        filtered_data = [d for d in data if d["name"] not in exclude_names]

        names = [d["name"] for d in filtered_data]
        counts = [d["count"] for d in filtered_data]

        # 棒グラフの描画（見やすくするため横向き棒グラフにします）
        # 上位が上に来るようにリストを逆順にする
        y_pos = range(len(names))
        self.canvas.axes.barh(y_pos, counts[::-1], color='skyblue')
        self.canvas.axes.set_yticks(y_pos)
        self.canvas.axes.set_yticklabels(names[::-1])
        
        self.canvas.axes.set_xlabel("トーク履歴の総行数")
        self.canvas.axes.set_title("トーク量 ランキング比較")
        
        self.canvas.axes.grid(axis='x', linestyle='--', alpha=0.7)
        self.canvas.fig.tight_layout()
        self.canvas.draw()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LineTalkAnalyzer()
    window.show()
    sys.exit(app.exec())
