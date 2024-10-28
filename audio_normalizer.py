import os
import sys
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QPushButton,
                            QLineEdit, QLabel, QVBoxLayout, QHBoxLayout, QWidget,
                            QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
                            QProgressDialog)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QFont, QFontDatabase
import subprocess

def init_font():
    # システムのデフォルトフォントを使用
    font_db = QFontDatabase()
    system_font = QFont(font_db.systemFont(QFontDatabase.GeneralFont))
    return system_font

class AudioNormalizer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Audio Normalizer")
        self.setGeometry(100, 100, 800, 600)

        # システムフォントを設定
        self.setFont(init_font())

        self.file_list = []  # [{'path': file_path, 'lufs': None}, ...]
        self.output_dir = ""
        self.ffmpeg_path = ""
        self.default_lufs = "-13"

        self.settings = QSettings("audio_normalizer.ini", QSettings.IniFormat)
        self.load_settings()

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # ファイル追加ボタン
        file_button_layout = QHBoxLayout()
        self.add_file_button = QPushButton("ファイルを追加")
        self.add_file_button.clicked.connect(self.select_files)
        self.clear_file_button = QPushButton("ファイル一覧をクリア")
        self.clear_file_button.clicked.connect(self.clear_files)
        file_button_layout.addWidget(self.add_file_button)
        file_button_layout.addWidget(self.clear_file_button)
        layout.addLayout(file_button_layout)

        # ファイル一覧テーブル
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(3)
        self.file_table.setHorizontalHeaderLabels(["ファイル名", "パス", "LUFS"])
        header = self.file_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        layout.addWidget(self.file_table)

        # 解析ボタン
        self.analyze_button = QPushButton("解析")
        self.analyze_button.clicked.connect(self.analyze_files)
        layout.addWidget(self.analyze_button)

        # 出力先ディレクトリ
        output_layout = QHBoxLayout()
        output_label = QLabel("出力先ディレクトリ:")
        self.output_edit = QLineEdit(self.output_dir)
        output_button = QPushButton("選択")
        output_button.clicked.connect(self.select_output_dir)
        output_layout.addWidget(output_label)
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(output_button)
        layout.addLayout(output_layout)

        # LUFS設定
        lufs_layout = QHBoxLayout()
        lufs_label = QLabel("ターゲットLUFS値:")
        self.lufs_edit = QLineEdit(self.settings.value("target_lufs", self.default_lufs))
        lufs_layout.addWidget(lufs_label)
        lufs_layout.addWidget(self.lufs_edit)
        layout.addLayout(lufs_layout)

        # 正規化ボタン
        self.normalize_button = QPushButton("正規化")
        self.normalize_button.clicked.connect(self.normalize_files)
        layout.addWidget(self.normalize_button)

        # FFmpegパス設定
        ffmpeg_layout = QHBoxLayout()
        ffmpeg_label = QLabel("ffmpegの実行ファイルパス:")
        self.ffmpeg_edit = QLineEdit(self.ffmpeg_path)
        ffmpeg_button = QPushButton("選択")
        ffmpeg_button.clicked.connect(self.select_ffmpeg_path)
        ffmpeg_layout.addWidget(ffmpeg_label)
        ffmpeg_layout.addWidget(self.ffmpeg_edit)
        ffmpeg_layout.addWidget(ffmpeg_button)
        layout.addLayout(ffmpeg_layout)

        self.setAcceptDrops(True)
        self.update_file_table()

    def closeEvent(self, event):
        self.save_settings()
        event.accept()

    def load_settings(self):
        self.output_dir = self.settings.value("output_dir", "")
        self.ffmpeg_path = self.settings.value("ffmpeg_path", self.find_ffmpeg())

    def save_settings(self):
        self.settings.setValue("output_dir", self.output_dir)
        self.settings.setValue("ffmpeg_path", self.ffmpeg_path)
        self.settings.setValue("target_lufs", self.lufs_edit.text())

    def find_ffmpeg(self):
        ffmpeg_path = ""
        for path in os.environ["PATH"].split(os.pathsep):
            ffmpeg = os.path.join(path, "ffmpeg.exe")
            if os.path.isfile(ffmpeg):
                ffmpeg_path = ffmpeg
                break
        return ffmpeg_path

    def select_output_dir(self):
        options = QFileDialog.Options()
        directory = QFileDialog.getExistingDirectory(self, "出力先ディレクトリを選択", options=options)
        if directory:
            self.output_dir = directory
            self.output_edit.setText(directory)

    def select_ffmpeg_path(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "ffmpegの実行ファイルを選択",
            "",
            "Executable Files (*.exe)",
            options=options
        )
        if file_path:
            self.ffmpeg_path = file_path
            self.ffmpeg_edit.setText(file_path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                self.add_file(file_path)
        self.update_file_table()

    def add_file(self, file_path):
        # 重複チェック
        if not any(f['path'] == file_path for f in self.file_list):
            self.file_list.append({'path': file_path, 'lufs': None})

    def clear_files(self):
        self.file_list.clear()
        self.update_file_table()

    def select_files(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "オーディオファイルを選択",
            "",
            "Audio Files (*.wav *.mp3 *.aac *.flac);;All Files (*)",
            options=options
        )
        if files:
            for file_path in files:
                self.add_file(file_path)
            self.update_file_table()

    def update_file_table(self):
        self.file_table.setRowCount(len(self.file_list))
        for row, file_info in enumerate(self.file_list):
            file_path = file_info['path']
            # ファイル名
            self.file_table.setItem(row, 0, QTableWidgetItem(os.path.basename(file_path)))
            # パス
            self.file_table.setItem(row, 1, QTableWidgetItem(file_path))
            # LUFS
            lufs_value = str(file_info['lufs']) if file_info['lufs'] is not None else ""
            self.file_table.setItem(row, 2, QTableWidgetItem(lufs_value))

    def analyze_files(self):
        if not self.file_list:
            QMessageBox.warning(self, "警告", "解析するファイルが選択されていません")
            return

        if not self.ffmpeg_path:
            QMessageBox.warning(self, "警告", "ffmpegの実行ファイルパスが指定されていません")
            return

        # プログレスダイアログを作成
        progress = QProgressDialog("オーディオファイルを解析中...", "キャンセル", 0, len(self.file_list), self)
        progress.setWindowTitle("解析中")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)

        for i, file_info in enumerate(self.file_list):
            if progress.wasCanceled():
                break

            file_path = file_info['path']
            progress.setLabelText(f"解析中: {os.path.basename(file_path)}")

            command = [
                self.ffmpeg_path,
                "-i", file_path,
                "-af", "loudnorm=I=-16:LRA=11:TP=-1.5:print_format=json",
                "-f", "null",
                "-"
            ]
            try:
                process = subprocess.Popen(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    encoding='utf-8',
                    errors='replace'
                )
                output, error = process.communicate()

                json_str = self.extract_json_from_output(error)
                if json_str:
                    import json
                    data = json.loads(json_str)
                    input_i = data.get('input_i')
                    if input_i is not None:
                        self.file_list[i]['lufs'] = float(input_i)
                        print(f"解析結果: {file_path} - LUFS: {input_i}")
                else:
                    print(f"LUFS値が見つかりません: {file_path}")
                    self.file_list[i]['lufs'] = None

            except Exception as e:
                print(f"解析エラー: {file_path}")
                print(e)
                self.file_list[i]['lufs'] = None

            progress.setValue(i + 1)

        self.update_file_table()

    def extract_json_from_output(self, output):
        """FFmpeg出力からJSON文字列を抽出する"""
        start = output.find('{')
        if start == -1:
            return None

        # JSONの終わりを探す
        count = 1
        for i in range(start + 1, len(output)):
            if output[i] == '{':
                count += 1
            elif output[i] == '}':
                count -= 1
                if count == 0:
                    return output[start:i+1]
        return None

    def normalize_files(self):
        if not self.file_list:
            QMessageBox.warning(self, "警告", "正規化するファイルが選択されていません")
            return

        if not self.output_dir:
            QMessageBox.warning(self, "警告", "出力先ディレクトリが指定されていません")
            return

        if not self.ffmpeg_path:
            QMessageBox.warning(self, "警告", "ffmpegの実行ファイルパスが指定されていません")
            return

        target_lufs = self.lufs_edit.text()
        if not target_lufs:
            QMessageBox.warning(self, "警告", "ターゲットLUFS値が指定されていません")
            return

        # プログレスダイアログを作成
        progress = QProgressDialog("オーディオファイルを正規化中...", "キャンセル", 0, len(self.file_list), self)
        progress.setWindowTitle("処理中")
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)

        processed_files = 0
        success_files = 0
        error_files = []

        for file_info in self.file_list:
            if progress.wasCanceled():
                break

            file_path = file_info['path']
            output_path = os.path.join(
                self.output_dir,
                f"{os.path.basename(file_path)}"
            )

            progress.setLabelText(f"処理中: {os.path.basename(file_path)}")

            try:
                # まず入力ファイルの情報を取得
                probe_command = [
                    self.ffmpeg_path,
                    "-i", file_path
                ]

                probe_process = subprocess.Popen(
                    probe_command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    encoding='utf-8',
                    errors='replace'
                )
                _, probe_output = probe_process.communicate()

                # ビットレートを抽出
                bitrate_match = re.search(r'bitrate:\s*(\d+)\s*kb/s', probe_output)
                # サンプルレートを抽出
                sample_rate_match = re.search(r'(\d+)\s*Hz', probe_output)
                # コーデックを抽出
                codec_match = re.search(r'Audio:\s*(\w+)', probe_output)

                bitrate = f"{bitrate_match.group(1)}k" if bitrate_match else "160k"
                sample_rate = sample_rate_match.group(1) if sample_rate_match else "44100"
                codec = codec_match.group(1) if codec_match else "mp3"

                # コーデックに応じたエンコーダーを選択
                codec_map = {
                    'mp3': 'libmp3lame',
                    'aac': 'aac',
                    'vorbis': 'libvorbis',
                    'opus': 'libopus',
                    'flac': 'flac'
                }
                encoder = codec_map.get(codec, 'copy')

                # 正規化コマンドを実行
                normalize_command = [
                    self.ffmpeg_path,
                    "-y",
                    "-i", file_path,
                    "-af", f"loudnorm=I={target_lufs}:LRA=11:TP=-1.5:linear=true",
                    "-ar", sample_rate,
                    "-c:a", encoder,
                    "-b:a", bitrate,
                    "-map_metadata", "0",  # メタデータを保持
                    "-map", "0:a:0",  # 最初のオーディオストリームのみを使用
                ]

                # フォーマット固有のオプションを追加
                if encoder == 'libmp3lame':
                    normalize_command.extend(["-q:a", "0"])  # 最高品質
                elif encoder == 'aac':
                    normalize_command.extend(["-strict", "experimental"])

                # 出力ファイルパスを追加
                normalize_command.append(output_path)

                # 処理実行
                process = subprocess.Popen(
                    normalize_command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    encoding='utf-8',
                    errors='replace'
                )
                _, error = process.communicate()

                if process.returncode == 0:
                    success_files += 1
                    print(f"正規化完了 ({success_files}/{len(self.file_list)}): {output_path}")
                    print(f"設定値 - ビットレート: {bitrate}, サンプルレート: {sample_rate}, コーデック: {encoder}")
                else:
                    error_files.append((file_path, error))
                    print(f"正規化エラー: {file_path}")
                    print(f"FFmpegエラー: {error}")

            except Exception as e:
                error_files.append((file_path, str(e)))
                print(f"正規化エラー: {file_path}")
                print(e)

            processed_files += 1
            progress.setValue(processed_files)

        # 完了メッセージを表示
        if error_files:
            error_msg = "以下のファイルで問題が発生しました:\n\n"
            for file_path, error in error_files:
                error_msg += f"- {os.path.basename(file_path)}\n"
            QMessageBox.warning(
                self,
                "完了（エラーあり）",
                f"処理が完了しました。\n成功: {success_files}個\n失敗: {len(error_files)}個\n\n{error_msg}"
            )
        else:
            QMessageBox.information(
                self,
                "完了",
                f"すべてのファイル({success_files}個)の正規化が完了しました"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(init_font())
    normalizer = AudioNormalizer()
    normalizer.show()
    sys.exit(app.exec_())
