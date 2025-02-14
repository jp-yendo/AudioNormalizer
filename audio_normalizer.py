import os
import sys
import re
import subprocess
import json
from PyQt5.QtWidgets import (QApplication, QMainWindow, QFileDialog, QPushButton,
                            QLineEdit, QLabel, QVBoxLayout, QHBoxLayout, QWidget,
                            QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView,
                            QProgressDialog)
from PyQt5.QtCore import Qt, QSettings, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QFontDatabase

def init_font():
    # システムのデフォルトフォントを使用
    font_db = QFontDatabase()
    system_font = QFont(font_db.systemFont(QFontDatabase.GeneralFont))
    return system_font

class AnalyzeWorker(QThread):
    progress = pyqtSignal(int, str)  # 進捗と現在のファイル名
    finished = pyqtSignal(list)  # 処理結果
    error = pyqtSignal(str)  # エラーメッセージ

    def __init__(self, file_list, ffmpeg_path):
        super().__init__()
        self.file_list = file_list  # 元のリストを参照として保持
        self.ffmpeg_path = ffmpeg_path
        self.is_cancelled = False

    def run(self):
        results = []
        for i, file_info in enumerate(self.file_list):
            if self.is_cancelled:
                break

            file_path = file_info['path']
            self.progress.emit(i, os.path.basename(file_path))

            try:
                command = [
                    self.ffmpeg_path,
                    "-i", file_path,
                    "-af", "loudnorm=I=-16:LRA=11:TP=-1.5:print_format=json",
                    "-f", "null",
                    "-"
                ]
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
                    data = json.loads(json_str)
                    input_i = data.get('input_i')
                    if input_i is not None:
                        # 直接file_infoを更新
                        file_info['lufs'] = float(input_i)
                        results.append(file_info)
                        self.progress.emit(i + 1, "")
                        continue

                # JSONの解析に失敗した場合のみNoneを設定
                file_info['lufs'] = None
                results.append(file_info)
                self.progress.emit(i + 1, "")

            except Exception as e:
                file_info['lufs'] = None
                results.append(file_info)
                self.error.emit(f"解析エラー: {file_path}\n{str(e)}")
                self.progress.emit(i + 1, "")

        if not self.is_cancelled:
            self.finished.emit(results)

    def extract_json_from_output(self, output):
        start = output.find('{')
        if start == -1:
            return None

        count = 1
        for i in range(start + 1, len(output)):
            if output[i] == '{':
                count += 1
            elif output[i] == '}':
                count -= 1
                if count == 0:
                    return output[start:i+1]
        return None


class NormalizeWorker(QThread):
    progress = pyqtSignal(int, str)  # 進捗と現在のファイル名
    finished = pyqtSignal(int, list)  # 成功数とエラーリスト
    error = pyqtSignal(str)  # エラーメッセージ

    def __init__(self, file_list, ffmpeg_path, output_dir, target_lufs):
        super().__init__()
        self.file_list = file_list
        self.ffmpeg_path = ffmpeg_path
        self.output_dir = output_dir
        self.target_lufs = target_lufs
        self.is_cancelled = False

    def run(self):
        success_files = 0
        error_files = []

        for i, file_info in enumerate(self.file_list):
            if self.is_cancelled:
                break

            file_path = file_info['path']
            output_path = os.path.join(
                self.output_dir,
                f"{os.path.basename(file_path)}"
            )

            self.progress.emit(i, os.path.basename(file_path))

            try:
                # 入力ファイルの情報を取得
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

                # 各種パラメータを抽出
                bitrate_match = re.search(r'bitrate:\s*(\d+)\s*kb/s', probe_output)
                sample_rate_match = re.search(r'(\d+)\s*Hz', probe_output)
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
                    "-af", f"loudnorm=I={self.target_lufs}:LRA=11:TP=-1.5:linear=true",
                    "-ar", sample_rate,
                    "-c:a", encoder,
                    "-b:a", bitrate,
                    "-map_metadata", "0",
                    "-map", "0:a:0",
                ]

                if encoder == 'libmp3lame':
                    normalize_command.extend(["-q:a", "0"])
                elif encoder == 'aac':
                    normalize_command.extend(["-strict", "experimental"])

                normalize_command.append(output_path)

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
                    self.progress.emit(i + 1, "")
                else:
                    error_files.append((file_path, error))
                    self.error.emit(f"正規化エラー: {file_path}\n{error}")
                    self.progress.emit(i + 1, "")

            except Exception as e:
                error_files.append((file_path, str(e)))
                self.error.emit(f"正規化エラー: {file_path}\n{str(e)}")
                self.progress.emit(i + 1, "")

        self.finished.emit(success_files, error_files)

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
        self.file_table.setHorizontalHeaderLabels(["ファイル名", "ディレクトリ", "LUFS"])
        header = self.file_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        # テーブルを編集不可に設定
        self.file_table.setEditTriggers(QTableWidget.NoEditTriggers)
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
            "ffmpeg.exe",
            "Executable Files (ffmpeg.exe);;All Files (*)",
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
        try:
            self.file_table.setRowCount(len(self.file_list))
            for row, file_info in enumerate(self.file_list):
                file_path = file_info['path']
                # ファイル名
                name_item = QTableWidgetItem(os.path.basename(file_path))
                name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
                self.file_table.setItem(row, 0, name_item)

                # ディレクトリパス
                dir_item = QTableWidgetItem(os.path.dirname(file_path))
                dir_item.setFlags(dir_item.flags() & ~Qt.ItemIsEditable)
                self.file_table.setItem(row, 1, dir_item)

                # LUFS
                lufs = file_info.get('lufs')
                if lufs is not None:
                    lufs_text = f"{lufs:.1f}"
                    lufs_item = QTableWidgetItem(lufs_text)
                    lufs_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    lufs_item = QTableWidgetItem("")
                lufs_item.setFlags(lufs_item.flags() & ~Qt.ItemIsEditable)
                self.file_table.setItem(row, 2, lufs_item)

            # テーブルの更新を強制
            self.file_table.viewport().update()
            self.file_table.repaint()

        except Exception as e:
            print(f"テーブル更新エラー: {str(e)}")

    def analyze_files(self):
        if not self.file_list:
            QMessageBox.warning(self, "警告", "解析するファイルが選択されていません")
            return

        if not self.ffmpeg_path:
            QMessageBox.warning(self, "警告", "ffmpegの実行ファイルパスが指定されていません")
            return

        # メインウィンドウを無効化
        self.setEnabled(False)

        try:
            # プログレスダイアログを作成
            self.progress_dialog = QProgressDialog("オーディオファイルを解析中...", "キャンセル", 0, len(self.file_list), self)
            self.progress_dialog.setWindowTitle("解析中")
            self.progress_dialog.setWindowModality(Qt.ApplicationModal)
            self.progress_dialog.setMinimumDuration(0)

            # ワーカーを作成
            self.analyze_worker = AnalyzeWorker(self.file_list, self.ffmpeg_path)

            # シグナル接続
            self.progress_dialog.canceled.connect(self.cancel_analyze)
            self.analyze_worker.progress.connect(self.update_analyze_progress)
            self.analyze_worker.error.connect(lambda msg: QMessageBox.warning(self, "解析エラー", msg))
            self.analyze_worker.finished.connect(self.handle_analyze_finished)

            # ワーカー開始
            self.analyze_worker.start()

        except Exception as e:
            self.setEnabled(True)
            QMessageBox.critical(self, "エラー", f"解析処理の初期化中にエラーが発生しました:\n{str(e)}")

    def update_analyze_progress(self, value, filename):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.setValue(value)
            if filename:
                self.progress_dialog.setLabelText(f"解析中: {filename}")

    def cleanup_progress_dialog(self):
        """プログレスダイアログを安全に削除"""
        try:
            if hasattr(self, 'progress_dialog'):
                self.progress_dialog.close()
                delattr(self, 'progress_dialog')
        except:
            pass

    def handle_analyze_finished(self, results):
        try:
            # プログレスダイアログを閉じる
            self.cleanup_progress_dialog()

            # 結果をメインのファイルリストに反映
            path_to_lufs = {result['path']: result['lufs'] for result in results}
            for file_info in self.file_list:
                if file_info['path'] in path_to_lufs:
                    file_info['lufs'] = path_to_lufs[file_info['path']]

            # テーブルを更新
            self.update_file_table()

        except Exception as e:
            QMessageBox.critical(self, "エラー", f"解析結果の処理中にエラーが発生しました:\n{str(e)}")

        finally:
            self.setEnabled(True)

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

        # メインウィンドウを無効化
        self.setEnabled(False)

        # プログレスダイアログを作成
        self.progress_dialog = QProgressDialog("オーディオファイルを正規化中...", "キャンセル", 0, len(self.file_list), self)
        self.progress_dialog.setWindowTitle("処理中")
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.setMinimumDuration(0)

        # ワーカーを作成
        self.normalize_worker = NormalizeWorker(self.file_list, self.ffmpeg_path, self.output_dir, target_lufs)
        self.progress_dialog.canceled.connect(self.cancel_normalize)
        self.normalize_worker.progress.connect(self.update_normalize_progress)
        self.normalize_worker.finished.connect(self.handle_normalize_finished)
        self.normalize_worker.error.connect(lambda msg: QMessageBox.warning(self, "正規化エラー", msg))
        self.normalize_worker.start()

    def update_normalize_progress(self, value, filename):
        if hasattr(self, 'progress_dialog'):
            self.progress_dialog.setValue(value)
            if filename:
                self.progress_dialog.setLabelText(f"処理中: {filename}")

    def handle_normalize_finished(self, success_files, error_files):
        # プログレスダイアログを閉じる
        self.cleanup_progress_dialog()

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
        self.setEnabled(True)

    def cancel_analyze(self):
        if hasattr(self, 'analyze_worker'):
            self.analyze_worker.is_cancelled = True
            self.analyze_worker.wait()
            self.cleanup_progress_dialog()
            self.setEnabled(True)

    def cancel_normalize(self):
        if hasattr(self, 'normalize_worker'):
            self.normalize_worker.is_cancelled = True
            self.normalize_worker.wait()
            self.cleanup_progress_dialog()
            self.setEnabled(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(init_font())
    normalizer = AudioNormalizer()
    normalizer.show()
    sys.exit(app.exec_())
