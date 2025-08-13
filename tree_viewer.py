import sys
import unicodedata
import ctypes
import os
import subprocess
import shutil
import pythoncom
import psutil
import win32gui
import win32process
import win32con
import csv
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QMenuBar, QTextEdit,
    QTreeWidget,QAction, QFileDialog, QInputDialog, QPushButton, QHBoxLayout, QLineEdit,QTextEdit,
    QMessageBox, QTreeView, QFileSystemModel, QMenu, QSplitter, QLabel, QAbstractItemView, QDialog, QTreeWidgetItem
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QIcon, QPixmap,QTextOption
from PyQt5.QtCore import Qt, QDir, QPoint,QTimer, QEvent
import qdarkstyle

EXCEL_APP = "Ket.Application"
VERSION = "Vβ01"
#EXCEL_APP = "Excel.Application"

def get_east_asian_width_count(text):
    count = 0
    for c in text:
        if unicodedata.east_asian_width(c) in 'FWA':
            count += 2
        else:
            count += 1
    return count

class WindowInfo:
    def __init__(self, hwnd, title, exe_path, icon, category):
        self.hwnd = hwnd
        self.title = title
        self.exe_path = exe_path
        self.icon = icon
        self.category = category

#config登録と保存----------------------------------------------------------------------------------------
CONFIG_FILE = "shortcut.config"

def load_config():
    shortcuts = []
    if not os.path.exists(CONFIG_FILE):
        return shortcuts
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) == 3:
                    shortcuts.append(tuple(row))
    except Exception as e:
        QMessageBox.critical(None, "エラー", f"コンフィグ読み込みエラー:\n{e}")
    return shortcuts

def save_config(shortcuts):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(shortcuts)
    except Exception as e:
        QMessageBox.critical(None, "エラー", f"コンフィグ保存エラー:\n{e}")

class RegisterDialog(QDialog):
    def __init__(self, on_submit, default_base_path):
        super().__init__()
        self.setWindowTitle("ショートカット登録")
        self.resize(400, 200)
        if default_base_path is None:
            default_base_path = os.path.expanduser("~")

        layout = QVBoxLayout()

        # 名前
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("名前:"))
        self.name_edit = QLineEdit()
        name_layout.addWidget(self.name_edit)
        layout.addLayout(name_layout)

        # 分類
        category_layout = QHBoxLayout()
        category_layout.addWidget(QLabel("分類:"))
        self.category_edit = QLineEdit()
        category_layout.addWidget(self.category_edit)
        layout.addLayout(category_layout)

        # フルパス（複数行対応）
        layout.addWidget(QLabel("フルパス:"))
        self.path_edit = QTextEdit()
        self.path_edit.setWordWrapMode(QTextOption.WrapAnywhere)
        self.path_edit.setText(default_base_path)
        self.path_edit.setFixedHeight(50)
        layout.addWidget(self.path_edit)

        # 参照ボタンを下に置く
        browse_btn = QPushButton("参照")
        def browse_file():
            file_path, _ = QFileDialog.getOpenFileName(self, "ファイルを選択", default_base_path)
            if file_path:
                self.path_edit.setText(file_path)
        browse_btn.clicked.connect(browse_file)
        layout.addWidget(browse_btn)


        browse_btn = QPushButton("参照")


        # ボタン
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("決定")
        ok_btn.clicked.connect(lambda: self.submit(on_submit))
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def submit(self, on_submit):
        name = self.name_edit.text().strip()
        category = self.category_edit.text().strip()
        path = self.path_edit.toPlainText().strip()
        if not name or not category or not path:
            QMessageBox.warning(self, "入力エラー", "すべての項目を入力してください。")
            return
        if not os.path.exists(path):
            QMessageBox.warning(self, "パスエラー", f"指定されたパスが存在しません:\n{path}")
            return

        on_submit(name, category, path)
        self.accept()

class FileExplorer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("tree viewer")
        self.setGeometry(0, 0, 340, 1200)




        self.current_path = QDir.homePath()
        self.favorites = []
        self.always_on_top = False
        self.excel_enabled = True
        self.excel_openflag = False
        

        self.excel_tabs_visible = False

        self.excel_cut_string = None
        self.excel_copy_string = None

        self.clipboard_path = None
        self.clipboard_cut = False
        self.excel_app = None
        self.current_workbook = None
        central = QWidget()
        self.setCentralWidget(central)
        self.layout = QVBoxLayout(central)


        top_bar = QSplitter(Qt.Horizontal)
        self.back_button = QPushButton("⬅ 上に戻る")
        self.back_button.setFixedWidth(120)
        self.back_button.setFixedHeight(30)
        self.back_button.clicked.connect(self.go_up)
        self.path_label = QLabel()
        top_bar.addWidget(self.back_button)
        top_bar.addWidget(self.path_label)




        #ツリービューの設定
        self.model = QFileSystemModel()
        self.model.setRootPath(self.current_path)
        self.tree = QTreeView()
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(self.current_path))
        self.tree.setColumnWidth(0, 300)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        self.tree.clicked.connect(self.on_tree_clicked)
        self.tree.doubleClicked.connect(self.on_tree_double_clicked)

        #ツリーアイテムの設定
        self.tree_item = QTreeView()
        self.layout.addWidget(self.tree_item)

        self.models = QStandardItemModel()
        self.models.setHorizontalHeaderLabels(['ウィンドウタイトル'])
        self.tree_item.setModel(self.models)
        self.tree_item.setEditTriggers(QAbstractItemView.NoEditTriggers)

        self.tree_item.viewport().installEventFilter(self)
        
        self.installEventFilter(self)

        self.tree_item.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_item.customContextMenuRequested.connect(self.item_context_menu)

        self.populate_windows()
        self.tree_item.doubleClicked.connect(self.on_item_double_clicked)
        
        self.tree_item.clicked.connect(self.on_item_clicked)
        
        self.tree_item_clickevnt = None

        #ショートカットビューの設定   
        self.shortcuts = load_config()

        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderHidden(True)
        self.tree_widget.itemDoubleClicked.connect(self.open_item)
        self.tree_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_widget.customContextMenuRequested.connect(self.shortcut_menu)
        self.populate_shortcut()



        # スプリッター（上下）
        splitter = QSplitter(Qt.Vertical)
        # ラベル（中間）
        self.label = QLabel("アクティブビュー")


        #self.label.setAlignment(Qt.AlignLeft)  
        #self.label.setMaximumHeight(30)
        #self.label.setStyleSheet("background-color: #f0f0f0; padding: 5px;")

        self.memo = QTextEdit()
        self.memo.setPlaceholderText("ショートカット")
        # スプリッターにウィジェット追加
        splitter.addWidget(self.label)
        splitter.addWidget(self.tree_item)
        splitter.addWidget(top_bar)

        splitter.addWidget(self.tree)
        self.label2 = QLabel("ショートカット")
        splitter.addWidget(self.label2)
        splitter.addWidget(self.tree_widget)
        

        # ラベルが大きくなりすぎないように制限
        self.label2.setMaximumHeight(30)

        # ストレッチ設定（ラベルは固定）
        splitter.setStretchFactor(0, 0)  # ラベル
        splitter.setStretchFactor(1, 2)  # ツリーアイテム
        splitter.setStretchFactor(2, 0)  # ラベル
        splitter.setStretchFactor(3, 3)  # ツリービュー
        splitter.setStretchFactor(4, 0)  # ラベル
        splitter.setStretchFactor(5, 2)  # シュートカットビュー

        self.layout.addWidget(splitter)

        self.menu_bar = QMenuBar()
        self.setMenuBar(self.menu_bar)
        self.setup_menus()
        

        start_path = os.getcwd()
        self.model.setRootPath(start_path)
        self.tree.setRootIndex(self.model.index(start_path))
        self.current_path = start_path
        self.update_path_label()

        # タイマーで定期更新#編集
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_active)
        self.timer.start(5000)  # 10秒ごと
    
    def check_active(self):
        if QApplication.activeWindow() is None:
            self.populate_windows()



    def setup_menus(self):
        file_menu = self.menu_bar.addMenu("ファイル")

        open_dir = QAction("別のディレクトリを開く", self)
        open_dir.triggered.connect(self.select_directory)
        file_menu.addAction(open_dir)


        exit_action = QAction("終了", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        self.favorite_menu = self.menu_bar.addMenu("お気に入り")

        config_save_action = QAction("ショートカットを登録", self)
        config_save_action.triggered.connect(self.open_register_dialog)
        self.favorite_menu.addAction(config_save_action)

        config_edit_action = QAction("ショートカットを編集", self)
        config_edit_action.triggered.connect(
            lambda: os.startfile(CONFIG_FILE)
            if os.path.exists(CONFIG_FILE)
            else QMessageBox.warning(self, "エラー", "コンフィグが存在しません。")
        )
        self.favorite_menu.addAction(config_edit_action)

        config_update_action = QAction("更新", self)
        config_update_action.triggered.connect(self.refresh_from_config)
        self.favorite_menu.addAction(config_update_action)


        tool_menu = self.menu_bar.addMenu("ツール")

        font_size_action = QAction("文字サイズを変更", self)
        font_size_action.triggered.connect(self.change_font_size)
        tool_menu.addAction(font_size_action)

        topmost_action = QAction("常に前面に表示", self, checkable=True)
        topmost_action.setChecked(True) 
        topmost_action.triggered.connect(self.toggle_always_on_top)
        tool_menu.addAction(topmost_action)

        self.toggle_always_on_top(Qt.Checked)

        excel_toggle = QAction("Excelを読み込む", self, checkable=True)
        excel_toggle.setChecked(True) 
        excel_toggle.triggered.connect(self.toggle_excel)
        tool_menu.addAction(excel_toggle)

        ver_menu = self.menu_bar.addMenu("バージョン")
        version_check = QAction(VERSION, self)
        ver_menu.addAction(version_check)

        

    def update_path_label(self):
        max_width = self.width() - 150  # ボタン等を考慮した余白
        metrics = self.path_label.fontMetrics()
        elided = metrics.elidedText(f"📁 : {self.current_path}", Qt.ElideLeft, max_width)
        self.path_label.setText(elided)

    def go_up(self):
        if self.excel_tabs_visible:
            self.populate_tree()
            self.excel_tabs_visible = False
        else:
            parent = os.path.dirname(self.current_path)
            if os.path.exists(parent):
                self.current_path = parent
                self.tree.setRootIndex(self.model.index(self.current_path))
                self.update_path_label()

        if(self.excel_openflag):
            self.excel_openflag = False
            
            #self.excel_app.Quit() #エクセルを消したいときは
            self.excel_app = None


    def on_tree_clicked(self, index):
        self.tree.expand(index)
        if (self.excel_openflag and self.excel_app):
            try:

                target_sheet = self.current_workbook.Sheets(index.data())  # または wb.Sheets(2) など
                target_sheet.Activate()  # これで画面上の表示もそのシートに切り替わる
                self.show_excel_tabs(self.current_workbook)
            except  Exception as e:
                QMessageBox.warning(self, "Excelエラー", f"Excelファイルの処理中にエラーが発生しました:\n{e}")
                self.go_up()

    def on_tree_double_clicked(self, index):
        self.tree.expand(index)
        
        if(not self.excel_openflag and not self.excel_app):
            path = self.model.filePath(index)

            ext = os.path.splitext(path)[1].lower()
            if not os.path.isdir(path):
                if self.excel_enabled and ext in [".xls", ".xlsx"]:
                    self.handle_excel_file(path)
                else:
                    self.open_with_default_app(path)


    def on_tree_load_clicked(self, path):
        
        
        if(not self.excel_openflag and not self.excel_app):
            #path = self.model.filePath(index)
            if os.path.isdir(path):
                self.current_path = path
                self.tree.setRootIndex(self.model.index(path))
                self.update_path_label()
            else:
                ext = os.path.splitext(path)[1].lower()
                if self.excel_enabled and ext in [".xls", ".xlsx"]:
                    self.handle_excel_file(path)
                else:
                    self.open_with_default_app(path)           

    def handle_excel_file(self, path):
        try:
            pythoncom.CoInitialize()
            # すでに起動している場合接続
            self.excel_app = win32com.client.GetActiveObject(EXCEL_APP)
            self.excel_openflag = True
        except Exception:
            pythoncom.CoInitialize()
            # 起動していなければ新しく起動
            self.excel_app = win32com.client.Dispatch(EXCEL_APP)
            self.excel_openflag = True


        try:
            

            wb = None
            for book in self.excel_app.Workbooks:
                if os.path.abspath(book.FullName) == os.path.abspath(path):
                    wb = book
                    break
            if wb is None:
                wb = self.excel_app.Workbooks.Open(path)
                self.excel_app.Visible = True
            self.current_workbook = wb
            self.show_excel_tabs(wb)
        except Exception as e:
            QMessageBox.warning(self, "Excelエラー", f"Excelファイルの処理中にエラーが発生しました:\n{e}")

    def show_excel_tabs(self, workbook):
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(["Excelシート"])
        for sheet in workbook.Sheets:
            item = QStandardItem(sheet.Name)
            item.setData(sheet.Name, Qt.UserRole)
            model.appendRow(item)
        self.tree.setModel(model)
        self.excel_tabs_visible = True

    def activate_excel_sheet(self, index):
        if not self.current_workbook:
            return
        sheet_name = index.data()
        try:
            self.current_workbook.Sheets(sheet_name).Activate()
        except Exception as e:
            QMessageBox.warning(self, "シート切り替えエラー", str(e))

    def open_with_default_app(self, path):
        try:
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            QMessageBox.warning(self, "エラー", f"ファイルを開けませんでした:\n{e}")

    def select_directory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "ディレクトリを選択", self.current_path)
        if dir_path:
            self.current_path = dir_path
            self.tree.setRootIndex(self.model.index(self.current_path))
            self.update_path_label()

    def change_font_size(self):
        size, ok = QInputDialog.getInt(self, "フォントサイズ変更", "新しいフォントサイズ:", 10, 6, 40)
        if ok:
            font = self.tree.font()
            font.setPointSize(size)
            self.tree.setFont(font)
            self.path_label.setFont(font)
            self.tree_item.setFont(font)
            self.tree_widget.setFont(font)
            self.label2.setFont(font)
            self.label.setFont(font)
            self.back_button.setFont(font)

    def toggle_always_on_top(self, checked):
        self.always_on_top = checked
        flags = self.windowFlags()
        if self.always_on_top:
            self.setWindowFlags(flags | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(flags & ~Qt.WindowStaysOnTopHint)
        self.show()

    def toggle_excel(self, checked):
        self.excel_enabled = checked



    def populate_tree(self):
        self.tree.setModel(self.model)
        self.tree.setRootIndex(self.model.index(self.current_path))
        self.update_path_label()
        self.tree.doubleClicked.connect(self.on_tree_double_clicked)

    def show_context_menu(self, position):
        if (not self.excel_openflag and not self.excel_app):
        
            index = self.tree.indexAt(position)
            if not index.isValid():
                return
            path = self.model.filePath(index)
            menu = QMenu()

            open_action = QAction("開く", self)
            open_action.triggered.connect(lambda: self.on_tree_load_clicked(path))
            menu.addAction(open_action)
            
            menu.addSeparator()

            newopen_action = QAction("デフォルトのアプリで開く", self)
            newopen_action.triggered.connect(lambda: self.open_with_default_app(path))
            menu.addAction(newopen_action)

            favorite_action = QAction("ショートカットに追加", self)
            favorite_action.triggered.connect(lambda: self.open_register_dialog(path))
            menu.addAction(favorite_action)

            if os.path.isfile(path):
                rename_action = QAction("名前の変更", self)
                rename_action.triggered.connect(lambda: self.rename_file(path))
                menu.addAction(rename_action)

            if os.path.isfile(path):
                delete_action = QAction("削除", self)
                delete_action.triggered.connect(lambda: self.delete_file(path))
                menu.addAction(delete_action)





            menu.exec_(self.tree.viewport().mapToGlobal(position))
        elif (self.excel_openflag and self.excel_app):
            index = self.tree.indexAt(position)
            if not index.isValid():
                return
            menu = QMenu()
            rename_action = QAction("名前の変更", self)
            rename_action.triggered.connect(lambda: self.rename_excel_tab(index))
            menu.addAction(rename_action)

            menu.addSeparator()

            add_action = QAction("新規シート追加", self)
            add_action.triggered.connect(lambda: self.add_excel_tab())
            menu.addAction(add_action)

            paste_action = QAction("貼付", self)
            paste_action.triggered.connect(lambda: self.paste_excel_tab(index))
            if(not self.excel_cut_string and not self.excel_copy_string):
                paste_action.setEnabled(False) 
            else:
                paste_action.setEnabled(True) 
            menu.addAction(paste_action)


            move_action = QAction("切り取り", self)
            move_action.triggered.connect(lambda: self.move_excel_tab(index))
            menu.addAction(move_action)



            copy_action = QAction("コピー", self)
            copy_action.triggered.connect(lambda: self.copy_excel_tab(index))
            menu.addAction(copy_action)

            delete_action = QAction("削除", self)
            delete_action.triggered.connect(lambda: self.delete_excel_tab(index))
            menu.addAction(delete_action)

            menu.addSeparator()

            save_action = QAction("上書き保存", self)
            save_action.triggered.connect(lambda: self.save_excel_tab())
            menu.addAction(save_action)

            newsave_action = QAction("別名で保存", self)
            newsave_action.triggered.connect(lambda: self.newsave_excel_tab())
            menu.addAction(newsave_action)

            menu.addSeparator()

            exit_action = QAction("excelを閉じる", self)
            exit_action.triggered.connect(lambda: self.exit_excel_tab())
            menu.addAction(exit_action)

            menu.exec_(self.tree.viewport().mapToGlobal(position))

    def delete_file(self, path):
        try:
            os.remove(path)
            QMessageBox.information(self, "削除完了", f"{os.path.basename(path)} を削除しました")
            self.model.refresh()
        except Exception as e:
            QMessageBox.warning(self, "削除失敗", str(e))

    def rename_file(self, path):
        new_name, ok = QInputDialog.getText(self, "名前の変更", "新しい名前:", text=os.path.basename(path))
        if ok and new_name:
            new_path = os.path.join(os.path.dirname(path), new_name)
            try:
                os.rename(path, new_path)
                QMessageBox.information(self, "名前変更", f"{path} → {new_path}")
                self.model.refresh()
            except Exception as e:
                QMessageBox.warning(self, "名前変更失敗", str(e))



    def rename_excel_tab(self, index):
        new_name, ok = QInputDialog.getText(self, "タブの名前の変更", "タブの新しい名前:", text=index.data())

        if(get_east_asian_width_count(new_name)>32):
            QMessageBox.warning(self, "タブは半角文字31文字全角、全角文字15文字までです。")
            return
        if ok and new_name:
            try:
                ws = self.current_workbook.Sheets(index.data())  


                ws.Name = new_name
                self.show_excel_tabs(self.current_workbook)
            except Exception as e:
                QMessageBox.warning(self, "名前変更失敗", str(e))

    def copy_excel_tab(self, index):
        self.excel_cut_string = None
        self.excel_copy_string = index.data()
        print("copy:",self.excel_copy_string," cut:",self.excel_cut_string)


    def move_excel_tab(self, index):
        self.excel_cut_string = index.data() 
        self.excel_copy_string = None      
        print("copy:",self.excel_copy_string," cut:",self.excel_cut_string)



    def paste_excel_tab(self, index):
        try:
            print("copy:",self.excel_copy_string," cut:",self.excel_cut_string)

            if(self.excel_cut_string and not self.excel_copy_string):

                worksheet = self.current_workbook.Sheets(self.excel_cut_string)

                worksheet.Move(Before=None,After=self.current_workbook.Sheets(index.data()))
                self.show_excel_tabs(self.current_workbook)
                self.excel_cut_string = None
                self.excel_copy_string = None


                self.show_excel_tabs(self.current_workbook)
            elif(not self.excel_cut_string and self.excel_copy_string):
                worksheet = self.current_workbook.Sheets(self.excel_copy_string)

                worksheet.Copy(Before=None,After=self.current_workbook.Sheets(index.data()))
                self.show_excel_tabs(self.current_workbook)
                self.excel_cut_string = None
        except Exception as e:
            QMessageBox.warning(self, "名前変更失敗", str(e))

    def add_excel_tab(self):
        new_name, ok = QInputDialog.getText(self, "新規シート挿入", "タブの新しい名前:", text="Sheet1")

        if(get_east_asian_width_count(new_name)>32):
            QMessageBox.warning(self, "タブは半角文字31文字全角、全角文字15文字までです。")
            return
        try:
            sheet_name = self.current_workbook.ActiveSheet.Name
            # 一番左のシートを取得
            ws = self.current_workbook.Sheets(sheet_name)

            # 一番左にシートを追加する
            self.current_workbook.Sheets.Add(Before=ws)

            ws.Name = new_name
            self.show_excel_tabs(self.current_workbook)
        except Exception as e:
            QMessageBox.warning(self, "追加失敗", str(e))
            self.show_excel_tabs(self.current_workbook)


    def delete_excel_tab(self,index):
        try:
            ws = self.current_workbook.Sheets(index.data())
            ws.Delete()
            self.show_excel_tabs(self.current_workbook)
        except Exception as e:
            QMessageBox.warning(self, "削除失敗", str(e))
            self.show_excel_tabs(self.current_workbook)


    def save_excel_tab(self):
        try:
            self.current_workbook.Save()
        except Exception as e:
            QMessageBox.warning(self, "上書き保存失敗", str(e))

    def newsave_excel_tab(self):
        
        try:
            oldfullpath = self.excel_app.ActiveWorkbook.FullName
            newfullpath = show_save_dialog(self.current_workbook.Name)
            self.current_workbook.SaveAs(newfullpath)
            
            self.show_excel_tabs(self.current_workbook)
            self.swich_UI(newfullpath,oldfullpath)
        except Exception as e:
            QMessageBox.warning(self, "新規保存失敗", str(e))

    def swich_UI(self,path,oldpath):
        buttonReply = QMessageBox.question(self, "メッセージ", "新規保存したシートにスイッチしますか？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if buttonReply == QMessageBox.Yes:
            self.handle_excel_file(path) 
        else:
            self.exit_excel_tab()
            self.handle_excel_file(oldpath) 

    def exit_excel_tab(self):

        self.excel_app.Quit()
        self.go_up()


    #上ビュー------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    def changeEvent(self, event):
        if event.type() ==  QEvent.FocusIn:
            if self.isActiveWindow():
                self.populate_windows()
        super().changeEvent(event)
    #'''

    def populate_windows(self):
        self.tree_item_clickevnt = None
        self.models.removeRows(0, self.models.rowCount())
        self.windows = []
        self.category_items = {}

        win32gui.EnumWindows(self.enum_callback, None)

        for win in self.windows:

            display_name = f"{self.get_emoji(win.category)} {win.title}"
            item = QStandardItem(display_name)
            item.setData(win.hwnd, Qt.UserRole)
            item.setData(win.exe_path, Qt.UserRole + 1)
            item.setToolTip(win.exe_path)
            if win.icon:
                item.setIcon(win.icon)

            self.category_items[win.category].appendRow(item)
        for i in range(self.models.rowCount()):
            index = self.models.index(i, 0)
            self.tree_item.expand(index)

    def enum_callback(self, hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if not title or title in ["Program Manager", "Windows Input Experience","Windows 入力エクスペリエンス"]:
            
            return
        
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        try:
            process = psutil.Process(pid)
            exe_path = process.exe()
            exe_name = os.path.basename(exe_path).lower()
        except Exception:
            return

        if exe_name in ["systemsettings.exe", "applicationframehost.exe"]:
            return


        category = self.classify_app(exe_name)
        icon = self.get_icon(exe_path)

        if category not in self.category_items:
            cat_item = QStandardItem(f"{self.get_emoji(category)} {category}")
            cat_item.setEditable(False)

            self.models.appendRow(cat_item)
            
            self.category_items[category] = cat_item

        self.windows.append(WindowInfo(hwnd, title, exe_path, icon, category))

    def classify_app(self, exe_name):
        if "excel" in exe_name:
            return "Excel"
        elif "chrome" in exe_name:
            return "Chrome"
        elif "msedge" in exe_name:
            return "Msedge"
        elif "python" in exe_name:
            return "python"
        elif "et.exe" in exe_name:
            return "Excel"
        elif "vscode" in exe_name or "code" in exe_name:
            return "VSCode"
        elif "photo" in exe_name:
            return "写真"
        elif "paint" in exe_name:
            return "ペイント"
        elif "duino" in exe_name:
            return "プログラム"
        elif "player" in exe_name:
            return "動画"
        elif "explorer" in exe_name:
            return "エクスプローラー"
        elif exe_name.endswith(".exe"):
            return exe_name.replace(".exe", "").capitalize()
        else:
            return "その他"

    def get_emoji(self, category):
        
        emojis = {
            "Excel": "📗",
            "Chrome": "🌐",
            "Msedge":"🌍",
            "VSCode": "🖊️",
            "ペイント":"🖌️",
            "プログラム":"✅",
            "エクスプローラー": "🗂️",
            "音楽": "🎵",
            "写真": "📷",
            "動画": "🎥",
            "PDF": "📕",
            "python":"🏴",
            "その他": "📒"
        }
        return emojis.get(category, "🏷️")
    


    def get_icon(self, exe_path):
        try:
            large = (ctypes.c_void_p * 1)()
            small = (ctypes.c_void_p * 1)()
            if ctypes.windll.shell32.ExtractIconExW(exe_path, 0, large, small, 1) > 0:
                hicon = large[0]
                pixmap = QPixmap.fromWinHICON(hicon)
                ctypes.windll.user32.DestroyIcon(hicon)
                return QIcon(pixmap)
        except Exception:
            pass
        return None


    def on_item_clicked(self, index):

        item = self.models.itemFromIndex(index)
        hwnd = item.data(Qt.UserRole)
        if hwnd:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                self.tree_item_clickevnt = index

                QTimer.singleShot(200, lambda: self.populate_windows())
            except Exception as e:
                print(f"アクティブ化失敗: {e}")

    def on_item_double_clicked(self,index):
        


        item = self.models.itemFromIndex(index)
        
        hwnd = item.data(Qt.UserRole)
        exe_path = item.data(Qt.UserRole + 1)  # exe_pathを格納しておいた

        if hwnd:
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)       
            except Exception as e:
                print(f"アクティブ化失敗: {e}")

        if("et.exe" in exe_path):   

            excel = win32com.client.Dispatch(EXCEL_APP)

            # アクティブなブックを取得
            wb = excel.ActiveWorkbook

            if wb is not None:
                # フルパスを取得
                path = wb.FullName
                

            ext = os.path.splitext(path)[1].lower()
            if not os.path.isdir(path):
                    if self.excel_enabled and ext in [".xls", ".xlsx"]:
                        self.handle_excel_file(path)
            self.populate_windows()
        if("explorer.exe" in exe_path):
            shell = win32com.client.Dispatch("Shell.Application")
            windows = shell.Windows()

            # アクティブウィンドウのハンドルを取得
            fg_hwnd = win32gui.GetForegroundWindow()

            for window in windows:
                
                    # Explorer ウィンドウのみ
                    if window.Name not in ("エクスプローラー", "Explorer"):
                        continue
                    
                    # ウィンドウのHWNDを取得
                    hwnd = window.HWND

                    # フォーカスのあるウィンドウと一致するか判定
                    if hwnd == fg_hwnd:
                        folder = window.Document.Folder
                        folder_path = folder.Self.Path



    def on_item_load_clicked(self,index):
        try:        
            item = self.models.itemFromIndex(index)
            
            hwnd = item.data(Qt.UserRole)
            exe_path = item.data(Qt.UserRole + 1) 
            if hwnd:
                
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)       

            if("et.exe" in exe_path):   

                excel = win32com.client.Dispatch(EXCEL_APP)

                # アクティブなブックを取得
                wb = excel.ActiveWorkbook

                if wb is not None:
                    # フルパスを取得
                    path = wb.FullName

                ext = os.path.splitext(path)[1].lower()
                if not os.path.isdir(path):
                        if self.excel_enabled and ext in [".xls", ".xlsx"]:
                            self.handle_excel_file(path)
        
            if("explorer.exe" in exe_path):
                shell = win32com.client.Dispatch("Shell.Application")
                windows = shell.Windows()

                # アクティブウィンドウのハンドルを取得
                fg_hwnd = win32gui.GetForegroundWindow()

                for window in windows:
                    
                        # Explorer ウィンドウのみ
                        if window.Name not in ("エクスプローラー", "Explorer"):
                            continue
                        
                        # ウィンドウのHWNDを取得
                        hwnd = window.HWND

                        # フォーカスのあるウィンドウと一致するか判定
                        if hwnd == fg_hwnd:
                            folder = window.Document.Folder
                            folder_path = folder.Self.Path
                if os.path.exists(folder_path):
                    self.current_path = folder_path
                    self.tree.setRootIndex(self.model.index(folder_path))
                    self.update_path_label()
            self.populate_windows()
        except Exception as e:
            print(f"アクティブ化失敗: {e}")



    def on_item_close(self,index):
        try:
            item = self.models.itemFromIndex(index)
            hwnd = item.data(Qt.UserRole)
            exe_path = item.data(Qt.UserRole + 1)  # exe_pathを格納しておいた
            if hwnd:
                
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)       



                    hwnd = win32gui.GetForegroundWindow()
                    if hwnd:
                        # WM_CLOSEメッセージを送ってウィンドウを閉じる
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        return True
                    return False        
        except Exception as e:
                    print(f"アクティブ化失敗: {e}")
        self.populate_windows()

    def item_context_menu(self,position):#編集
        index = self.tree_item.indexAt(position)
        if not index.isValid():
                return
        menu = QMenu()

        open_action = QAction("ツリービューに反映", self)
        
        open_action.triggered.connect(lambda: self.on_item_load_clicked(index))
        menu.addAction(open_action)
        
        menu.addSeparator()
        
        close_action = QAction("閉じる", self)
        close_action.triggered.connect(lambda: self.on_item_close(index))
        menu.addAction(close_action)

        menu.exec_(self.tree_item.viewport().mapToGlobal(position))
    




    #ここからショートカットビュー--------------------------------------------------------------------
    def create_menu(self):
        menubar = self.menuBar()
        config_menu = menubar.addMenu("コンフィグ")

        open_action = QAction("コンフィグを開く", self)
        open_action.triggered.connect(
            lambda: os.startfile(CONFIG_FILE)
            if os.path.exists(CONFIG_FILE)
            else QMessageBox.warning(self, "エラー", "コンフィグが存在しません。")
        )
        config_menu.addAction(open_action)

    def open_register_dialog(self,path):
        if not(path):
            path = r"C:/"
        def on_submit(name, category, path):
            
            self.shortcuts.append((name, category, path))
            save_config(self.shortcuts)
            self.populate_shortcut()
            QMessageBox.information(self, "登録完了", "ショートカットを登録しました。")

        dialog = RegisterDialog(on_submit,path)
        dialog.exec_()

    def populate_shortcut(self):
        self.tree_widget.clear()
        categories = {}
        for name, category, path in self.shortcuts:
            if category not in categories:
                cat_item = QTreeWidgetItem([f" {category}"])
                categories[category] = cat_item
                self.tree_widget.addTopLevelItem(cat_item)

            ext = os.path.splitext(path)[1].lower()
            rogid = self.classify_path(ext)
            icon=self.get_emoji(rogid)
            #icon = ICON_MAP.get(ext, "📦")
            item = QTreeWidgetItem([f"{icon} {name}"])
            item.setData(0, Qt.UserRole, path)
            categories[category].addChild(item)

    def refresh_from_config(self):
        try:
            self.shortcuts = load_config()
            self.populate_shortcut()
            #QMessageBox.information(self, "反映完了", "コンフィグを再読み込みしました。")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"コンフィグ読み込みエラー:\n{e}")

    def open_item(self, item, column):
        path = item.data(0, Qt.UserRole)
        if path:
            try:
                if(path in [".xls", ".xlsx"] or os.path.exists(path)):
                    self.on_tree_load_clicked(path)
                else:
                    os.startfile(path)
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"ファイルを開けません:\n{e}")


    def classify_path(self, exe_name):
        if exe_name in [".xls", ".xlsx"] :
            return "Excel"
        elif ".html" in exe_name:
            return "Msedge"
        elif ".py" in exe_name:
            return "python"
        elif ".ino" in exe_name:
            return "VSCode"
        elif  exe_name in [".png",".jpg"]:
            return "写真"
        elif  exe_name in [".mp4",".mav",".mkv"]:
            return "動画"
        elif ".pdf" in exe_name:
            return "PDF"
        elif exe_name.endswith(".exe"):
            return exe_name.replace(".exe", "").capitalize()
        else:
            return "エクスプローラー"
    
    def shortcut_menu(self, position):
        index = self.tree_widget.indexAt(position)
        if not index.isValid():
            return
        path = index.data(Qt.UserRole)

        menu = QMenu()
        if(path in [".xls", ".xlsx"] or os.path.exists(path)):

            open_action = QAction("ツリービューに反映", self)
            open_action.triggered.connect(lambda: self.on_tree_load_clicked(path))
            menu.addAction(open_action)
            
            menu.addSeparator()
        
        newopen_action = QAction("デフォルトのアプリで開く", self)
        newopen_action.triggered.connect(lambda: os.startfile(path))
        menu.addAction(newopen_action)

        menu.exec_(self.tree_widget.viewport().mapToGlobal(position))



def show_save_dialog(defult_name):
    window_saveas = QWidget()

    # 名前を付けて保存ダイアログを表示
    file_path, _ = QFileDialog.getSaveFileName(
        parent=window_saveas,
        caption="名前を付けて保存",
        directory=defult_name,
        filter="Excelファイル (*.xlsx *.xls);;CSVファイル (*.csv);;すべてのファイル (*.*)"
    )

    if file_path:
        print("保存先パス（絶対パス）:", file_path)
        return file_path
    else:
        print("キャンセルされました")
        return None



if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet())
    window = FileExplorer()
    window.show()
    sys.exit(app.exec_())
