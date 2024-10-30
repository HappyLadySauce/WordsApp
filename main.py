import json
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, font, Frame, Text, Scrollbar
from tkinter import ttk
from threading import Thread, Event
import random
import os


class WordsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("单词发音及解释")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)  # 添加关闭事件处理

        # 设置窗口的背景颜色
        self.root.configure(bg='#4d6c85')

        # 创建左侧框架用于输入和按钮
        left_frame = Frame(root, bg='#4d6c85')
        left_frame.grid(row=0, column=0, sticky='nsew', padx=10, pady=10)

        # 创建右侧框架用于显示单词表
        self.right_frame = Frame(root, bg='#82c6f8')
        self.right_frame.grid(row=0, column=1, sticky='nsew', padx=10, pady=10)

        # 设置字体为微软雅黑
        self.font = font.Font(family='Microsoft YaHei', size=12)

        # 创建ttk风格
        style = ttk.Style(root)
        style.configure('TButton', font=self.font, background='#82c6f8', foreground='black')
        style.configure('TFrame', background='#b8bdf5')
        style.configure('TLabel', font=self.font)

        # 创建标签和输入框
        ttk.Label(left_frame, text="输入第几个单词开始:", style='TLabel').grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.start_entry = Entry(left_frame)
        self.start_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        ttk.Label(left_frame, text="输入第几个单词结束:", style='TLabel').grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.end_entry = Entry(left_frame)
        self.end_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        # 创建听写间隔输入框和标签
        ttk.Label(left_frame, text="单词发音间隔(s):", style='TLabel').grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.interval_entry = Entry(left_frame)
        self.interval_entry.grid(row=2, column=1, sticky='ew', padx=5, pady=5)

        # 创建按钮
        self.pronounce_button = ttk.Button(left_frame, text="范围顺序发音", command=self.start_pronouncing, style='TButton')
        self.pronounce_button.grid(row=3, column=0, columnspan=2, sticky='ew', padx=5, pady=5)
        self.random_pronounce_button = ttk.Button(left_frame, text="范围随机听写", command=self.random_pronounce, style='TButton')
        self.random_pronounce_button.grid(row=4, column=0, columnspan=2, sticky='ew', padx=5, pady=5)
        self.stop_pronounce_button = ttk.Button(left_frame, text="停止发音", command=self.stop_pronouncing, style='TButton')
        self.stop_pronounce_button.grid(row=5, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        # 创建滚动文本框，并设置字体
        self.text_area = Text(self.right_frame, font=self.font, bg='#ffffff', relief='sunken', wrap='word')
        self.text_area.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        scrollbar = Scrollbar(self.right_frame, command=self.text_area.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.text_area.config(yscrollcommand=scrollbar.set)

        # 创建单词发音输入框和按钮
        ttk.Label(left_frame, text="单词发音(输入单词):", style='TLabel').grid(row=8, column=0, sticky='e', padx=5, pady=5)
        self.word_entry = Entry(left_frame)
        self.word_entry.grid(row=8, column=1, sticky='ew', padx=5, pady=5)
        self.pronounce_word_button = ttk.Button(left_frame, text="发音", command=self.pronounce_word, style='TButton')
        self.pronounce_word_button.grid(row=9, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        # 创建状态标签
        self.status_label = ttk.Label(left_frame, text="等待操作", style='TLabel')
        self.status_label.grid(row=6, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        # 添加选择单词表按钮
        self.select_file_button = ttk.Button(left_frame, text="选择单词表", command=self.select_word_file, style='TButton')
        self.select_file_button.grid(row=7, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        # 初始化单词表文件名
        self.word_file_name = None

        # 加载单词列表
        self.words_dict = None
        self.word_file_name = None
        self.words_loaded = False  # 添加一个标志来跟踪单词列表是否已加载
        self.load_word_list()

        # 打印单词表
        self.print_words()

        # 初始化线程控制
        self.pronounce_thread = None
        self.stop_event = Event()
        self.driver = None

    def load_word_list(self):
        if not self.words_loaded:  # 如果单词列表未加载，尝试加载
            try:
                with open('words.json', 'r', encoding='utf-8') as f:
                    self.words_dict = json.load(f)['words']
                self.print_words()
                self.words_loaded = True  # 标记单词列表已加载
            except (FileNotFoundError, json.JSONDecodeError) as e:
                messagebox.showerror("错误", f"发生错误：{e}")
                self.select_word_file()

    def select_word_file(self):
        file_path = filedialog.askopenfilename(
            initialdir=os.getcwd(),
            title="选择单词表",
            filetypes=(("JSON 文件", "*.json"), ("所有文件", "*.*"))
        )
        if file_path:  # 如果用户选择了文件
            self.load_word_file(file_path)
        else:  # 如果用户关闭了选择文件窗口
            sys.exit(0)  # 退出程序

    def load_word_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.words_dict = json.load(f)['words']
            self.word_file_name = file_path
            self.print_words()
            messagebox.showinfo("成功", "单词表加载成功。")
            self.words_loaded = True  # 标记单词列表已加载
        except (FileNotFoundError, json.JSONDecodeError) as e:
            messagebox.showerror("错误", f"发生错误：{e}")
            sys.exit(0)  # 如果文件未找到或格式不正确，退出程序

    def init_browser(self):
        if self.driver is None or not self.is_browser_open():
            self.status_label.config(text="正在初始化浏览器")
            self.service = Service(EdgeChromiumDriverManager().install())
            self.driver = webdriver.Edge(service=self.service)
            self.status_label.config(text="浏览器初始化完成")

    def is_browser_open(self):
        try:
            self.driver.current_window_handle
            return True
        except:
            return False

    def start_pronouncing(self):
        if self.pronounce_thread is not None and self.pronounce_thread.is_alive():
            self.status_label.config(text="发音线程已在运行...")
            return

        self.init_browser()  # 初始化浏览器

        self.stop_event.clear()  # 清除停止事件

        start_index = self.get_index_from_entry(self.start_entry)
        end_index = self.get_index_from_entry(self.end_entry)

        if start_index is None:
            start_index = 1
        if end_index is None:
            end_index = len(self.words_dict)
        elif end_index < start_index:
            self.status_label.config(text="错误：结束索引必须大于或等于开始索引。")
            return

        words_list = list(self.words_dict.keys())[start_index-1:end_index]
        interval = self.get_interval_from_entry(self.interval_entry)
        if interval is None or interval < 1 or interval > 90:
            interval = 3  # 如果输入的值不合法，使用默认值10秒
        self.pronounce_words(words_list, interval)

    def random_pronounce(self):
        if self.pronounce_thread is not None and self.pronounce_thread.is_alive():
            self.status_label.config(text="发音线程已在运行...")
            return

        self.init_browser()  # 初始化浏览器

        self.stop_event.clear()  # 清除停止事件

        start_index = self.get_index_from_entry(self.start_entry)
        end_index = self.get_index_from_entry(self.end_entry)

        if start_index is None:
            start_index = 1
        if end_index is None:
            end_index = len(self.words_dict)
        elif end_index < start_index:
            self.status_label.config(text="错误：结束索引必须大于或等于开始索引。")
            return

        words_list = list(self.words_dict.keys())[start_index-1:end_index]
        interval = self.get_interval_from_entry(self.interval_entry)
        if interval is None or interval < 1 or interval > 90:
            interval = 10  # 如果输入的值不合法，使用默认值10秒
        random.shuffle(words_list)  # 随机打乱单词顺序
        self.pronounce_words(words_list, interval)

    def stop_pronouncing(self):
        self.stop_event.set()  # 设置停止事件
        self.status_label.config(text="发音已停止。")

    def get_index_from_entry(self, entry):
        try:
            return int(entry.get())
        except ValueError:
            return None  # 不再显示错误消息

    def get_interval_from_entry(self, entry):
        try:
            interval = int(entry.get())
            return interval
        except ValueError:
            return None

    def pronounce_words(self, words, interval=3):
        def pronounce():
            self.status_label.config(text="正在发音...")
            for word in words:
                if self.stop_event.is_set():
                    self.status_label.config(text="发音已停止。")
                    break
                try:
                    self.driver.get(f"https://dict.youdao.com/result?word={word}&lang=en")
                    wait = WebDriverWait(self.driver, 10)
                    pronounce_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "pronounce")))
                    pronounce_button.click()
                    time.sleep(interval)  # 使用用户设定的间隔时间
                except Exception as e:
                    print(f"在单词 '{word}' 发生错误: {e}")
            self.status_label.config(text="发音完成。")
            self.stop_event.clear()

        self.pronounce_thread = Thread(target=pronounce)
        self.pronounce_thread.start()

    def pronounce_word(self):
        word = self.word_entry.get()
        if not word:
            messagebox.showwarning("警告", "请输入一个单词。")
            return
        self.init_browser()  # 初始化浏览器
        self.start_pronounce_word_thread(word)

    def start_pronounce_word_thread(self, word):
        self.pronounce_word_thread = Thread(target=self.pronounce_word_in_thread, args=(word,))
        self.pronounce_word_thread.start()

    def pronounce_word_in_thread(self, word):
        try:
            self.driver.get(f"https://dict.youdao.com/result?word={word}&lang=en")
            wait = WebDriverWait(self.driver, 10)
            pronounce_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "pronounce")))
            pronounce_button.click()
            time.sleep(3)  # 等待发音完成
        except Exception as e:
            messagebox.showerror("错误", f"在单词 '{word}' 发生错误: {e}")
        finally:
            # 注意：这里不关闭浏览器，以符合题目要求
            pass

    def print_words(self):
        self.text_area.delete('1.0', 'end')
        for i, (word, meanings) in enumerate(self.words_dict.items(), start=1):
            explanations = "; ".join([f"{part[:-1]}:{meanings[part]}" for part in meanings])
            self.text_area.insert('end', f"{i}. {word}: {explanations}\n")

    def on_closing(self):
        if self.pronounce_thread is not None and self.pronounce_thread.is_alive():
            self.stop_pronouncing()
        # 不关闭浏览器
        # self.driver.quit()
        self.root.destroy()
        if self.word_file_name:
            # 保存当前选择的单词表文件名到配置文件
            pass


if __name__ == "__main__":
    root = Tk()
    app = WordsApp(root)
    root.mainloop()