# 导入 tkinter 库，用于创建 GUI 应用程序
import tkinter as tk
# 从 tkinter 中导入文件对话框和消息框模块
from tkinter import filedialog, messagebox
# 从 utils 模块中导入读取 Excel 文件的函数
from utils import read_excel
# 导入 requests 库，用于发送 HTTP 请求
import requests
# 导入 threading 模块，用于实现多线程处理
import threading
# 从 config 模块中导入 DeepSeek API 的密钥
from config import DEEPSEEK_API_KEY

class ExcelChatApp:
    def __init__(self, root):
        """
        初始化应用程序界面组件和布局
        :param root: tkinter 根窗口对象
        """
        self.root = root
        # 设置窗口标题
        self.root.title("DeepSeek Excel Chat")
        # 设置窗口大小
        self.root.geometry("500x400")

        # 主框架容器，用于布局其他组件
        main_frame = tk.Frame(root, padx=20, pady=20)
        # 使主框架填充整个窗口并可扩展
        main_frame.pack(expand=True, fill=tk.BOTH)

        # 文件选择组件框架
        file_frame = tk.Frame(main_frame)
        # 使文件选择组件框架水平填充
        file_frame.pack(fill=tk.X, pady=5)
        
        # 创建文件选择提示标签
        self.label = tk.Label(file_frame, text="Select an Excel file:")
        # 将标签放置在框架左侧
        self.label.pack(side=tk.LEFT)
        
        # 创建打开 Excel 文件的按钮，点击时调用 open_file 方法
        self.open_button = tk.Button(file_frame, text="Open Excel", command=self.open_file)
        # 将按钮放置在框架右侧
        self.open_button.pack(side=tk.RIGHT)

        # 问题输入组件，用于用户输入问题
        self.query_entry = tk.Entry(main_frame, width=50)
        # 将问题输入框水平填充并添加垂直间距
        self.query_entry.pack(pady=10, fill=tk.X)
        
        # 创建提交问题的按钮，点击时调用 ask_deepseek 方法
        self.submit_button = tk.Button(main_frame, text="Ask DeepSeek", command=self.ask_deepseek)
        # 将提交按钮添加垂直间距
        self.submit_button.pack(pady=5)

        # 结果显示组件框架，带有边框样式
        result_frame = tk.Frame(main_frame, relief=tk.GROOVE, borderwidth=2)
        # 使结果显示框架填充整个窗口并可扩展
        result_frame.pack(expand=True, fill=tk.BOTH)
        
        # 创建结果显示标签，初始显示提示信息
        self.result_label = tk.Label(
            result_frame, 
            text="Answer will be shown here...",
            wraplength=400,
            justify=tk.LEFT,
            anchor=tk.NW
        )
        # 将结果显示标签添加内边距并填充整个框架
        self.result_label.pack(padx=5, pady=5, expand=True, fill=tk.BOTH)

    def open_file(self):
        """
        打开文件选择对话框并验证文件类型
        """
        # 打开文件选择对话框，只允许选择 Excel 文件
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                # 验证文件是否可读，传入 cell_range 参数
                cell_range = "A1:B10"  # 假设的 cell range
                read_excel(filename, cell_range)
                # 保存选中的文件名
                self.filename = filename
                # 显示文件加载成功的消息框
                messagebox.showinfo("File Selected", f"Successfully loaded:\n{filename}")
            except Exception as e:
                # 显示文件读取失败的消息框
                messagebox.showerror("Error", f"Failed to read file:\n{str(e)}")
    def _process_query(self, question):
        """
        后台处理查询逻辑
        :param question: 用户输入的问题
        """
        try:
            # 读取 Excel 数据（根据需求调整数据处理逻辑）
            cell_range = "A1:B10"  # 假设的 cell range
            data = read_excel(self.filename, cell_range)
            
            # 构造包含数据的提示（示例）
            # 将列表转换为字符串表示
            data_str = "\n".join(str(row) for row in data)
            enhanced_question = f"{question}\n\nExcel Data Summary:\n{data_str}"
            
            # 获取 API 响应
            response = query_deepseek(
                api_key=DEEPSEEK_API_KEY,
                question=enhanced_question
            )
            
            # 更新 UI
            self.root.after(0, lambda: self._update_ui(response))
        except Exception as e:
            # 显示错误信息
            self.root.after(0, lambda: self._show_error(str(e)))
        finally:
            # 启用提交按钮
            self.root.after(0, lambda: self.submit_button.config(state=tk.NORMAL))

    def ask_deepseek(self):
        """
        处理用户提问请求
        """
        # 获取用户输入的问题并去除首尾空格
        question = self.query_entry.get().strip()
        
        if not question:
            # 如果问题为空，显示警告消息框
            messagebox.showwarning("Empty Question", "Please enter your question.")
            return
            
        if not hasattr(self, 'filename'):
            # 如果没有选择文件，显示警告消息框
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return

        # 在问题前添加中文回答提示
        question = f"{question} 请用中文回答。"

        # 禁用提交按钮，避免重复请求
        self.submit_button.config(state=tk.DISABLED)
        # 更新结果显示标签为处理中提示信息
        self.result_label.config(text="Processing your question...")

        # 在独立线程中处理 API 请求
        threading.Thread(target=self._process_query, args=(question,), daemon=True).start()

    def _update_ui(self, response):
        """
        更新结果显示
        :param response: API 返回的响应内容
        """
        # 更新结果显示标签为 API 响应内容
        self.result_label.config(text=f"Answer:\n\n{response}")

    def _show_error(self, message):
        """
        显示错误信息
        :param message: 错误消息内容
        """
        # 显示错误消息框
        messagebox.showerror("Error", message)
        # 更新结果显示标签为错误提示信息
        self.result_label.config(text="Error occurred. Please try again.")

def query_deepseek(api_key, question):
    """
    向 DeepSeek API 发送查询请求并处理响应
    :param api_key: DeepSeek API 的密钥
    :param question: 用户输入的问题
    :return: API 返回的响应内容
    """
    # DeepSeek API 的请求 URL
    url = "https://api.deepseek.com/chat/completions"
    # 请求头，包含授权信息和内容类型
    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    
    # 请求体，包含模型名称、消息内容、温度参数以及语言偏好
    payload = {
        "model": "deepseek-chat",
        "messages": [{
            "role": "user",
            "content": question
        }],
        "temperature": 0.7,
        "language": "zh-CN"  # 明确指定语言为中文
    }

    try:
        # 发送 POST 请求
        response = requests.post(
            url,
            json=payload,
            headers=headers,
            timeout=15
        )
        # 检查响应状态码
        response.raise_for_status()
        
        # 解析响应为 JSON 格式
        result = response.json()
        if 'choices' in result and len(result['choices']) > 0:
            # 返回 API 响应中的消息内容
            return result['choices'][0]['message']['content']
        # 如果没有有效响应，返回提示信息
        return "No valid response from API."
        
    except requests.exceptions.RequestException as e:
        # 处理网络请求异常
        raise Exception(f"Network error: {str(e)}")
    except (KeyError, IndexError) as e:
        # 处理 API 响应格式异常
        raise Exception(f"Invalid API response format: {str(e)}")
    except Exception as e:
        # 处理其他异常
        raise Exception(f"Unexpected error: {str(e)}")
    
if __name__ == "__main__":
    # 创建 tkinter 根窗口
    root = tk.Tk()
    # 创建 ExcelChatApp 应用实例
    app = ExcelChatApp(root)
    # 进入主事件循环
    root.mainloop()
