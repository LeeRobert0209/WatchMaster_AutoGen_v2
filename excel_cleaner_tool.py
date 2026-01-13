import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD
import win32com.client
import pythoncom

class PsdProcessor:
    def __init__(self, log_callback):
        self.log = log_callback
        self.app = None

    def connect_photoshop(self):
        try:
            self.app = win32com.client.Dispatch("Photoshop.Application")
            return True
        except Exception as e:
            self.log(f"无法连接到 Photoshop: {e}")
            return False

    def find_layer(self, parent, layer_name):
        """Recursively find a layer by name."""
        try:
            # First pass: direct children match
            for layer in parent.Layers:
                if layer.Name == layer_name:
                    return layer
            
            # Second pass: go deep into groups
            for layer in parent.Layers:
                if layer.TypeName == "LayerSet": # It's a group
                    found = self.find_layer(layer, layer_name)
                    if found:
                        return found
        except Exception:
            pass # Handle cases where layers might not be accessible
        return None

    def update_text_layer(self, doc, layer_name, text, width_px=None):
        layer = self.find_layer(doc, layer_name)
        
        # FIX: If we found a Group (LayerSet) instead of a Layer, try to find the layer INSIDE the group
        # This handles the case where there is a Group named "匠龄" containing a Text Layer named "匠龄"
        if layer and hasattr(layer, 'TypeName') and layer.TypeName == "LayerSet":
             # Try to find the actual layer inside this group
             # We assume the text layer inside has the SAME name, or we just look for it recursively
             inner_layer = self.find_layer(layer, layer_name)
             if inner_layer:
                 layer = inner_layer
             else:
                 # If exact name not found inside, maybe just pick the first Text Layer inside?
                 # Dangerous, but better than failing. Let's stick to exact name first.
                 pass

        if layer:
            try:
                if hasattr(layer, 'TypeName') and layer.TypeName == "LayerSet":
                     self.log(f"警告: 找到的图层 '{layer_name}' 仍是一个图层组，无法修改。")
                     return False
                
                if layer.Kind == 2: # 2 = Text Layer
                    try:
                        # Activate the layer first!
                        # Modifying properties like Kind/Width for layers inside groups often fails 
                        # if the layer is not the active layer.
                        doc.ActiveLayer = layer
                        text_item = layer.TextItem
                        
                        # Reverted Punctuation Hack:
                        # Relying on Photoshop's native "Adobe World-Ready Paragraph Composer" and "Kinsoku Shori"
                        # settings in the template is the correct way to handle line-start punctuation.
                        text_item.Contents = text
                        
                        if width_px: 
                            if text_item.Kind != 2:
                                text_item.Kind = 2 
                            
                            # CRITICAL: Reverting to unit conversion logic which worked for descriptions.
                            # Calculate width in Points (1/72 inch) relative to Document DPI manually 
                            # Formula: pt = px * 72 / dpi
                            resolution = doc.Resolution
                            width_pt = width_px * 72 / resolution
                            
                            # Set calculated Width
                            text_item.Width = width_pt
                            
                            # Safe Height
                            if text_item.Height < width_pt / 4: 
                                text_item.Height = width_pt * 2 
                                
                        return True
                    except Exception as e:
                         self.log(f"修改图层 '{layer_name}' 出错: {e}")
                else:
                    self.log(f"警告: 图层 '{layer_name}' 不是文本图层。")
            except Exception as e:
                self.log(f"错误: 操作图层 '{layer_name}' 失败: {e}")
        else:
            self.log(f"警告: 未找到图层 '{layer_name}'。")
        return False

    def process_batch(self, excel_path, template_path, output_dir):
        pythoncom.CoInitialize() # Required for COM in thread
        
        doc = None
        try:
            if not self.connect_photoshop():
                return
            
            # --- Force Preferences ---
            # 1 = Pixels, 2 = Points, 3 = CM
            try:
                self.app.Preferences.RulerUnits = 1 
                self.app.Preferences.TypeUnits = 1
            except:
                pass

            # --- Smart Header Detection (Same as Step 1) ---
            self.log(f"读取 Excel 数据: {excel_path}")
            # Read first few lines without header to find the real header row
            df_temp = pd.read_excel(excel_path, header=None, nrows=10)
            header_row_idx = -1
            
            # Look for a row containing "姓名" and "门店"
            for idx, row in df_temp.iterrows():
                row_str = " ".join([str(x) for x in row.values])
                if "姓名" in row_str and "门店" in row_str:
                    header_row_idx = idx
                    break
            
            if header_row_idx != -1:
                self.log(f"自动检测到表头在第 {header_row_idx + 1} 行")
                df = pd.read_excel(excel_path, header=header_row_idx)
            else:
                self.log("未检测到标准表头，尝试默认设置 (header=0)...")
                df = pd.read_excel(excel_path, header=0) # Cleaned file usually has header at 0

            # Handle Merged Cells for '门店' (Forward Fill) - Safety net
            if "门店" in df.columns:
                df["门店"] = df["门店"].fillna(method='ffill')
            
            # Verify columns - STRICT CHECK
            # We must ensure the user is using the CLEANED file, which has "描述1", "匠人独白", etc.
            required_cols = ["姓名", "门店"]
            processed_cols = ["描述1", "匠人独白", "标题1"]
            
            missing_basic = [c for c in required_cols if c not in df.columns]
            if missing_basic:
                self.log(f"错误: Excel 缺少基础列 {missing_basic}")
                messagebox.showerror("文件错误", f"所选 Excel 缺少必要列: {missing_basic}\n请检查文件格式。")
                return

            missing_processed = [c for c in processed_cols if c not in df.columns]
            if missing_processed:
                self.log(f"错误: Excel 缺少清洗后的数据列 {missing_processed}")
                self.log("提示: 您似乎选择了原始数据文件？请选择步骤1生成的 '_清洗版.xlsx' 文件。")
                messagebox.showerror("选错文件了？", 
                    f"检测到 Excel 文件缺少 {missing_processed} 等清洗列。\n\n"
                    "您可能选择了【原始 Excel】文件！\n"
                    "请务必选择步骤 1 生成的【_清洗版.xlsx】文件进行生成。"
                )
                return

            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            total = len(df)
            self.log(f"开始处理 {total} 个维修师数据...")
            
            self.log(f"打开模板: {template_path}")
            doc = self.app.Open(template_path)
            
            # Re-apply Preferences AFTER opening doc just in case
            try:
                self.app.Preferences.RulerUnits = 1 
                self.app.Preferences.TypeUnits = 1
            except:
                pass
            
            self.log(f"当前文档分辨率: {doc.Resolution} DPI")

            count = 0
            for idx, row in df.iterrows():
                name = str(row.get("姓名", "")).strip()
                store = str(row.get("门店", "")).strip()
                
                if not name: 
                    self.log(f"跳过第 {idx+1} 行: 姓名为空")
                    continue
                    
                target_filename = f"{store}_{name}.psd"
                target_filename = re.sub(r'[\\/*?:"<>|]', "", target_filename)
                save_path = os.path.join(output_dir, target_filename)
                
                self.log(f"[{idx+1}/{total}] 处理: {name} @ {store}")
                
                # Mapping dict: Excel Column -> (Layer Name, Width)
                # Width=None means no change/constraint
                # Since the template already has Paragraph Text boxes set up, we should avoid resetting widths 
                # to prevent unit conversion errors.
                mapping = {
                    "姓名": ("姓名", None),
                    "匠龄": ("匠龄", None),
                    "标题1": ("标题1", None),
                    "描述1": ("描述1", None),
                    "标题2": ("标题2", None),
                    "描述2": ("描述2", None),
                    "标题3": ("标题3", None),
                    "描述3": ("描述3", None),
                    "匠人独白": ("匠人独白", None)
                }

                for col_name, (layer_name, width_val) in mapping.items():
                    if col_name in df.columns:
                        content = str(row[col_name]).strip()
                        if content == "nan": content = ""
                        self.update_text_layer(doc, layer_name, content, width_px=width_val)
                
                # Save as PSD Copy
                # FIX: Correct ProgID is "Photoshop.PhotoshopSaveOptions"
                try:
                    options = win32com.client.Dispatch("Photoshop.PhotoshopSaveOptions")
                    options.EmbedColorProfile = True
                    options.AlphaChannels = True
                    options.Layers = True
                    doc.SaveAs(save_path, options, True) # True = asCopy
                except Exception as save_err:
                    self.log(f"保存失败 ({target_filename}): {save_err}")
                    # Fallback: maybe try saving without options if dispatch fails?
                
                count += 1
            
            self.log(f"处理完成！成功生成 {count} 个文件。")
            self.log(f"保存位置: {output_dir}")
            messagebox.showinfo("完成", f"PSD 批量生成完成！\n共生成 {count} 个文件。\n位置: {output_dir}")

        except Exception as e:
            self.log(f"批量处理出错: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            if doc:
                try:
                    doc.Close(2) # 2 = ppDoNotSaveChanges
                except:
                    pass
            pythoncom.CoUninitialize()


class ExcelCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("智能设计工坊")
        self.root.geometry("700x760")
        
        # Base paths
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.psd_tool = PsdProcessor(self.log)

        # UI Setup
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="智能设计工坊", font=("Microsoft YaHei", 18, "bold"), fg="#333").pack(pady=(0, 20))

        # === Step 1: Data Cleaning ===
        step1_frame = tk.LabelFrame(main_frame, text="步骤 1: 数据清洗", font=("Microsoft YaHei", 10, "bold"), fg="#2E7D32", bg="#F1F8E9", padx=10, pady=10)
        step1_frame.pack(fill=tk.X, pady=10)
        
        self.drop_frame1 = tk.Frame(step1_frame, bg="white", bd=1, relief="solid")
        self.drop_frame1.pack(fill=tk.X, ipady=15, pady=5)
        
        self.drop_label1 = tk.Label(
            self.drop_frame1, 
            text="拖拽原始 Excel 文件到此处 或 点击浏览", 
            font=("Microsoft YaHei", 10),
            bg="white", fg="#555"
        )
        self.drop_label1.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_frame1.drop_target_register(DND_FILES)
        self.drop_frame1.dnd_bind('<<Drop>>', self.handle_drop_clean)
        self.drop_frame1.bind("<Button-1>", lambda e: self.browse_file_clean())
        self.drop_label1.bind("<Button-1>", lambda e: self.browse_file_clean())

        self.clean_file_var = tk.StringVar()
        tk.Label(step1_frame, textvariable=self.clean_file_var, bg="#F1F8E9", fg="gray", font=("SimSun", 9)).pack(fill=tk.X)
        
        tk.Label(step1_frame, text="* 导出文件将自动保存在原文件同目录下 (_清洗版.xlsx)", bg="#F1F8E9", fg="#666", font=("Microsoft YaHei", 8)).pack(anchor="w", padx=5, pady=(0, 5))

        self.btn_clean = tk.Button(step1_frame, text="清洗数据并导出 Excel", command=self.start_cleaning, 
                                     bg="#4CAF50", fg="white", font=("Microsoft YaHei", 10, "bold"), height=2)
        self.btn_clean.pack(fill=tk.X, pady=5)

        # === Step 2: PSD Generation ===
        step2_frame = tk.LabelFrame(main_frame, text="步骤 2: PSD 批量生成", font=("Microsoft YaHei", 10, "bold"), fg="#1565C0", bg="#E3F2FD", padx=10, pady=10)
        step2_frame.pack(fill=tk.X, pady=10)

        # Replaced input frame with Drop Zone
        self.drop_frame2 = tk.Frame(step2_frame, bg="white", bd=1, relief="solid")
        self.drop_frame2.pack(fill=tk.X, ipady=40, pady=5)
        
        self.drop_label2 = tk.Label(
            self.drop_frame2, 
            text="拖拽清洗后的 Excel 文件到此处 或 点击浏览\n(步骤1完成后将自动填入)", 
            font=("Microsoft YaHei", 10),
            bg="white", fg="#555"
        )
        self.drop_label2.place(relx=0.5, rely=0.5, anchor="center")

        self.drop_frame2.drop_target_register(DND_FILES)
        self.drop_frame2.dnd_bind('<<Drop>>', self.handle_drop_psd)
        self.drop_frame2.bind("<Button-1>", lambda e: self.browse_file_psd())
        self.drop_label2.bind("<Button-1>", lambda e: self.browse_file_psd())

        self.psd_input_var = tk.StringVar()
        # Path label (subtle)
        tk.Label(step2_frame, textvariable=self.psd_input_var, bg="#E3F2FD", fg="gray", font=("SimSun", 8)).pack(fill=tk.X)

        # === Status Indicators Area ===
        status_frame = tk.Frame(step2_frame, bg="#E3F2FD")
        status_frame.pack(fill=tk.X, pady=5)

        # 1. Data Status
        self.data_status_var = tk.StringVar(value="❌ 未选择数据文件")
        self.lbl_data_status = tk.Label(status_frame, textvariable=self.data_status_var, bg="#E3F2FD", fg="red", anchor="w", font=("Microsoft YaHei", 9, "bold"))
        self.lbl_data_status.pack(fill=tk.X)

        # 2. Template Status
        self.template_status_var = tk.StringVar(value="检查中...")
        self.lbl_template_status = tk.Label(status_frame, textvariable=self.template_status_var, bg="#E3F2FD", anchor="w", font=("Microsoft YaHei", 9))
        self.lbl_template_status.pack(fill=tk.X, pady=(2,0))
        
        # Start periodic check
        self.check_template_status()

        out_dir = os.path.join(self.base_dir, 'output_psds')
        tk.Label(step2_frame, text=f"* PSD 导出位置: {out_dir}", bg="#E3F2FD", fg="#666", font=("Microsoft YaHei", 8)).pack(anchor="w", padx=5, pady=(5, 5))

        self.btn_gen_psd = tk.Button(step2_frame, text="启动 Photoshop 批量生成", command=self.start_psd_gen, 
                                     bg="#1976D2", fg="white", font=("Microsoft YaHei", 10, "bold"), height=2)
        self.btn_gen_psd.pack(fill=tk.X, pady=5)


        # === Logging ===
        tk.Label(main_frame, text="系统日志:", font=("Microsoft YaHei", 9)).pack(anchor="w", pady=(10,0))
        self.log_text = scrolledtext.ScrolledText(main_frame, height=6, state='disabled', font=("SimSun", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log("系统就绪。请从步骤 1 开始。")

    def check_template_status(self):
        """Periodically checks if the template file exists AND auto-detects data file if empty."""
        # 1. Check Template
        template_path = os.path.join(self.base_dir, 'model', '维修师-模板.psd')
        if os.path.exists(template_path):
            filename = os.path.basename(template_path)
            self.template_status_var.set(f"✅ 模板已就绪: {filename}")
            self.lbl_template_status.config(fg="green")
        else:
            self.template_status_var.set("❌ 未找到模板 (model/维修师-模板.psd)")
            self.lbl_template_status.config(fg="red")
            
        # 2. Check Data Source (Only if input is empty)
        if not self.psd_input_var.get():
            data_dir = os.path.join(self.base_dir, 'data')
            if os.path.exists(data_dir):
                try:
                    # Find all files ending with '_清洗版.xlsx' but EXCLUDE temporary files (~$)
                    files = [f for f in os.listdir(data_dir) 
                             if f.endswith('_清洗版.xlsx') and not f.startswith('~$')]
                    if files:
                        # Pick the most recently modified one
                        full_paths = [os.path.join(data_dir, f) for f in files]
                        latest_file = max(full_paths, key=os.path.getmtime)
                        
                        # Auto-fill
                        self.psd_input_var.set(latest_file)
                        self.update_data_status(latest_file)
                        self.log(f"自动检测到已就绪的数据文件: {os.path.basename(latest_file)}")
                except Exception:
                    pass
        
        # Check again in 2 seconds
        if self.root:
            self.root.after(2000, self.check_template_status)

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        if self.root: self.root.update()
    
    def update_data_status(self, file_path):
        """Updates the data status label based on the file path."""
        if not file_path:
            self.data_status_var.set("❌ 未选择数据文件")
            self.lbl_data_status.config(fg="red")
            return

        if not os.path.exists(file_path):
            self.data_status_var.set("❌ 文件不存在")
            self.lbl_data_status.config(fg="red")
            return

        if not file_path.lower().endswith(('.xlsx', '.xls')):
            self.data_status_var.set("⚠️ 格式错误 (非 Excel 文件)")
            self.lbl_data_status.config(fg="#FF9800") # Orange for warning
            return

        filename = os.path.basename(file_path)
        self.data_status_var.set(f"✅ 数据源已就绪: {filename}")
        self.lbl_data_status.config(fg="green")


    # --- Step 1 Handlers ---
    def handle_drop_clean(self, event):
        file_path = self.parse_path(event.data)
        self.clean_file_var.set(file_path)
        self.log(f"[清洗] 已选择文件: {file_path}")

    def browse_file_clean(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.clean_file_var.set(filename)
            self.log(f"[清洗] 已选择文件: {filename}")

    def start_cleaning(self):
        input_path = self.clean_file_var.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("提示", "请先选择有效的Excel文件！")
            return
        
        self.btn_clean.config(state='disabled')
        self.log("正在启动数据清洗任务...")
        thread = threading.Thread(target=self.cleaning_logic, args=(input_path,))
        thread.start()

    def cleaning_logic(self, input_path):
        try:
            self.log(f"正在读取: {os.path.basename(input_path)}")
            
            # --- Smart Header Detection ---
            # Read first 10 rows without header to scan for keywords
            df_temp = pd.read_excel(input_path, header=None, nrows=10)
            header_row_idx = -1
            
            # Keywords to identify the header row
            required_keywords = ["姓名", "门店"] 
            
            for idx, row in df_temp.iterrows():
                row_str = " ".join([str(x) for x in row.values])
                if all(k in row_str for k in required_keywords):
                    header_row_idx = idx
                    break
            
            if header_row_idx != -1:
                self.log(f"自动检测到表头在第 {header_row_idx + 1} 行")
                df = pd.read_excel(input_path, header=header_row_idx)
            else:
                self.log("⚠️ 未检测到标准表头(姓名/门店)，尝试默认位置 (header=2)...")
                df = pd.read_excel(input_path, header=2)

            # --- Flexible Column Mapping ---
            # Identify columns by name rather than fixed index
            col_map = {}
            for col in df.columns:
                c_str = str(col).strip()
                if "姓名" in c_str: col_map["name"] = col
                elif "门店" in c_str: col_map["store"] = col
                elif "内容" in c_str or "文案" in c_str or "介绍" in c_str: col_map["content"] = col
                # If specific column names are known, add them here
            
            # Fallback if columns not found by name (try standard indices as backup)
            if "name" not in col_map and df.shape[1] > 1: col_map["name"] = df.columns[1]
            # --- Intelligent Column Mapping based on CONTENT ---
            # Instead of trusting headers, we check the content of the first few rows (non-empty)
            # to find which column contains the rich text (Skills, Monologue, etc.)
            
            content_col_name = None
            name_col_name = None
            store_col_name = None
            
            # Helper to find column by keyword in its values
            def find_col_by_value_keyword(dataframe, keywords, search_rows=10):
                for col in dataframe.columns:
                    # Check first N non-null values
                    sample_values = dataframe[col].dropna().head(search_rows).astype(str).tolist()
                    joined_sample = " ".join(sample_values)
                    if any(k in joined_sample for k in keywords):
                        return col
                return None

            # 1. Find Content Column (Look for "匠人独白" or "维修技师")
            content_col_name = find_col_by_value_keyword(df, ["匠人独白", "维修技师", "深耕", "服务至上"])
            
            # 2. Find Name Column (Look for "姓名" in header or header-like logic previously done)
            # If we already have col_map from previous step, use it, otherwise refined search
            if "name" in col_map: 
                name_col_name = col_map["name"]
            else:
                # Fallback: usually column 1
                if df.shape[1] > 1: name_col_name = df.columns[1]

            # 3. Find Store Column (Look for "门店" in header or "店" in values)
            if "store" in col_map:
                store_col_name = col_map["store"]
            else:
                 store_col_name = find_col_by_value_keyword(df, ["店", "服务点", "中心"])
                 if not store_col_name and df.shape[1] > 4: store_col_name = df.columns[4]

            # 4. Find Experience Column (Look for "匠龄" in header or values)
            # This is the Safety Net for people like Chen Xin who miss it in the text.
            exp_col_name = find_col_by_value_keyword(df, ["匠龄", "年", "从业"])
            # Fallback if specific column name exists from header detection
            if "匠龄" in df.columns: exp_col_name = "匠龄"

            self.log(f"锁定关键列 -> 姓名: [{name_col_name}], 门店: [{store_col_name}], 文案: [{content_col_name}], 匠龄(备用): [{exp_col_name}]")

            # Validate
            if not content_col_name:
                self.log("❌ 严重警告: 无法在任何列中找到包含'匠人独白'或'维修技师'的内容。")
                self.log("请检查Excel中是否包含完整的文案列。尝试使用默认列索引继续...")
                # Fallback to old behavior just in case
                if "content" in col_map: content_col_name = col_map["content"]
                elif df.shape[1] > 6: content_col_name = df.columns[6]

            # Handle Merged Cells for '门店'
            if store_col_name:
                df[store_col_name] = df[store_col_name].fillna(method='ffill')

            # --- Processing Loop ---
            
            new_data = []
            
            for idx, row in df.iterrows():
                # Get Basic Info
                name = str(row[name_col_name]).strip() if name_col_name and pd.notna(row[name_col_name]) else ""
                if not name: continue # Skip empty names
                
                store = str(row[store_col_name]).strip() if store_col_name and pd.notna(row[store_col_name]) else ""
                # store = store.replace("某某品牌后缀", "").strip() # 可在此处添加特定品牌后缀清洗逻辑

                # Get Rich Text
                full_text = str(row[content_col_name]).strip() if content_col_name and pd.notna(row[content_col_name]) else ""
                
                # Initialize fields
                # Default Experience from Excel Column (Safety Net)
                experience = ""
                if exp_col_name and pd.notna(row[exp_col_name]):
                    experience = str(row[exp_col_name]).strip()

                monologue = ""
                titles = []
                descriptions = []
                
                # --- Advanced Parsing Logic (State Machine Style) ---
                # To handle both single-line "*Title: Desc" and multi-line "*Title:\nDesc" formats.
                
                parts_list = [] # Will store (title, description) tuples
                current_title = ""
                current_desc_lines = []
                
                lines = full_text.split('\n')
                
                def flush_current_section():
                    nonlocal current_title, current_desc_lines, parts_list
                    if current_title or current_desc_lines:
                        d_text = " ".join(current_desc_lines).strip()
                        
                        # --- Feature Restoration: Ensure description ends with period ---
                        if d_text and not re.search(r'[。！!？?\.]$', d_text):
                            d_text += "。"
                        
                        if current_title or d_text:
                            parts_list.append((current_title, d_text))
                    current_title = ""
                    current_desc_lines = []

                for line in lines:
                    line = line.strip()
                    if not line: continue
                    
                    # 1. Monologue (High Priority)
                    if "匠人独白" in line:
                        flush_current_section() # Close previous section
                        parts = re.split(r"[:：]", line, maxsplit=1)
                        if len(parts) > 1:
                            monologue = self.clean_text(parts[1])
                            # --- Feature Restoration: Remove trailing punctuation ---
                            monologue = re.sub(r'[。！!.\s,，]+$', '', monologue)
                        continue
                        
                    # 2. Tech / Experience (High Priority)
                    if "维修技师" in line:
                        flush_current_section()
                        # Strategy 1: Look for explicit "匠龄" keyword
                        exp_match = re.search(r"匠龄[:：]\s*(\S+)", line)
                        
                        # Strategy 2: If keyword missing, look for standalone duration pattern
                        # Matches: "10+年", "近10年", "20年", "6+年"
                        if not exp_match:
                            exp_match = re.search(r'((?:近)?\d{1,2}\+?年)', line)
                        
                        if exp_match:
                            raw_exp = exp_match.group(1).strip()
                            # Extra cleanup: remove any accidental "匠龄" or colons if data was malformed
                            # e.g. "匠龄：匠龄：10年" -> "10年"
                            experience = re.sub(r'^(匠龄|[:：])+', '', raw_exp).strip()
                        
                        # IMPORTANT: REMOVED THE NAME OVERWRITE LOGIC HERE based on User Feedack
                        # We trust the Excel Name Column (A) more than the copy-pasted text content.
                        continue
                    
                    # 3. Titles (Lines starting with *)
                    # Also try to detect lines that look like titles even without * if they end with colon
                    is_title_line = False
                    if line.startswith("*") or line.startswith("●") or line.startswith("•"):
                        is_title_line = True
                    # Regex for "Text:" pattern which is likely a title
                    elif re.match(r'^.{2,15}[:：]\s*$', line): # "Short Text:"
                         is_title_line = True
                    
                    if is_title_line:
                        # This starts a new section
                        flush_current_section()
                        
                        # Try to split if it's "Title: Content" on one line
                        if "：" in line or ":" in line:
                            split_parts = re.split(r"[:：]", line, maxsplit=1)
                            t_raw = split_parts[0].strip().replace("*", "").replace("●", "").replace("•", "")
                            d_raw = split_parts[1].strip()
                            
                            current_title = self.clean_text(t_raw)
                            if d_raw:
                                current_desc_lines.append(self.clean_text(d_raw))
                        else:
                            # Just a title line without colon? specific case
                            current_title = self.clean_text(line.replace("*", ""))
                            
                    else:
                        # 4. Content Line
                        # If we have a current title, this is its description
                        if current_title:
                            current_desc_lines.append(self.clean_text(line))
                        else:
                            # Orphaned text? Only if it looks like description content
                            # Check if line contains a colon, might be a Title we missed?
                            if ("：" in line or ":" in line) and len(line) < 50:
                                # Treat as new title-desc pair
                                flush_current_section()
                                split_parts = re.split(r"[:：]", line, maxsplit=1)
                                t_raw = split_parts[0].strip()
                                d_raw = split_parts[1].strip()
                                current_title = self.clean_text(t_raw)
                                if d_raw:
                                    current_desc_lines.append(self.clean_text(d_raw))
                            else:
                                # Just append to previous if exists, or ignore?
                                # For robustness, if parts_list is not empty, append to last one?
                                # Or just ignore "2025入职" type junk lines if they appear early.
                                if current_title:
                                     current_desc_lines.append(self.clean_text(line))

                # End loop
                flush_current_section()
                
                # Fill titles/descriptions lists
                for t, d in parts_list:
                    if len(titles) < 3:
                        titles.append(t)
                        descriptions.append(d)

                # Fallback: If Monologue is STILL empty, use the specific fallback user rejected? 
                # NO, user said monologue EXISTS. If we fail here, it's better to leave empty 
                # than to guess incorrectly, but we should log it.
                if not monologue:
                    # Try one last regex on full string in case it wasn't on a single line
                    mono_match = re.search(r"匠人独白[:：]\s*(.*)", full_text, re.DOTALL)
                    if mono_match:
                        monologue = self.clean_text(mono_match.group(1))

                new_data.append({
                    "姓名": name,
                    "门店": store,
                    "匠龄": experience,
                    "匠人独白": monologue,
                    "标题1": titles[0], "描述1": descriptions[0],
                    "标题2": titles[1], "描述2": descriptions[1],
                    "标题3": titles[2], "描述3": descriptions[2]
                })

            # Create DataFrame
            df_cleaned = pd.DataFrame(new_data)
            
            # Save Cleaned Excel
            output_path = os.path.splitext(input_path)[0] + "_清洗版.xlsx"
            df_cleaned.to_excel(output_path, index=False)
            self.log(f"清洗完成！已保存为: {os.path.basename(output_path)}")
            
            # --- Generate Verification Report (Checklist) ---
            check_file_path = os.path.splitext(input_path)[0] + "_数据核对单.txt"
            with open(check_file_path, "w", encoding="utf-8") as f:
                f.write("=== 数据核对报告 ===\n")
                f.write("请务必检查以下信息是否与原始Excel对应。\n\n")
                
                for idx, row in df_cleaned.iterrows():
                    name = row.get("姓名", "N/A")
                    store = row.get("门店", "N/A")
                    monologue = row.get("匠人独白", "N/A")
                    t1 = row.get("标题1", "")
                    d1 = row.get("描述1", "")[:15] + "..." if row.get("描述1") else ""
                    
                    f.write(f"[{idx+1}] {name} @ {store}\n")
                    f.write(f"     独白: {monologue}\n")
                    f.write(f"     T1: {t1} | D1: {d1}\n")
                    f.write("-" * 50 + "\n")
            
            self.log(f"已生成核对报告: {os.path.basename(check_file_path)}")
            # Delay opening slightly
            self.root.after(500, lambda: os.startfile(check_file_path))
            
            messagebox.showinfo("完成", f"数据清洗完成！\n\n请查看打开的【数据核对报告】\n")

        except Exception as e:
            self.log(f"清洗数据失败: {e}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("错误", f"清洗失败: {e}")
        finally:
            self.btn_clean.config(state='normal')

    def clean_text(self, text):
        if not text: return ""
        text = str(text).strip()
        
        # 1. Normalize all whitespace
        text = re.sub(r'\s+', ' ', text)
        
        # 2. Fix specific symbols
        text = text.replace(' +', '+').replace('+ ', '+')
        
        # 3. Handle Duplicate Punctuation (Cleanup before space removal)
        # Replace multiple Chinese commas/periods with single ones
        # Also handles mixed like ",，" or "。."
        text = re.sub(r'[，,]{2,}', '，', text)
        text = re.sub(r'[。.]{2,}', '。', text)
        text = re.sub(r'[！!]{2,}', '！', text)
        text = re.sub(r'[？?]{2,}', '？', text)
        
        # 4. Remove spaces strictly for Chinese context
        # Strategy: If a space is adjacent to ANY Chinese character or Full-width Punctuation, remove it.
        # This covers cases like: "字 “" -> "字“", "字 ，" -> "字，"
        
        # Pattern 1: (Chinese/FullWidth) <Space> (Any Non-Space)
        # We assume if the left side is Chinese/FullWidth, we don't need a space after it.
        # Range includes Chinese \u4e00-\u9fa5 and Fullwidth Punctuation \uff00-\uffef
        ch_punct_range = r'[\u4e00-\u9fa5\uff00-\uffef]'
        
        # Remove space AFTER Chinese/Punctuation
        text = re.sub(f'(?<={ch_punct_range})\\s+(?=[\\S])', '', text)
        
        # Remove space BEFORE Chinese/Punctuation
        text = re.sub(f'(?<=[\\S])\\s+(?={ch_punct_range})', '', text)
        
        # 5. Case specific for user: "50+" space removal (Number to Number/Symbol)
        text = re.sub(r'(?<=[\d])\s+(?=[\d+])', '', text)
        
        return text

    # --- Step 2 Handlers ---
    def handle_drop_psd(self, event):
        file_path = self.parse_path(event.data)
        self.psd_input_var.set(file_path)
        self.update_data_status(file_path)
        self.log(f"[PSD] 已选择文件: {file_path}")

    def browse_file_psd(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.psd_input_var.set(filename)
            self.update_data_status(filename)
            self.log(f"[PSD] 已选择文件: {filename}")

    def start_psd_gen(self):
        excel_path = self.psd_input_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showwarning("提示", "请提供有效的清洗版Excel文件！")
            return
        
        # Verify it's an Excel file
        if not excel_path.lower().endswith(('.xlsx', '.xls')):
            messagebox.showerror("错误", "文件格式错误！\n\n请选择【Excel 文件】(.xlsx)，\n而不是 PSD 模板文件。")
            return
        
        template_path = os.path.join(self.base_dir, 'model', '维修师-模板.psd')
        if not os.path.exists(template_path):
            messagebox.showerror("错误", f"未找到模板文件: {template_path}")
            return
        
        output_dir = os.path.join(self.base_dir, 'output_psds')

        self.btn_gen_psd.config(state='disabled')
        self.log("正在启动 Photoshop 生成任务 (请勿关闭 Photoshop)...")
        
        # Threading for PSD
        thread = threading.Thread(target=self.psd_tool.process_batch, args=(excel_path, template_path, output_dir))
        thread.start()
        
        # Monitor thread
        self.monitor_psd_thread(thread)

    def monitor_psd_thread(self, thread):
        if thread.is_alive():
            self.root.after(1000, lambda: self.monitor_psd_thread(thread))
        else:
            self.btn_gen_psd.config(state='normal')


    def parse_path(self, data):
        if data.startswith('{') and data.endswith('}'):
            return data[1:-1]
        return data

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()
