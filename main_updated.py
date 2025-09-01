import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os
import platform
import json

# Create folders if they don't exist
for folder in ['data', 'receipts_in', 'receipts_out']:
    if not os.path.exists(folder):
        os.makedirs(folder)

class ScrapShopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมคำนวณรับซื้อและจำหน่ายของเก่า V6.1")
        self.root.state("zoomed")

        # --- Initialize Thai font ---
        self.thai_font_name = self.register_thai_font()

        # --- Style ---
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TLabel", font=("TH Sarabun New", 18))
        style.configure("TButton", font=("TH Sarabun New", 18), padding=10)
        style.configure("TEntry", font=("TH Sarabun New", 18))

        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        # --- Notebook ---
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Load prices
        self.load_prices()

        # Setup Tabs
        self.incoming_tab = ttk.Frame(self.notebook, padding="10")
        self.outgoing_tab = ttk.Frame(self.notebook, padding="10")
        self.history_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.incoming_tab, text="รับเข้า")
        self.notebook.add(self.outgoing_tab, text="จำหน่ายออก")
        self.notebook.add(self.history_tab, text="ประวัติและคงคลัง")

        self._setup_transaction_tab(
            self.incoming_tab, "in", self.BUY_PRICES, self.update_buy_price,
            "ชื่อผู้ขาย:", "ชื่อผู้รับ:"
        )
        self._setup_transaction_tab(
            self.outgoing_tab, "out", self.SELL_PRICES, self.update_sell_price,
            "ชื่อผู้จ่าย:", "ชื่อผู้รับ (เช่น โรงงาน):"
        )
        self.setup_history_tab()

        # Excel and JSON files
        self.incoming_excel = os.path.join('data', 'incoming_scrap_records.xlsx')
        self.outgoing_excel = os.path.join('data', 'outgoing_scrap_records.xlsx')
        self.receipt_history_file = os.path.join('data', 'receipt_history.json')

        if not os.path.exists(self.incoming_excel):
            self.create_excel_file(self.incoming_excel)
        if not os.path.exists(self.outgoing_excel):
            self.create_excel_file(self.outgoing_excel)

        # Load histories
        self.load_excel_history(self.incoming_excel, self.incoming_calc_tree)
        self.load_excel_history(self.outgoing_excel, self.outgoing_calc_tree)
        self.load_receipt_history()

        # Compute inventory
        self.compute_inventory()

        # Current data
        self.current_in_data = None
        self.current_out_data = None

    def register_thai_font(self):
        """ลงทะเบียนฟอนต์ภาษาไทยสำหรับ PDF"""
        try:
            # ลำดับการค้นหาไฟล์ฟอนต์
            font_paths = [
                # ในโฟลเดอร์เดียวกับโปรแกรม
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "TH Sarabun New Bold.ttf"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "THSarabunNew Bold.ttf"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "TH Sarabun New Bold.ttf"),
                
                # ใน Windows
                "C:/Windows/Fonts/THSarabunNew Bold.ttf",
                "C:/Windows/Fonts/TH Sarabun New Bold.ttf",
                
                # ใน Linux
                "/usr/share/fonts/truetype/thai/TH Sarabun New Bold.ttf",
                "/usr/local/share/fonts/TH Sarabun New Bold.ttf",
                
                # ใน macOS
                "/System/Library/Fonts/TH Sarabun New Bold.ttf",
                "/Library/Fonts/TH Sarabun New Bold.ttf"
            ]
            
            # ลองหาฟอนต์ที่มีอยู่
            for font_path in font_paths:
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont('THSarabun', font_path))
                        print(f"✅ ลงทะเบียนฟอนต์สำเร็จ: {font_path}")
                        return 'THSarabun'
                    except Exception as e:
                        print(f"❌ ไม่สามารถลงทะเบียนฟอนต์ {font_path}: {e}")
                        continue
            
            # หากไม่พบฟอนต์ไทย ให้แสดงคำเตือน
            messagebox.showwarning(
                "ไม่พบฟอนต์ภาษาไทย", 
                "ไม่พบไฟล์ฟอนต์ TH Sarabun New Bold.ttf\n\n"
                "กรุณาดาวน์โหลดฟอนต์จาก https://fonts.google.com/specimen/Sarabun\n"
                "แล้ววางไฟล์ในโฟลเดอร์เดียวกับโปรแกรม\n\n"
                "ตอนนี้จะใช้ฟอนต์ Helvetica แทน (อาจแสดงภาษาไทยไม่ถูกต้อง)"
            )
            return 'Helvetica'
            
        except Exception as e:
            print(f"❌ เกิดข้อผิดพลาดในการลงทะเบียนฟอนต์: {e}")
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถลงทะเบียนฟอนต์ได้: {e}")
            return 'Helvetica'

    def load_prices(self):
        try:
            with open('prices.json', 'r', encoding='utf-8') as f:
                prices = json.load(f)
                self.BUY_PRICES = prices['BUY_PRICES']
                self.SELL_PRICES = prices['SELL_PRICES']
        except (FileNotFoundError, json.JSONDecodeError):
            self.BUY_PRICES = {
                "กระดาษลัง": 3.50, "กระดาษขาว-ดำ": 6.00, "หนังสือพิมพ์": 7.50,
                "ขวดพลาสติกใส (PET)": 13.50, "พลาสติกขาวขุ่น/ขวดน้ำ": 11.00,
                "เหล็กหนา": 9.50, "เหล็กบาง": 8.00, "กระป๋องอลูมิเนียม": 60.00,
                "ทองแดง (เบอร์ 1)": 295.00, "สแตนเลส (แท้)": 32.50,
                "ขวดเบียร์ (ช้าง, ลีโอ)": 13.00, "เศษแก้วขาว": 1.50
            }
            self.SELL_PRICES = {k: v * 1.1 for k, v in self.BUY_PRICES.items()}
            self.save_prices()

    def save_prices(self):
        prices = {'BUY_PRICES': self.BUY_PRICES, 'SELL_PRICES': self.SELL_PRICES}
        with open('prices.json', 'w', encoding='utf-8') as f:
            json.dump(prices, f, ensure_ascii=False, indent=4)

    def _setup_transaction_tab(self, tab, mode, prices, update_price_cmd, label1_text, label2_text):
        vcmd = (self.root.register(self.validate_numeric), '%P')

        ttk.Label(tab, text=label1_text).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        name1_var = tk.StringVar()
        ttk.Entry(tab, textvariable=name1_var, width=30).grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ttk.Label(tab, text=label2_text).grid(row=1, column=0, padx=10, pady=10, sticky="w")
        name2_var = tk.StringVar()
        ttk.Entry(tab, textvariable=name2_var, width=30).grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        ttk.Label(tab, text="เลือกสินค้า:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        item_var = tk.StringVar()
        item_combobox = ttk.Combobox(tab, textvariable=item_var, values=list(prices.keys()), width=30, font=("TH Sarabun New", 18))
        item_combobox.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        item_combobox.bind("<<ComboboxSelected>>", update_price_cmd)

        ttk.Label(tab, text="ราคา/กก. (บาท):").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        price_var = tk.DoubleVar()
        ttk.Entry(tab, textvariable=price_var, width=30, validate="key", validatecommand=vcmd).grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        ttk.Label(tab, text="น้ำหนัก (กก.):").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        weight_var = tk.DoubleVar()
        ttk.Entry(tab, textvariable=weight_var, width=30, validate="key", validatecommand=vcmd).grid(row=4, column=1, padx=10, pady=10, sticky="ew")

        button_frame = ttk.Frame(tab, padding="10")
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="🧮 คำนวณ", command=lambda: self._calculate(mode)).grid(row=0, column=0, padx=20)
        save_print_button = ttk.Button(button_frame, text="💾🖨️ บันทึกและพิมพ์", command=lambda: self._save_print(mode), state="disabled")
        save_print_button.grid(row=0, column=1, padx=20)

        result_label = ttk.Label(tab, text="กรอกข้อมูลแล้วกดคำนวณ", font=("TH Sarabun New", 22, "bold"), foreground="darkgreen")
        result_label.grid(row=6, column=0, columnspan=2, pady=10)

        if mode == 'in':
            self.seller_var, self.buyer_var = name1_var, name2_var
            self.item_in_var, self.price_in_var, self.weight_in_var = item_var, price_var, weight_var
            self.save_print_in_button, self.result_in_label = save_print_button, result_label
        else:
            self.payer_var, self.recipient_var = name1_var, name2_var
            self.item_out_var, self.price_out_var, self.weight_out_var = item_var, price_var, weight_var
            self.save_print_out_button, self.result_out_label = save_print_button, result_label

        if prices:
            item_var.set(list(prices.keys())[0])
            update_price_cmd()

    def setup_history_tab(self):
        tables_frame = ttk.Frame(self.history_tab, padding="10")
        tables_frame.pack(fill=tk.BOTH, expand=True)
        tables_frame.grid_rowconfigure((1, 3, 5, 7, 9), weight=1)
        tables_frame.grid_columnconfigure(0, weight=1)

        # Incoming Calc History
        ttk.Label(tables_frame, text="📊 สินค้ารับเข้า", font=("TH Sarabun New", 20, "bold")).grid(row=0, column=0, sticky="w")
        self.incoming_calc_tree = ttk.Treeview(tables_frame, columns=("date", "seller", "buyer", "item", "price", "weight", "total"), show="headings", height=6)
        for col, text in zip(("date", "seller", "buyer", "item", "price", "weight", "total"), ("วันที่", "ผู้ขาย", "ผู้รับ", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม")):
            self.incoming_calc_tree.heading(col, text=text)
            self.incoming_calc_tree.column(col, width=150, anchor="center")
        self.incoming_calc_tree.grid(row=1, column=0, sticky="nsew", pady=5)

        # Outgoing Calc History
        ttk.Label(tables_frame, text="📊 สินค้าจำหน่ายออก", font=("TH Sarabun New", 20, "bold")).grid(row=2, column=0, sticky="w", pady=(20, 0))
        self.outgoing_calc_tree = ttk.Treeview(tables_frame, columns=("date", "payer", "recipient", "item", "price", "weight", "total"), show="headings", height=6)
        for col, text in zip(("date", "payer", "recipient", "item", "price", "weight", "total"), ("วันที่", "ผู้จ่าย", "ผู้รับ", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม")):
            self.outgoing_calc_tree.heading(col, text=text)
            self.outgoing_calc_tree.column(col, width=150, anchor="center")
        self.outgoing_calc_tree.grid(row=3, column=0, sticky="nsew", pady=5)

        # Receipt In History
        ttk.Label(tables_frame, text="🧾 ประวัติใบเสร็จรับเข้า", font=("TH Sarabun New", 20, "bold")).grid(row=4, column=0, sticky="w", pady=(20, 0))
        self.receipt_in_tree = ttk.Treeview(tables_frame, columns=("date", "seller", "buyer", "item", "price", "weight", "total", "file"), show="headings", height=6)
        for col, text in zip(("date", "seller", "buyer", "item", "price", "weight", "total", "file"), ("วันที่", "ผู้ขาย", "ผู้รับ", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม", "ไฟล์ PDF")):
            self.receipt_in_tree.heading(col, text=text)
            self.receipt_in_tree.column(col, width=150, anchor="center")
        self.receipt_in_tree.grid(row=5, column=0, sticky="nsew", pady=5)

        # Receipt Out History
        ttk.Label(tables_frame, text="🧾 ประวัติใบเสร็จจำหน่ายออก", font=("TH Sarabun New", 20, "bold")).grid(row=6, column=0, sticky="w", pady=(20, 0))
        self.receipt_out_tree = ttk.Treeview(tables_frame, columns=("date", "payer", "recipient", "item", "price", "weight", "total", "file"), show="headings", height=6)
        for col, text in zip(("date", "payer", "recipient", "item", "price", "weight", "total", "file"), ("วันที่", "ผู้จ่าย", "ผู้รับ", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม", "ไฟล์ PDF")):
            self.receipt_out_tree.heading(col, text=text)
            self.receipt_out_tree.column(col, width=150, anchor="center")
        self.receipt_out_tree.grid(row=7, column=0, sticky="nsew", pady=5)

        # Inventory
        ttk.Label(tables_frame, text="📦 สินค้าคงคลัง", font=("TH Sarabun New", 20, "bold")).grid(row=8, column=0, sticky="w", pady=(20, 0))
        self.inventory_tree = ttk.Treeview(tables_frame, columns=("item", "total_in", "total_out", "stock"), show="headings", height=6)
        for col, text in zip(("item", "total_in", "total_out", "stock"), ("สินค้า", "รวมรับเข้า", "รวมจำหน่ายออก", "คงคลัง")):
            self.inventory_tree.heading(col, text=text)
            self.inventory_tree.column(col, width=150, anchor="center")
        self.inventory_tree.grid(row=9, column=0, sticky="nsew", pady=5)

    def validate_numeric(self, new_value):
        if new_value == "":
            return True
        try:
            float(new_value)
            return True
        except ValueError:
            return False

    def update_buy_price(self, event=None):
        selected_item = self.item_in_var.get()
        if selected_item in self.BUY_PRICES:
            self.price_in_var.set(self.BUY_PRICES[selected_item])

    def update_sell_price(self, event=None):
        selected_item = self.item_out_var.get()
        if selected_item in self.SELL_PRICES:
            self.price_out_var.set(self.SELL_PRICES[selected_item])

    def create_excel_file(self, excel_file):
        wb = Workbook()
        ws = wb.active
        ws.title = "Records"
        ws.append(["วันที่", "ชื่อผู้ขาย/ผู้จ่าย", "ชื่อผู้รับ", "สินค้า", "ราคา/กก.", "น้ำหนัก (กก.)", "รวม (บาท)"])
        wb.save(excel_file)

    def load_excel_history(self, excel_file, tree):
        if os.path.exists(excel_file):
            try:
                wb = load_workbook(excel_file, read_only=True)
                ws = wb.active
                for row in ws.iter_rows(min_row=2, values_only=True):
                    tree.insert("", "end", values=row)
            except Exception as e:
                messagebox.showwarning("ข้อผิดพลาดในการโหลด", f"ไม่สามารถโหลดประวัติจาก Excel: {e}")

    def load_receipt_history(self):
        try:
            with open(self.receipt_history_file, 'r', encoding='utf-8') as f:
                history = json.load(f)
            for record in history.get('in', []):
                self.receipt_in_tree.insert("", "end", values=tuple(record))
            for record in history.get('out', []):
                self.receipt_out_tree.insert("", "end", values=tuple(record))
        except (FileNotFoundError, json.JSONDecodeError):
            pass  # No history file yet, start with empty history

    def save_receipt_history(self, mode, data, filename):
        history = {'in': [], 'out': []}
        try:
            with open(self.receipt_history_file, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            pass

        record = list(data) + [filename]
        history[mode].append(record)
        try:
            with open(self.receipt_history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกประวัติใบเสร็จ: {e}")

    def compute_inventory(self):
        stock = {}
        if os.path.exists(self.incoming_excel):
            wb_in = load_workbook(self.incoming_excel, read_only=True)
            ws_in = wb_in.active
            for row in ws_in.iter_rows(min_row=2, values_only=True):
                if row[0] is None:  # Skip empty rows
                    continue
                _, _, _, item, _, weight, _ = row
                if item and weight:  # Make sure both exist
                    if item not in stock:
                        stock[item] = {'in': 0, 'out': 0}
                    stock[item]['in'] += float(weight)

        if os.path.exists(self.outgoing_excel):
            wb_out = load_workbook(self.outgoing_excel, read_only=True)
            ws_out = wb_out.active
            for row in ws_out.iter_rows(min_row=2, values_only=True):
                if row[0] is None:  # Skip empty rows
                    continue
                _, _, _, item, _, weight, _ = row
                if item and weight:  # Make sure both exist
                    if item not in stock:
                        stock[item] = {'in': 0, 'out': 0}
                    stock[item]['out'] += float(weight)

        self.inventory_tree.delete(*self.inventory_tree.get_children())
        for item, data in sorted(stock.items()):
            self.inventory_tree.insert("", "end", values=(item, f"{data['in']:.2f}", f"{data['out']:.2f}", f"{data['in'] - data['out']:.2f}"))

    def _calculate(self, mode):
        try:
            name1, name2, item, price_var, weight_var, result_label, save_print_button, tree = (
                (self.seller_var, self.buyer_var, self.item_in_var, self.price_in_var, self.weight_in_var,
                 self.result_in_label, self.save_print_in_button, self.incoming_calc_tree)
                if mode == 'in' else
                (self.payer_var, self.recipient_var, self.item_out_var, self.price_out_var, self.weight_out_var,
                 self.result_out_label, self.save_print_out_button, self.outgoing_calc_tree)
            )

            name1, name2, item, price_per_kg, weight = (
                name1.get().strip(), name2.get().strip(), item.get(), price_var.get(), weight_var.get()
            )

            if not name1 or not name2 or not item or price_per_kg <= 0 or weight <= 0:
                messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณากรอกข้อมูลให้ครบถ้วนและถูกต้อง")
                return

            total = price_per_kg * weight
            current_data = (
                datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                name1, name2, item, price_per_kg, weight, total
            )

            result_label.config(text=f"รวมทั้งสิ้น: {total:,.2f} บาท")
            save_print_button.config(state="normal")
            tree.insert("", "end", values=current_data)

            if mode == 'in':
                self.current_in_data = current_data
            else:
                self.current_out_data = current_data

        except ValueError:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกข้อมูลตัวเลขให้ถูกต้อง")

    def _save_print(self, mode):
        data, excel_file, tree = (
            (self.current_in_data, self.incoming_excel, self.receipt_in_tree)
            if mode == 'in' else
            (self.current_out_data, self.outgoing_excel, self.receipt_out_tree)
        )

        if not data:
            messagebox.showerror("ไม่มีข้อมูล", "กรุณาคำนวณก่อนบันทึก")
            return

        self.save_excel(data, excel_file)
        filename = self.print_receipt(data, mode)
        if filename:
            self.save_receipt_history(mode, data, filename)
            tree.insert("", "end", values=(*data, filename))
        self.compute_inventory()

    def save_excel(self, data, excel_file):
        try:
            if os.path.exists(excel_file):
                wb = load_workbook(excel_file)
                ws = wb.active
            else:
                self.create_excel_file(excel_file)
                wb = load_workbook(excel_file)
                ws = wb.active

            ws.append(data)
            wb.save(excel_file)
            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลลง {excel_file} เรียบร้อยแล้ว")
        except PermissionError:
            messagebox.showerror("สิทธิ์การเข้าถึงถูกปฏิเสธ", "ไม่สามารถบันทึกไฟล์ Excel ได้ กรุณาปิดไฟล์หากเปิดอยู่แล้วลองอีกครั้ง")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึก Excel ได้: {e}")

    def print_receipt(self, data, mode):
        """สร้างใบเสร็จ PDF พร้อมฟอนต์ภาษาไทย"""
        try:
            date, name1, name2, item, price_per_kg, weight, total = data
            dt_str = datetime.now().strftime('%Y%m%d_%H%M%S')
            folder = 'receipts_in' if mode == 'in' else 'receipts_out'
            filename = os.path.join(folder, f"receipt_{dt_str}.pdf")
            
            # สร้าง PDF
            c = canvas.Canvas(filename, pagesize=letter)
            width, height = letter

            # ใช้ฟอนต์ที่ลงทะเบียนไว้แล้ว
            try:
                c.setFont(self.thai_font_name, 20)
            except:
                c.setFont('Helvetica', 20)
                print("⚠️ ไม่สามารถใช้ฟอนต์ไทยได้ ใช้ Helvetica แทน")

            # หัวเรื่อง
            title = "ใบเสร็จรับซื้อของเก่า" if mode == 'in' else "ใบเสร็จจำหน่ายของเก่า"
            title_width = c.stringWidth(title, self.thai_font_name, 20)
            c.drawString((width - title_width) / 2, height - 100, title)

            # ข้อมูลหลัก
            y_position = height - 150
            line_height = 25

            try:
                c.setFont(self.thai_font_name, 16)
            except:
                c.setFont('Helvetica', 16)

            # ข้อมูลรายการ
            label1 = "ผู้ขาย:" if mode == 'in' else "ผู้จ่าย:"
            lines = [
                f"{label1} {name1}",
                f"ผู้รับ: {name2}",
                f"สินค้า: {item}",
                f"ราคาต่อหน่วย: {price_per_kg:,.2f} บาท/กก.",
                f"น้ำหนัก: {weight:,.2f} กก.",
            ]

            for line in lines:
                c.drawString(100, y_position, line)
                y_position -= line_height

            # เส้นคั่น
            c.line(100, y_position - 10, width - 100, y_position - 10)
            y_position -= 30

            # ยอดรวม
            try:
                c.setFont(self.thai_font_name, 18)
            except:
                c.setFont('Helvetica-Bold', 18)

            total_text = f"รวมทั้งสิ้น: {total:,.2f} บาท"
            c.drawString(100, y_position, total_text)

            # เส้นคั่นล่าง
            c.line(100, y_position - 20, width - 100, y_position - 20)
            y_position -= 40

            # วันที่และเวลา
            try:
                c.setFont(self.thai_font_name, 14)
            except:
                c.setFont('Helvetica', 14)

            c.drawString(100, y_position, f"วันที่: {date}")

            # ลายเซ็น (ถ้าต้องการ)
            y_position -= 80
            c.drawString(100, y_position, "ลายเซ็นผู้รับ: _____________________")
            c.drawString(350, y_position, "ลายเซ็นผู้จ่าย: _____________________")

            # บันทึกไฟล์
            c.save()
            
            # เปิดไฟล์
            self.open_file(filename)
            messagebox.showinfo("สำเร็จ", f"สร้างใบเสร็จ PDF เรียบร้อยแล้ว\nไฟล์: {filename}")
            return filename

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถสร้างใบเสร็จ PDF ได้: {e}")
            print(f"❌ PDF Error Details: {e}")
            return None

    def open_file(self, filename):
        """เปิดไฟล์ด้วยโปรแกรมเริ่มต้นของระบบ"""
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(filename)
            elif system == "Darwin":  # macOS
                os.system(f"open '{filename}'")
            elif system == "Linux":
                os.system(f"xdg-open '{filename}'")
            else:
                messagebox.showinfo("เปิดไฟล์", f"ไฟล์ {filename} ถูกสร้างแล้ว แต่ไม่สามารถเปิดอัตโนมัติได้บนระบบนี้ กรุณาเปิดด้วยตัวเอง")
        except Exception as e:
            print(f"❌ ไม่สามารถเปิดไฟล์ได้: {e}")
            messagebox.showinfo("เปิดไฟล์", f"ไฟล์ {filename} ถูกสร้างแล้ว กรุณาเปิดด้วยตัวเอง")

if __name__ == "__main__":
    root = tk.Tk()
    app = ScrapShopApp(root)
    root.mainloop()