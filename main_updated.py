import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os

# ราคากลางของเก่า
SCRAP_PRICES = {
    "กระดาษลัง": 3.50,
    "กระดาษขาว-ดำ": 6.00,
    "หนังสือพิมพ์": 7.50,
    "ขวดพลาสติกใส (PET)": 13.50,
    "พลาสติกขาวขุ่น/ขวดน้ำ": 11.00,
    "เหล็กหนา": 9.50,
    "เหล็กบาง": 8.00,
    "กระป๋องอลูมิเนียม": 60.00,
    "ทองแดง (เบอร์ 1)": 295.00,
    "สแตนเลส (แท้)": 32.50,
    "ขวดเบียร์ (ช้าง, ลีโอ)": 13.00,
    "เศษแก้วขาว": 1.50
}

class ScrapShopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมคำนวณรับซื้อของเก่า V5.0")
        self.root.state("zoomed")  # เต็มหน้าจอ

        # --- Style ---
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TLabel", font=("TH Sarabun New", 18))
        style.configure("TButton", font=("TH Sarabun New", 18), padding=10)
        style.configure("TEntry", font=("TH Sarabun New", 18))

        # --- Main Frame ---
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Item ---
        ttk.Label(main_frame, text="เลือกสินค้า:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.item_var = tk.StringVar()
        self.item_combobox = ttk.Combobox(main_frame, textvariable=self.item_var, values=list(SCRAP_PRICES.keys()), width=30, font=("TH Sarabun New", 18))
        self.item_combobox.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.item_combobox.bind("<<ComboboxSelected>>", self.update_price)

        # --- Price ---
        ttk.Label(main_frame, text="ราคา/กก. (บาท):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.price_var = tk.DoubleVar()
        self.price_entry = ttk.Entry(main_frame, textvariable=self.price_var, width=30)
        self.price_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # --- Weight ---
        ttk.Label(main_frame, text="น้ำหนัก (กก.):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.weight_var = tk.DoubleVar()
        self.weight_entry = ttk.Entry(main_frame, textvariable=self.weight_var, width=30)
        self.weight_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

        # --- Buttons ---
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        self.calc_button = ttk.Button(button_frame, text="🧮 คำนวณ", command=self.calculate)
        self.calc_button.grid(row=0, column=0, padx=20)

        self.save_button = ttk.Button(button_frame, text="💾 บันทึก Excel", command=self.save_excel, state="disabled")
        self.save_button.grid(row=0, column=1, padx=20)

        self.print_button = ttk.Button(button_frame, text="🖨️ พิมพ์ใบเสร็จ (PDF)", command=self.print_receipt, state="disabled")
        self.print_button.grid(row=0, column=2, padx=20)

        # --- Result ---
        self.result_label = ttk.Label(main_frame, text="กรอกข้อมูลแล้วกดคำนวณ", font=("TH Sarabun New", 22, "bold"), foreground="darkgreen")
        self.result_label.grid(row=4, column=0, columnspan=2, pady=10)

        # --- Tables Frame ---
        tables_frame = ttk.Frame(main_frame, padding="10")
        tables_frame.grid(row=5, column=0, columnspan=2, sticky="nsew")

        # ตารางประวัติการคำนวณ
        ttk.Label(tables_frame, text="📊 ประวัติการคำนวณ", font=("TH Sarabun New", 20, "bold")).pack(anchor="w")
        self.calc_tree = ttk.Treeview(tables_frame, columns=("date", "item", "price", "weight", "total"), show="headings", height=6)
        for col, text in zip(("date", "item", "price", "weight", "total"), ("วันที่", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม")):
            self.calc_tree.heading(col, text=text)
            self.calc_tree.column(col, width=150, anchor="center")
        self.calc_tree.pack(fill=tk.X, pady=5)

        # ตารางประวัติการออกใบเสร็จ
        ttk.Label(tables_frame, text="🧾 ประวัติการออกใบเสร็จ", font=("TH Sarabun New", 20, "bold")).pack(anchor="w", pady=(20,0))
        self.receipt_tree = ttk.Treeview(tables_frame, columns=("date", "item", "price", "weight", "total", "file"), show="headings", height=6)
        for col, text in zip(("date", "item", "price", "weight", "total", "file"), ("วันที่", "สินค้า", "ราคา/กก.", "น้ำหนัก", "รวม", "ไฟล์ PDF")):
            self.receipt_tree.heading(col, text=text)
            self.receipt_tree.column(col, width=150, anchor="center")
        self.receipt_tree.pack(fill=tk.X, pady=5)

        # Excel file
        self.excel_file = "scrap_records.xlsx"
        if not os.path.exists(self.excel_file):
            self.create_excel_file()

        # ตัวแปรเก็บผลลัพธ์
        self.current_data = None

    def update_price(self, event=None):
        selected_item = self.item_var.get()
        if selected_item in SCRAP_PRICES:
            self.price_var.set(SCRAP_PRICES[selected_item])

    def create_excel_file(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Records"
        ws.append(["วันที่", "สินค้า", "ราคา/กก.", "น้ำหนัก (กก.)", "รวม (บาท)"])
        wb.save(self.excel_file)

    def calculate(self):
        try:
            item = self.item_var.get()
            price_per_kg = self.price_var.get()
            weight = self.weight_var.get()

            if not item or price_per_kg <= 0 or weight <= 0:
                messagebox.showerror("ข้อมูลไม่ครบถ้วน", "กรุณาเลือกสินค้าและกรอกข้อมูลราคาและน้ำหนักให้ถูกต้อง")
                return

            total = price_per_kg * weight
            self.current_data = (
                datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                item,
                price_per_kg,
                weight,
                total
            )

            self.result_label.config(text=f"รวมทั้งสิ้น: {total:,.2f} บาท")
            self.save_button.config(state="normal")
            self.print_button.config(state="normal")

            # เพิ่มในตารางคำนวณ
            self.calc_tree.insert("", "end", values=self.current_data)

        except ValueError:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกข้อมูลตัวเลขให้ถูกต้อง")

    def save_excel(self):
        if not self.current_data:
            messagebox.showerror("ไม่มีข้อมูล", "กรุณาคำนวณก่อนบันทึก")
            return

        try:
            if os.path.exists(self.excel_file):
                wb = load_workbook(self.excel_file)
                ws = wb.active
            else:
                self.create_excel_file()
                wb = load_workbook(self.excel_file)
                ws = wb.active

            ws.append(self.current_data)
            wb.save(self.excel_file)

            messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลลง {self.excel_file} เรียบร้อยแล้ว")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึก Excel ได้: {e}")

    def print_receipt(self):
        if not self.current_data:
            messagebox.showerror("ไม่มีข้อมูล", "กรุณาคำนวณก่อนพิมพ์")
            return

        try:
            # --- Register Thai Font ---
            font_path = "TH Sarabun New Bold.ttf"
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('THSarabun', font_path))
                font_name = "THSarabun"
            else:
                font_name = "Helvetica"

            date, item, price_per_kg, weight, total = self.current_data
            filename = f"receipt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            c = canvas.Canvas(filename, pagesize=letter)

            c.setFont(font_name, 18)
            c.drawString(100, 800, "ใบเสร็จรับเงิน (Receipt)")

            c.setFont(font_name, 16)
            c.drawString(100, 770, f"สินค้า: {item}")
            c.drawString(100, 750, f"ราคาต่อหน่วย: {price_per_kg:,.2f} บาท/กก.")
            c.drawString(100, 730, f"น้ำหนัก: {weight:,.2f} กก.")
            c.line(100, 720, 500, 720)

            c.setFont(font_name, 18)
            c.drawString(100, 700, f"รวมทั้งสิ้น: {total:,.2f} บาท")
            c.line(100, 690, 500, 690)

            c.setFont(font_name, 14)
            c.drawString(100, 660, f"วันที่: {date}")

            c.save()
            os.startfile(filename)

            # เพิ่มในตารางประวัติใบเสร็จ
            self.receipt_tree.insert("", "end", values=(date, item, price_per_kg, weight, total, filename))

        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถสร้างใบเสร็จ PDF ได้: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScrapShopApp(root)
    root.mainloop()
