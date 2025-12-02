import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pyodbc
import xlsxwriter  # <--- Thư viện bạn yêu cầu






# ====== 1. CẤU HÌNH KẾT NỐI ======
def connect_db():
    try:
        conn = pyodbc.connect(
            "DRIVER={SQL Server};"
            "SERVER=LAPTOP-NI0S7IB6\\SQL2017EXPRESS;"  # <--- Kiểm tra tên Server
            "DATABASE=QLDIEMSV;"
            "Trusted_Connection=yes;"
        )
        return conn
    except Exception as e:
        messagebox.showerror("Lỗi Kết Nối", f"Lỗi: {e}")
        return None


# ====== HÀM CĂN GIỮA MÀN HÌNH ======
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')


# =========================================================================
# LOGIC QUẢN LÝ SINH VIÊN (TAB 1)
# =========================================================================
def load_sv(tree):
    for i in tree.get_children(): tree.delete(i)
    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            cur.execute("SELECT MaSV, HoTen, NgaySinh, GioiTinh, MaLop, MaKhoa FROM SINHVIEN")
            for row in cur.fetchall(): tree.insert("", tk.END, values=list(row))
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
        finally:
            conn.close()


def them_sv(tree, e_ma, e_ten, e_ngay, e_gt, e_lop, e_khoa):
    masv, hoten = e_ma.get(), e_ten.get()
    ngaysinh = str(e_ngay.get_date())
    makhoa = e_khoa.get()
    if not masv or not hoten: return messagebox.showwarning("Lỗi", "Thiếu Mã SV hoặc Tên")
    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            sql = "INSERT INTO SINHVIEN (MaSV, HoTen, NgaySinh, GioiTinh, MaLop, MaKhoa) VALUES (?, ?, ?, ?, ?, ?)"
            cur.execute(sql, (masv, hoten, ngaysinh, e_gt.get(), e_lop.get(), makhoa))
            conn.commit()
            messagebox.showinfo("Thành công", "Đã thêm SV mới")
            load_sv(tree)
        except Exception as e:
            messagebox.showerror("Lỗi SQL", str(e))
        finally:
            conn.close()


def xoa_sv(tree):
    sel = tree.selection()
    if not sel: return
    masv = tree.item(sel)["values"][0]
    if messagebox.askyesno("Xóa", f"Xóa SV {masv} sẽ mất hết điểm. Tiếp tục?"):
        conn = connect_db()
        if conn:
            try:
                cur = conn.cursor()
                cur.execute("DELETE FROM DIEM WHERE MaSV=?", (masv,))
                cur.execute("DELETE FROM SINHVIEN WHERE MaSV=?", (masv,))
                conn.commit()
                load_sv(tree)
                messagebox.showinfo("Xong", "Đã xóa!")
            except Exception as e:
                messagebox.showerror("Lỗi", str(e))
            finally:
                conn.close()


# =========================================================================
# LOGIC QUẢN LÝ ĐIỂM & XUẤT EXCEL (TAB 2)
# =========================================================================
def load_diem(tree):
    for i in tree.get_children(): tree.delete(i)
    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            sql = """
                SELECT SV.MaSV, SV.HoTen, MH.TenMon, D.DiemQuaTrinh, D.DiemGiuaKy, D.DiemCuoiKy, D.DiemTongKet
                FROM DIEM D
                JOIN SINHVIEN SV ON D.MaSV = SV.MaSV
                JOIN HOCPHAN HP ON D.MaHocPhan = HP.MaHocPhan
                JOIN MONHOC MH ON HP.MaMon = MH.MaMon
            """
            cur.execute(sql)
            for row in cur.fetchall():
                tong = row[6] if row[6] else 0
                ketqua = "ĐẬU" if tong >= 4.0 else "RỚT"
                data = list(row)
                data.append(ketqua)
                tree.insert("", tk.END, values=data)
        except Exception as e:
            messagebox.showerror("Lỗi", str(e))
        finally:
            conn.close()


def xuat_excel_xlsxwriter():
    """Hàm xuất Excel sử dụng thư viện xlsxwriter"""
    conn = connect_db()
    if not conn: return

    try:

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 title="Lưu file Bảng Điểm")
        if not file_path: return


        cursor = conn.cursor()
        sql = """
            SELECT SV.MaSV, SV.HoTen, MH.TenMon, 
                   D.DiemQuaTrinh, D.DiemGiuaKy, D.DiemCuoiKy, D.DiemTongKet,
                   CASE WHEN D.DiemTongKet >= 4.0 THEN N'ĐẬU' ELSE N'RỚT' END
            FROM DIEM D
            JOIN SINHVIEN SV ON D.MaSV = SV.MaSV
            JOIN HOCPHAN HP ON D.MaHocPhan = HP.MaHocPhan
            JOIN MONHOC MH ON HP.MaMon = MH.MaMon
        """
        cursor.execute(sql)
        rows = cursor.fetchall()


        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet("BangDiem")


        worksheet.set_column('A:A', 5)  # STT
        worksheet.set_column('B:B', 15)  # Mã SV
        worksheet.set_column('C:C', 25)  # Họ Tên
        worksheet.set_column('D:D', 20)  # Môn Học
        worksheet.set_column('E:H', 10)  # Các cột điểm
        worksheet.set_column('I:I', 15)  # Kết quả


        bold_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D7E4BC', 'align': 'center'})
        normal_format = workbook.add_format({'border': 1})
        center_format = workbook.add_format({'border': 1, 'align': 'center'})


        headers = ['STT', 'MÃ SV', 'HỌ TÊN', 'MÔN HỌC', 'ĐIỂM QT', 'ĐIỂM GK', 'ĐIỂM CK', 'TỔNG', 'KẾT QUẢ']
        for col, text in enumerate(headers):
            worksheet.write(0, col, text, bold_format)


        for i, row in enumerate(rows):
            row_num = i + 1

            worksheet.write(row_num, 0, i + 1, center_format)  # Cột STT
            worksheet.write(row_num, 1, row[0], center_format)  # Mã SV
            worksheet.write(row_num, 2, row[1], normal_format)  # Họ Tên
            worksheet.write(row_num, 3, row[2], normal_format)  # Môn
            worksheet.write(row_num, 4, row[3], center_format)  # QT
            worksheet.write(row_num, 5, row[4], center_format)  # GK
            worksheet.write(row_num, 6, row[5], center_format)  # CK
            worksheet.write(row_num, 7, row[6], center_format)  # Tổng


            if row[7] == 'RỚT':
                red_format = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'red', 'bold': True})
                worksheet.write(row_num, 8, row[7], red_format)
            else:
                worksheet.write(row_num, 8, row[7], center_format)



        workbook.close()
        messagebox.showinfo("Thành công", f"Đã xuất file Excel tại:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Lỗi Xuất Excel", str(e))
    finally:
        conn.close()


def popup_nhap_diem(parent_root, tree_diem):
    win = tk.Toplevel(parent_root)
    win.title("Nhập Điểm")
    center_window(win, 400, 500)  # Căn giữa popup

    tk.Label(win, text="NHẬP ĐIỂM", font=("bold", 14), fg="blue").pack(pady=10)
    f = tk.Frame(win);
    f.pack(pady=5)
    tk.Label(f, text="Mã SV:").grid(row=0, column=0);
    e_msv = tk.Entry(f);
    e_msv.grid(row=0, column=1)
    tk.Label(f, text="Mã HP:").grid(row=1, column=0);
    e_mhp = tk.Entry(f);
    e_mhp.grid(row=1, column=1)
    tk.Label(f, text="Điểm QT (20%):").grid(row=3, column=0);
    e_qt = tk.Entry(f);
    e_qt.grid(row=3, column=1)
    tk.Label(f, text="Điểm GK (30%):").grid(row=4, column=0);
    e_gk = tk.Entry(f);
    e_gk.grid(row=4, column=1)
    tk.Label(f, text="Điểm CK (50%):").grid(row=5, column=0);
    e_ck = tk.Entry(f);
    e_ck.grid(row=5, column=1)

    def luu():
        try:
            qt, gk, ck = float(e_qt.get()), float(e_gk.get()), float(e_ck.get())
        except:
            return messagebox.showerror("Lỗi", "Điểm phải là số")
        if not (0 <= qt <= 10 and 0 <= gk <= 10 and 0 <= ck <= 10): return messagebox.showerror("Lỗi", "Điểm 0-10")
        tong = round(qt * 0.2 + gk * 0.3 + ck * 0.5, 2)
        conn = connect_db()
        if conn:
            try:
                cur = conn.cursor()
                cur.execute("SELECT * FROM DIEM WHERE MaSV=? AND MaHocPhan=?", (e_msv.get(), e_mhp.get()))
                if cur.fetchone():
                    cur.execute(
                        "UPDATE DIEM SET DiemQuaTrinh=?, DiemGiuaKy=?, DiemCuoiKy=?, DiemTongKet=? WHERE MaSV=? AND MaHocPhan=?",
                        (qt, gk, ck, tong, e_msv.get(), e_mhp.get()))
                else:
                    cur.execute(
                        "INSERT INTO DIEM (MaSV, MaHocPhan, DiemQuaTrinh, DiemGiuaKy, DiemCuoiKy"
                        ", DiemTongKet) VALUES (?, ?, ?, ?, ?, ?)",
                        (e_msv.get(), e_mhp.get(), qt, gk, ck, tong))
                conn.commit()
                messagebox.showinfo("OK", f"Đã lưu! Tổng: {tong}")
                load_diem(tree_diem)
                win.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi SQL", str(e))
            finally:
                conn.close()

    tk.Button(win, text="LƯU", bg="blue", fg="white", command=luu).pack(pady=15)


# =========================================================================
# GIAO DIỆN CHÍNH
# =========================================================================
root = tk.Tk()
root.title("QUẢN LÝ ĐIỂM SINH VIÊN")
center_window(root, 1500, 600)  # Căn giữa cửa sổ chính

notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill="both")
tab1 = tk.Frame(notebook);
notebook.add(tab1, text=" SINH VIÊN ")
tab2 = tk.Frame(notebook);
notebook.add(tab2, text=" BẢNG ĐIỂM ")

# --- TAB 1 ---
f1 = tk.LabelFrame(tab1, text="Thông tin");
f1.pack(fill="x", padx=10)
tk.Label(f1, text="Mã SV:").grid(row=0, column=0);
e1_ma = tk.Entry(f1);
e1_ma.grid(row=0, column=1)
tk.Label(f1, text="Họ Tên:").grid(row=0, column=2);
e1_ten = tk.Entry(f1);
e1_ten.grid(row=0, column=3)
tk.Label(f1, text="Ngày Sinh:").grid(row=0, column=4);
e1_ngay = DateEntry(f1, date_pattern="yyyy-mm-dd");
e1_ngay.grid(row=0, column=5)
tk.Label(f1, text="Giới Tính:").grid(row=1, column=0);
e1_gt = ttk.Combobox(f1, values=["Nam", "Nu"]);
e1_gt.grid(row=1, column=1);
e1_gt.set("Nam")
tk.Label(f1, text="Mã Lớp:").grid(row=1, column=2);
e1_lop = tk.Entry(f1);
e1_lop.grid(row=1, column=3)
tk.Label(f1, text="Mã Khoa:").grid(row=1, column=4);
e1_khoa = tk.Entry(f1);
e1_khoa.grid(row=1, column=5)

btn = tk.Frame(tab1);
btn.pack(pady=5)
tk.Button(btn, text="Thêm", bg="green", fg="white",
          command=lambda: them_sv(tree1, e1_ma, e1_ten, e1_ngay, e1_gt, e1_lop, e1_khoa)).pack(side="left", padx=5)
tk.Button(btn, text="Xóa", bg="red", fg="white", command=lambda: xoa_sv(tree1)).pack(side="left", padx=5)
tk.Button(btn, text="Tải Lại", command=lambda: load_sv(tree1)).pack(side="left", padx=5)
cols1 = ("MaSV", "HoTen", "NgaySinh", "GioiTinh", "MaLop", "MaKhoa")
tree1 = ttk.Treeview(tab1, columns=cols1, show="headings", height=15)
for c in cols1: tree1.heading(c, text=c)
tree1.pack(fill="both", expand=True, padx=10)
load_sv(tree1)

# --- TAB 2 ---
f2_top = tk.Frame(tab2);
f2_top.pack(fill="x", padx=10, pady=10)
tk.Button(f2_top, text="+ NHẬP ĐIỂM", bg="blue", fg="white", command=lambda: popup_nhap_diem(root, tree2)).pack(
    side="left", padx=5)

# Nút Xuất Excel
tk.Button(f2_top, text="XUẤT EXCEL", bg="green", fg="white", command=xuat_excel_xlsxwriter).pack(side="left", padx=5)

tk.Button(f2_top, text="Tải lại", command=lambda: load_diem(tree2)).pack(side="left", padx=5)

cols2 = ("MaSV", "HoTen", "TenMon", "DiemQT", "DiemGK", "DiemCK", "TongKet", "KetQua")
tree2 = ttk.Treeview(tab2, columns=cols2, show="headings", height=15)
for c in cols2: tree2.heading(c, text=c)
tree2.pack(fill="both", expand=True, padx=10)
load_diem(tree2)

root.mainloop()