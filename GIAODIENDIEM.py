import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pyodbc


# ====== 1. CẤU HÌNH KẾT NỐI ======
def connect_db():
    try:
        conn = pyodbc.connect(
            "DRIVER={SQL Server};"
            "SERVER=LAPTOP-NI0S7IB6\SQL2017EXPRESS;"  # <--- HÃY SỬA TÊN SERVER CỦA BẠN Ở ĐÂY
            "DATABASE=QLDIEMSV;"
            "Trusted_Connection=yes;"
        )
        return conn
    except Exception as e:
        messagebox.showerror("Lỗi Kết Nối", f"Lỗi: {e}")
        return None


# =========================================================================
# TAB 1: LOGIC QUẢN LÝ SINH VIÊN
# =========================================================================
def load_sv(tree):
    for i in tree.get_children(): tree.delete(i)
    conn = connect_db()
    if conn:
        cur = conn.cursor()
        cur.execute("SELECT MaSV, HoTen, NgaySinh, GioiTinh, MaLop, MaKhoa FROM SINHVIEN")
        for row in cur.fetchall(): tree.insert("", tk.END, values=list(row))
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
    if messagebox.askyesno("Xóa", f"Xóa SV {masv} sẽ mất hết điểm của SV này. Tiếp tục?"):
        conn = connect_db()
        if conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM DIEM WHERE MaSV=?", (masv,))  # Xóa điểm trước
            cur.execute("DELETE FROM SINHVIEN WHERE MaSV=?", (masv,))  # Xóa SV sau
            conn.commit()
            conn.close()
            load_sv(tree)
            messagebox.showinfo("Xong", "Đã xóa!")


# =========================================================================
# TAB 2: LOGIC QUẢN LÝ ĐIỂM
# =========================================================================
def load_diem(tree):
    for i in tree.get_children(): tree.delete(i)
    conn = connect_db()
    if conn:
        try:
            cur = conn.cursor()
            # Kỹ thuật JOIN bảng để lấy Tên Môn thay vì chỉ hiện ID
            sql = """
                SELECT 
                    SV.MaSV, SV.HoTen, 
                    MH.TenMon, 
                    D.DiemQuaTrinh, D.DiemGiuaKy, D.DiemCuoiKy, D.DiemTongKet,
                    HP.MaHocPhan
                FROM DIEM D
                JOIN SINHVIEN SV ON D.MaSV = SV.MaSV
                JOIN HOCPHAN HP ON D.MaHocPhan = HP.MaHocPhan
                JOIN MONHOC MH ON HP.MaMon = MH.MaMon
            """
            cur.execute(sql)
            for row in cur.fetchall():
                # Xử lý xếp loại đơn giản
                tong = row[6] if row[6] else 0
                ketqua = "ĐẬU" if tong >= 4.0 else "RỚT"

                # Chèn vào bảng (Thêm cột Kết quả vào cuối)
                data = list(row)
                data.insert(7, ketqua)  # Chèn chữ Đậu/Rớt trước cột Mã HP ẩn
                tree.insert("", tk.END, values=data)
        except Exception as e:
            messagebox.showerror("Lỗi Tải Điểm", str(e))
        finally:
            conn.close()


def tim_kiem_diem(tree, keyword):

    for i in tree.get_children(): tree.delete(i)
    conn = connect_db()
    if conn:
        cur = conn.cursor()
        sql = """
            SELECT SV.MaSV, SV.HoTen, MH.TenMon, D.DiemQuaTrinh, D.DiemGiuaKy, D.DiemCuoiKy, D.DiemTongKet, HP.MaHocPhan
            FROM DIEM D
            JOIN SINHVIEN SV ON D.MaSV = SV.MaSV
            JOIN HOCPHAN HP ON D.MaHocPhan = HP.MaHocPhan
            JOIN MONHOC MH ON HP.MaMon = MH.MaMon
            WHERE SV.HoTen LIKE ? OR SV.MaSV LIKE ?
        """
        cur.execute(sql, (f'%{keyword}%', f'%{keyword}%'))
        for row in cur.fetchall():
            tong = row[6] if row[6] else 0
            ketqua = "ĐẬU" if tong >= 4.0 else "RỚT"
            data = list(row)
            data.insert(7, ketqua)
            tree.insert("", tk.END, values=data)
        conn.close()


def popup_nhap_diem(parent_root, tree_diem):
    win = tk.Toplevel(parent_root)
    win.title("Nhập / Sửa Điểm")
    win.geometry("400x500")

    tk.Label(win, text="NHẬP ĐIỂM SINH VIÊN", font=("bold", 14), fg="blue").pack(pady=10)

    # Form nhập
    f = tk.Frame(win);
    f.pack(pady=5)

    tk.Label(f, text="Mã Sinh Viên:").grid(row=0, column=0, pady=5, sticky="e")
    e_msv = tk.Entry(f);
    e_msv.grid(row=0, column=1)

    tk.Label(f, text="Mã Học Phần (ID):").grid(row=1, column=0, pady=5, sticky="e")
    e_mhp = tk.Entry(f);
    e_mhp.grid(row=1, column=1)
    tk.Label(f, text="(Xem bảng HOCPHAN trong SQL)", font=("Arial", 7), fg="gray").grid(row=2, column=1)

    tk.Label(f, text="Điểm Quá Trình (20%):").grid(row=3, column=0, pady=5, sticky="e")
    e_qt = tk.Entry(f);
    e_qt.grid(row=3, column=1)

    tk.Label(f, text="Điểm Giữa Kỳ (30%):").grid(row=4, column=0, pady=5, sticky="e")
    e_gk = tk.Entry(f);
    e_gk.grid(row=4, column=1)

    tk.Label(f, text="Điểm Cuối Kỳ (50%):").grid(row=5, column=0, pady=5, sticky="e")
    e_ck = tk.Entry(f);
    e_ck.grid(row=5, column=1)

    lbl_kq = tk.Label(win, text="...", font=("bold", 12), fg="red")
    lbl_kq.pack(pady=10)

    def xu_ly_luu():
        msv, mhp = e_msv.get(), e_mhp.get()
        try:
            qt, gk, ck = float(e_qt.get()), float(e_gk.get()), float(e_ck.get())
        except:
            return messagebox.showerror("Lỗi", "Điểm phải là số!")

        if not (0 <= qt <= 10 and 0 <= gk <= 10 and 0 <= ck <= 10):
            return messagebox.showerror("Lỗi", "Điểm từ 0-10")

        tong = round(qt * 0.2 + gk * 0.3 + ck * 0.5, 2)
        lbl_kq.config(text=f"Tổng kết: {tong}")

        conn = connect_db()
        if conn:
            try:
                cur = conn.cursor()

                cur.execute("SELECT * FROM DIEM WHERE MaSV=? AND MaHocPhan=?", (msv, mhp))
                if cur.fetchone():
                    sql = "UPDATE DIEM SET DiemQuaTrinh=?, DiemGiuaKy=?, DiemCuoiKy=?, DiemTongKet=? WHERE MaSV=? AND MaHocPhan=?"
                    cur.execute(sql, (qt, gk, ck, tong, msv, mhp))
                else:
                    sql = "INSERT INTO DIEM (MaSV, MaHocPhan, DiemQuaTrinh, DiemGiuaKy, DiemCuoiKy, DiemTongKet) VALUES (?, ?, ?, ?, ?, ?)"
                    cur.execute(sql, (msv, mhp, qt, gk, ck, tong))
                conn.commit()
                messagebox.showinfo("OK", "Đã lưu điểm!")
                load_diem(tree_diem)  # Load lại bảng điểm ở giao diện chính
                win.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi SQL", f"Kiểm tra lại Mã SV hoặc Mã HP.\n{e}")
            finally:
                conn.close()

    tk.Button(win, text="LƯU ĐIỂM", bg="blue", fg="white", font=("bold", 10), command=xu_ly_luu).pack(pady=5)


# =========================================================================
# GIAO DIỆN CHÍNH (MAIN WINDOW)
# =========================================================================
root = tk.Tk()
root.title("HỆ THỐNG QUẢN LÝ ĐIỂM SINH VIÊN")
root.geometry("900x600")

# --- TẠO TAB (NOTEBOOK) ---
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill="both")

# Tạo 2 khung cho 2 tab
tab1 = tk.Frame(notebook)
tab2 = tk.Frame(notebook)

notebook.add(tab1, text=" QUẢN LÝ HỒ SƠ SINH VIÊN ")
notebook.add(tab2, text=" QUẢN LÝ BẢNG ĐIỂM & KẾT QUẢ ")

# -------------------------------------------------------------------------
# THIẾT KẾ TAB 1: SINH VIÊN
# -------------------------------------------------------------------------
f1 = tk.LabelFrame(tab1, text="Nhập liệu")
f1.pack(fill="x", padx=10, pady=5)

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
e1_gt = ttk.Combobox(f1, values=["Nam", "Nu"], width=17);
e1_gt.grid(row=1, column=1);
e1_gt.set("Nam")
tk.Label(f1, text="Mã Lớp:").grid(row=1, column=2);
e1_lop = tk.Entry(f1);
e1_lop.grid(row=1, column=3)
tk.Label(f1, text="Mã Khoa:").grid(row=1, column=4);
e1_khoa = tk.Entry(f1);
e1_khoa.grid(row=1, column=5)

btn_f1 = tk.Frame(tab1);
btn_f1.pack(pady=5)
tk.Button(btn_f1, text="Thêm SV", bg="green", fg="white",
          command=lambda: them_sv(tree1, e1_ma, e1_ten, e1_ngay, e1_gt, e1_lop, e1_khoa)).pack(side="left", padx=5)
tk.Button(btn_f1, text="Xóa SV", bg="red", fg="white", command=lambda: xoa_sv(tree1)).pack(side="left", padx=5)
tk.Button(btn_f1, text="Tải lại danh sách", command=lambda: load_sv(tree1)).pack(side="left", padx=5)

cols1 = ("MaSV", "HoTen", "NgaySinh", "GioiTinh", "MaLop", "MaKhoa")
tree1 = ttk.Treeview(tab1, columns=cols1, show="headings", height=15)
for c in cols1: tree1.heading(c, text=c)
tree1.pack(fill="both", expand=True, padx=10, pady=5)
load_sv(tree1)

# -------------------------------------------------------------------------
# TAB 2: BẢNG ĐIỂM
# -------------------------------------------------------------------------
f2_top = tk.Frame(tab2);
f2_top.pack(fill="x", padx=10, pady=10)

tk.Button(f2_top, text="+ NHẬP ĐIỂM MỚI", bg="blue", fg="white", font=("bold", 10), height=2,
          command=lambda: popup_nhap_diem(root, tree2)).pack(side="left", padx=5)

tk.Label(f2_top, text="Tìm tên SV:").pack(side="left", padx=(30, 5))
e2_tim = tk.Entry(f2_top)
e2_tim.pack(side="left")
tk.Button(f2_top, text="Tìm kiếm", command=lambda: tim_kiem_diem(tree2, e2_tim.get())).pack(side="left", padx=5)
tk.Button(f2_top, text="Hiện tất cả", command=lambda: load_diem(tree2)).pack(side="left", padx=5)

# Bảng điểm chi tiết
cols2 = ("MaSV", "HoTen", "TenMon", "DiemQT", "DiemGK", "DiemCK", "TongKet", "KetQua")
tree2 = ttk.Treeview(tab2, columns=cols2, show="headings", height=15)

tree2.heading("MaSV", text="Mã SV")
tree2.heading("HoTen", text="Họ Tên")
tree2.heading("TenMon", text="Môn Học")
tree2.heading("DiemQT", text="QT (20%)")
tree2.heading("DiemGK", text="GK (30%)")
tree2.heading("DiemCK", text="CK (50%)")
tree2.heading("TongKet", text="T.Kết")
tree2.heading("KetQua", text="Kết Quả")

tree2.column("MaSV", width=80)
tree2.column("HoTen", width=150)
tree2.column("TenMon", width=150)
tree2.column("DiemQT", width=60, anchor="center")
tree2.column("DiemGK", width=60, anchor="center")
tree2.column("DiemCK", width=60, anchor="center")
tree2.column("TongKet", width=60, anchor="center")
tree2.column("KetQua", width=80, anchor="center")

tree2.pack(fill="both", expand=True, padx=10, pady=5)
load_diem(tree2)

root.mainloop()