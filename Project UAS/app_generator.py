import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime



# Daftar Dosen  
dosen_input = [
    'Alun Sujjada, ST., M.Kom', 'Anggun Fergina, M.Kom', 'Gina Purnama Insany, S.Si.T., M.Kom',
    'Ir. Somantri, ST., M.Kom', 'Ivana Lucia Kharisma, M.Kom', 'Ir. Kamdan, M.Kom',
    'Dhita Diana Dewi, M.Stat', 'Lusiana Sani Parwati, M.Mat', 'Drs. Nuzwan Sudariana, MM',
    'Syahid Abdullah, S.Si., M.Kom', 'Hermanto, M.Kom', 'Nugraha, M.Kom', 'Imam Sanjaya, SP., M.Kom',
    'Zaenal Alamsyah, M.Kom', 'M.Ikhsan Thohir, M.Kom', 'Adrian Reza, M.Kom', 'Shinta Ayuningtyas, M.Kom',
    'Moneyta Dholah Rosita, M.Kom', 'Mega Lumbia Octavia Sinaga, M.Kom', 'Indra Yustiana, M.Kom',
    'Harris Al Qodri Maarif, S.T., M.Sc. PhD', 'Dr. Iwan Setiawan, S.T., M.T', 'Dede Sukmawan, M.Kom',
    'Falentino Sembiring, M.Kom', 'Dr. Huang Gan', 'Muchtar Ali Setyo Yudono, S.T., M.T',
    'Dr. Deni Hasman', 'Dr. Nurkhan', 'Dr. Yurman Zaenal', 'Zaenal Alamsyah, M.Kom',
    'Indra Yustiana, M.Kom', 'Ir. Somantri', 'Anggun Fergina, M.Kom', 'Gina Purnama Insany, S.Si.T., M.Kom',
    'Moneyta Dholah Rosita, M.Kom', 'Mega Lumbia Octavia Sinaga, M.Kom', 'Shinta Ayuningtyas, M.Kom'
]
DOSEN_LIST = sorted(list(set(dosen_input)))

# Daftar Mata Kuliah  
MATAKULIAH_LIST = sorted([
    'Algoritma dan Struktur Data', 'Pemrograman Berbasis Platform', 'Kompleksitas Algoritma',
    'Pengolahan Citra Digital', 'Pemrograman Berbasis Web', 'Jaringan Komputer dan Keamanan Informasi',
    'Sistem Paralel dan Terdistribusi', 'Rekayasa Perangkat Lunak', 'Basis Data',
    'Projek Perangkat Lunak', 'Logika Informatika', 'Statistika dan Probabilitas',
    'Metodologi Penelitian', 'Data Science', 'Cyber Security', 'Sistem Informasi Geografis',
    'Big Data Arsitektur dan Infrastruktur', 'Interaksi Manusia dan Komputer', 'Computer Vision',
    'Deep Learning', 'Organisasi dan Arsitektur Komputer', 'Kalkulus', 'Metode Numerik',
    'Pemrograman Berbasis Mobile', 'Pengolahan Perangkat Lunak', 'Etika Profesi', 'Teknologi Blockchain'
])

# --- MODIFIKASI: Menambahkan daftar ruangan yang bisa dipilih ---
RUANGAN_LIST = sorted([
    "B4H","B3B","B4C","B4A"
])

SEMESTER_LIST = list(range(1, 9))
SKS_LIST = list(range(1, 7))
MODE_LIST = ['OFFLINE', 'ONLINE']

# --- KONFIGURASI DAN DATA ---
AVAILABLE_ROOMS = {'A': {'floors': list(range(1, 6))}, 'B': {'floors': list(range(1, 7))}}
ALLOWED_DAYS = ['SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT']
MIN_TIME = datetime.strptime("08:00", "%H:%M").time()
MAX_TIME = datetime.strptime("20:00", "%H:%M").time()
BREAKS = [(datetime.strptime("12:00", "%H:%M").time(), datetime.strptime("13:00", "%H:%M").time()), (datetime.strptime("18:00", "%H:%M").time(), datetime.strptime("19:00", "%H:%M").time())]
NAMA_FILE_EXCEL = 'reservasi_ruangan.xlsx'

KOLOM_WAJIB = ['HARI', 'DOSEN', 'MATAKULIAH', 'SEMESTER', 'SKS', 'KELAS', 'MODE', 'PRODI', 'GEDUNG', 'LANTAI', 'RUANGAN', 'MULAI', 'SELESAI']
KOLOM_FORM = ['HARI', 'DOSEN', 'MATAKULIAH', 'SEMESTER', 'SKS', 'KELAS', 'MODE', 'GEDUNG', 'LANTAI', 'RUANGAN', 'MULAI', 'SELESAI']

# (Fungsi validasi tidak berubah)
def is_time_slot_valid(start_str, end_str):
    try: start_t, end_t = datetime.strptime(start_str, "%H:%M").time(), datetime.strptime(end_str, "%H:%M").time()
    except (ValueError, TypeError): return False, "Format waktu salah. Gunakan HH:MM."
    if not (MIN_TIME <= start_t < MAX_TIME and MIN_TIME < end_t <= MAX_TIME and start_t < end_t): return False, f"Waktu reservasi harus antara {MIN_TIME:%H:%M} dan {MAX_TIME:%H:%M}."
    for s, e in BREAKS:
        if s <= start_t < e: return False, f"Reservasi tidak boleh DIMULAI di dalam jam istirahat ({s:%H:%M}-{e:%H:%M})."
    return True, ""
def is_room_available(df, day, start_str, end_str, gedung, lantai, ruangan, ignore_index=None):
    temp_df = df.copy()
    if ignore_index is not None: temp_df = temp_df.drop(index=ignore_index)
    start_t, end_t = datetime.strptime(start_str, "%H:%M").time(), datetime.strptime(end_str, "%H:%M").time()
    bookings = temp_df[(temp_df['HARI'] == day) & (temp_df['GEDUNG'] == gedung) & (temp_df['LANTAI'] == lantai) & (temp_df['RUANGAN'] == ruangan)]
    for _, row in bookings.iterrows():
        booked_start, booked_end = datetime.strptime(row['MULAI'], "%H:%M").time(), datetime.strptime(row['SELESAI'], "%H:%M").time()
        if max(start_t, booked_start) < min(end_t, booked_end): return False, f"Lokasi ini sudah dipesan dari jam {row['MULAI']}-{row['SELESAI']}."
    return True, ""

# ===================================================================================
# KELAS APLIKASI GUI
# ===================================================================================

class ReservationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Reservasi Ruangan")
        self.root.geometry("1366x768")
        self.df = self.load_data()
        
        main_frame = ttk.Frame(self.root, padding="10"); main_frame.pack(fill=tk.BOTH, expand=True)
        tree_frame = ttk.Frame(main_frame); tree_frame.pack(fill=tk.BOTH, expand=True)
        tree_scroll_y = ttk.Scrollbar(tree_frame); tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient='horizontal'); tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree = ttk.Treeview(tree_frame, columns=KOLOM_WAJIB, show='headings', yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        tree_scroll_y.config(command=self.tree.yview); tree_scroll_x.config(command=self.tree.xview)
        for col in KOLOM_WAJIB:
            self.tree.heading(col, text=col); self.tree.column(col, width=100, anchor='w')
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_item_select)

        form_frame = ttk.LabelFrame(main_frame, text="Detail Reservasi", padding="10"); form_frame.pack(fill=tk.X, pady=10)
        self.entries = {}
        
        # --- MODIFIKASI: Widget untuk 'RUANGAN' diubah dari Entry menjadi Combobox ---
        widget_configs = {
            'HARI': {'widget': ttk.Combobox, 'options': {'values': ALLOWED_DAYS}}, 'DOSEN': {'widget': ttk.Combobox, 'options': {'values': DOSEN_LIST}},
            'MATAKULIAH': {'widget': ttk.Combobox, 'options': {'values': MATAKULIAH_LIST}}, 'SEMESTER': {'widget': ttk.Combobox, 'options': {'values': SEMESTER_LIST}},
            'SKS': {'widget': ttk.Combobox, 'options': {'values': SKS_LIST}}, 'KELAS': {'widget': ttk.Entry},
            'MODE': {'widget': ttk.Combobox, 'options': {'values': MODE_LIST}}, 'GEDUNG': {'widget': ttk.Combobox, 'options': {'values': list(AVAILABLE_ROOMS.keys())}},
            'LANTAI': {'widget': ttk.Combobox, 'options': {}}, 
            'RUANGAN': {'widget': ttk.Combobox, 'options': {'values': RUANGAN_LIST}}, # <-- DIUBAH DI SINI
            'MULAI': {'widget': ttk.Entry}, 'SELESAI': {'widget': ttk.Entry}
        }
        items_per_row = 6
        for i, (label, config) in enumerate(widget_configs.items()):
            row_num, col_num = i // items_per_row, (i % items_per_row) * 2
            lbl = ttk.Label(form_frame, text=label); lbl.grid(row=row_num, column=col_num, padx=5, pady=5, sticky='w')
            widget = config['widget'](form_frame, width=25, **config.get('options', {}))
            if isinstance(widget, ttk.Combobox): widget.config(state='readonly')
            widget.grid(row=row_num, column=col_num + 1, padx=5, pady=5, sticky='ew')
            self.entries[label] = widget
            
        self.entries['MODE'].bind('<<ComboboxSelected>>', self.on_mode_select)
        self.entries['GEDUNG'].bind('<<ComboboxSelected>>', self.update_location_options)

        button_frame = ttk.Frame(main_frame, padding="10"); button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Tambah Reservasi", command=self.add_reservation).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(button_frame, text="Update Reservasi", command=self.update_reservation).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(button_frame, text="Hapus Reservasi", command=self.delete_reservation).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(button_frame, text="Clear Form", command=self.clear_form).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(button_frame, text="Ekspor ke Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.populate_treeview()
        self.on_mode_select()

    def save_data_auto(self):
        try: self.df.to_excel(NAMA_FILE_EXCEL, index=False)
        except Exception as e: messagebox.showerror("Error Penyimpanan Otomatis", f"Gagal menyimpan ke {NAMA_FILE_EXCEL}:\n{e}")

    def export_to_excel(self):
        if self.df.empty: messagebox.showwarning("Peringatan", "Tidak ada data untuk diekspor."); return
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")], title="Simpan Data Reservasi")
            if file_path: self.df.to_excel(file_path, index=False); messagebox.showinfo("Sukses", f"Data berhasil diekspor ke:\n{file_path}")
        except Exception as e: messagebox.showerror("Error Ekspor", f"Gagal mengekspor file:\n{e}")

    # --- MODIFIKASI: Logika untuk widget RUANGAN disesuaikan karena sudah menjadi Combobox ---
    def on_mode_select(self, event=None):
        mode = self.entries['MODE'].get()
        gedung_widget = self.entries['GEDUNG']
        lantai_widget = self.entries['LANTAI']
        ruangan_widget = self.entries['RUANGAN'] # Ini sekarang adalah Combobox

        if mode == 'ONLINE':
            gedung_widget.set('-'); gedung_widget.config(state='disabled')
            lantai_widget.set('-'); lantai_widget.config(state='disabled')
            ruangan_widget.set('-'); ruangan_widget.config(state='disabled') # Cukup set dan disable
        else: # OFFLINE
            gedung_widget.config(state='readonly')
            lantai_widget.config(state='readonly')
            ruangan_widget.config(state='readonly') # Ganti ke readonly untuk Combobox
            
            # Bersihkan nilai default jika ada
            if gedung_widget.get() == '-': gedung_widget.set('')
            if lantai_widget.get() == '-': lantai_widget.set('')
            if ruangan_widget.get() == '-': ruangan_widget.set('')
            
            self.update_location_options()

    def update_location_options(self, event=None):
        selected_gedung = self.entries['GEDUNG'].get()
        self.entries['LANTAI'].set('')
        if selected_gedung and selected_gedung != '-':
            self.entries['LANTAI']['values'] = AVAILABLE_ROOMS[selected_gedung]['floors']
        else:
            self.entries['LANTAI']['values'] = []

    def on_item_select(self, event):
        if not self.tree.selection(): return
        values = self.tree.item(self.tree.selection()[0], 'values')
        self.clear_form(clear_selection=False)
        for key in KOLOM_FORM:
            # Kode ini sudah dapat menangani Combobox secara otomatis, tidak perlu diubah
            value = values[KOLOM_WAJIB.index(key)]
            if isinstance(self.entries[key], ttk.Combobox):
                self.entries[key].set(value)
            else:
                self.entries[key].insert(0, value)
        self.on_mode_select()
        self.update_location_options()
        self.entries['LANTAI'].set(values[KOLOM_WAJIB.index('LANTAI')])

    def load_data(self):
        try:
            df = pd.read_excel(NAMA_FILE_EXCEL)
            for col in KOLOM_WAJIB:
                if col not in df.columns: df[col] = ''
            for col in ['LANTAI', 'SEMESTER', 'SKS']:
                # Mengubah 'coerce' menjadi 'ignore' bisa lebih aman jika ada nilai non-numerik seperti '-'
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
            df = df[KOLOM_WAJIB]
            return df
        except FileNotFoundError: return pd.DataFrame(columns=KOLOM_WAJIB)
    
    def populate_treeview(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for _, row in self.df.iterrows(): self.tree.insert("", tk.END, values=list(row.astype(str)))
    
    def clear_form(self, clear_selection=True):
        if clear_selection and self.tree.selection(): self.tree.selection_remove(self.tree.selection()[0])
        for widget in self.entries.values():
            if isinstance(widget, ttk.Combobox): widget.set('')
            else: 
                if widget.cget('state') != 'disabled':
                    widget.delete(0, tk.END)
        self.on_mode_select()

    def _get_and_validate_form_data(self):
        data = {key: widget.get().strip() for key, widget in self.entries.items()}
        required_fields = [f for f in KOLOM_FORM if f not in ['GEDUNG', 'LANTAI', 'RUANGAN']]
        for field in required_fields:
            if not data.get(field): messagebox.showerror("Input Tidak Lengkap", f"Kolom '{field}' harus diisi."); return None
        mode = data.get('MODE')
        if mode == 'OFFLINE':
            for field in ['GEDUNG', 'LANTAI', 'RUANGAN']:
                if not data.get(field) or data.get(field) == '-':
                    messagebox.showerror("Input Tidak Lengkap", f"Untuk mode OFFLINE, kolom '{field}' harus diisi."); return None
        try:
            data['LANTAI_INT'] = int(data['LANTAI']) if mode == 'OFFLINE' else 0
            data['SEMESTER_INT'] = int(data['SEMESTER']); data['SKS_INT'] = int(data['SKS'])
        except (ValueError, TypeError): messagebox.showerror("Input Salah", "Kolom numerik (Lantai, Semester, SKS) harus berupa angka."); return None
        data['HARI_UPPER'] = data['HARI'].upper()
        # Logika ini tetap valid, karena .get() bekerja untuk Combobox juga
        data['RUANGAN_UPPER'] = data.get('RUANGAN', '').upper() if mode == 'OFFLINE' else '-'
        data['GEDUNG_UPPER'] = data['GEDUNG'].upper() if mode == 'OFFLINE' else '-'
        valid, msg = is_time_slot_valid(data['MULAI'], data['SELESAI']);
        if not valid: messagebox.showerror("Error Waktu", msg); return None
        return data
        
    def add_reservation(self):
        data = self._get_and_validate_form_data()
        if data is None: return
        if data.get('MODE') == 'OFFLINE':
            valid, msg = is_room_available(self.df, data['HARI_UPPER'], data['MULAI'], data['SELESAI'], data['GEDUNG_UPPER'], data['LANTAI_INT'], data['RUANGAN_UPPER'])
            if not valid: messagebox.showerror("Jadwal Bentrok", msg); return
        
        new_data = {
            'HARI': data['HARI_UPPER'], 'DOSEN': data['DOSEN'], 'MATAKULIAH': data['MATAKULIAH'],
            'SEMESTER': data['SEMESTER_INT'], 'SKS': data['SKS_INT'], 'KELAS': data['KELAS'].upper(),
            'MODE': data['MODE'], 'PRODI': data['KELAS'][:2].upper(), 'GEDUNG': data['GEDUNG_UPPER'],
            'LANTAI': data['LANTAI_INT'], 'RUANGAN': data['RUANGAN_UPPER'], 'MULAI': data['MULAI'], 'SELESAI': data['SELESAI']
        }
        new_df = pd.DataFrame([new_data]); self.df = pd.concat([self.df, new_df], ignore_index=True)
        messagebox.showinfo("Sukses", "Reservasi berhasil ditambahkan."); self.save_data_auto(); self.populate_treeview(); self.clear_form()

    def update_reservation(self):
        if not self.tree.selection(): messagebox.showwarning("Peringatan", "Pilih data yang ingin diupdate dari tabel."); return
        data = self._get_and_validate_form_data()
        if data is None: return
        df_index_to_update = self.df.index[self.tree.index(self.tree.selection()[0])]
        if data.get('MODE') == 'OFFLINE':
            valid, msg = is_room_available(self.df, data['HARI_UPPER'], data['MULAI'], data['SELESAI'], data['GEDUNG_UPPER'], data['LANTAI_INT'], data['RUANGAN_UPPER'], ignore_index=df_index_to_update)
            if not valid: messagebox.showerror("Jadwal Bentrok", msg); return
        
        self.df.loc[df_index_to_update, 'HARI'] = data['HARI_UPPER']
        self.df.loc[df_index_to_update, 'DOSEN'] = data['DOSEN']
        self.df.loc[df_index_to_update, 'MATAKULIAH'] = data['MATAKULIAH']
        self.df.loc[df_index_to_update, 'SEMESTER'] = data['SEMESTER_INT']
        self.df.loc[df_index_to_update, 'SKS'] = data['SKS_INT']
        self.df.loc[df_index_to_update, 'KELAS'] = data['KELAS'].upper()
        self.df.loc[df_index_to_update, 'MODE'] = data['MODE']
        self.df.loc[df_index_to_update, 'PRODI'] = data['KELAS'][:2].upper()
        self.df.loc[df_index_to_update, 'GEDUNG'] = data['GEDUNG_UPPER']
        self.df.loc[df_index_to_update, 'LANTAI'] = data['LANTAI_INT']
        self.df.loc[df_index_to_update, 'RUANGAN'] = data['RUANGAN_UPPER']
        self.df.loc[df_index_to_update, 'MULAI'] = data['MULAI']
        self.df.loc[df_index_to_update, 'SELESAI'] = data['SELESAI']

        messagebox.showinfo("Sukses", "Reservasi berhasil diperbarui."); self.save_data_auto(); self.populate_treeview(); self.clear_form()

    def delete_reservation(self):
        if not self.tree.selection(): messagebox.showwarning("Peringatan", "Pilih data yang ingin dihapus dari tabel."); return
        if messagebox.askyesno("Konfirmasi Hapus", "Apakah Anda yakin ingin menghapus data ini?"):
            indices_to_drop = [self.df.index[self.tree.index(item)] for item in self.tree.selection()]
            self.df = self.df.drop(indices_to_drop).reset_index(drop=True)
            self.save_data_auto(); self.populate_treeview(); self.clear_form()
            messagebox.showinfo("Sukses", "Reservasi berhasil dihapus.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReservationApp(root)
    root.mainloop() 