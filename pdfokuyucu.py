import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber
import os
import re
import requests
import json
import pandas as pd
import sv_ttk  # Modern tema için

# --- Backend Ayarları ---
BASE_URL = "https://isiyer-app-a98cc9d8425a.herokuapp.com"
YARDS_API_URL = f"{BASE_URL}/api/yards"
PRODUCTS_API_URL = f"{BASE_URL}/api/products"
santiyeler_map = {}

# --- GENEL FONKSİYONLAR ---

def santiyeleri_getir(combobox_widget):
    try:
        response = requests.get(YARDS_API_URL)
        if response.status_code == 200:
            santiyeler = response.json()
            santiye_isimleri = [santiye.get("yardName") for santiye in santiyeler if santiye.get("yardName")]
            
            santiyeler_map.clear()
            for santiye in santiyeler:
                santiyeler_map[santiye.get("yardName")] = santiye.get("id")
            
            combobox_widget['values'] = santiye_isimleri
            if santiye_isimleri:
                combobox_widget.current(0)
            return True
        else:
            messagebox.showerror("Bağlantı Hatası", f"Şantiyeler alınamadı. Sunucu: {response.status_code}")
            return False
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}")
        return False

# --- SEKME 1: FATURA GİRİŞ ARAYÜZÜNÜ OLUŞTURAN FONKSİYON ---

def create_fatura_tab(tab):
    main_frame = ttk.Frame(tab, padding="10")
    main_frame.pack(fill="both", expand=True)

    top_frame = ttk.LabelFrame(main_frame, text="Kontroller", padding="10")
    top_frame.pack(fill="x", pady=(0, 10))
    top_frame.columnconfigure(1, weight=1)

    ttk.Label(top_frame, text="Kaydedilecek Şantiye:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
    santiye_secimi_cb = ttk.Combobox(top_frame, state="readonly")
    santiye_secimi_cb.grid(row=0, column=1, pady=5, sticky="ew")

    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill="both", expand=True)
    tree = ttk.Treeview(tree_frame, columns=("h_kod", "h_aciklama", "miktar", "birim"), show="headings")
    tree.heading("h_kod", text="Hizmet Kodu"); tree.column("h_kod", width=150, anchor="w")
    tree.heading("h_aciklama", text="Hizmet Açıklaması"); tree.column("h_aciklama", width=450, anchor="w")
    tree.heading("miktar", text="Miktar"); tree.column("miktar", width=100, anchor="center")
    tree.heading("birim", text="Birim"); tree.column("birim", width=100, anchor="center")
    tree.tag_configure('oddrow', background='#F0F0F0')
    tree.tag_configure('evenrow', background='white')
    
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)

    def pdf_oku_ve_doldur():
        dosya_yolu = filedialog.askopenfilename(title="Fatura PDF'ini Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not dosya_yolu: return
        for i in tree.get_children(): tree.delete(i)
        try:
            with pdfplumber.open(dosya_yolu) as pdf:
                sayfa = pdf.pages[0]
                tablolar = sayfa.extract_tables()
                hedef_tablo = None
                for tablo in tablolar:
                    if tablo:
                        baslik_metni = " ".join(filter(None, tablo[0])).replace('\n', ' ')
                        if "Malzeme/Hizmet Kodu" in baslik_metni and "KDV Oran" in baslik_metni:
                            hedef_tablo = tablo
                            break
                if not hedef_tablo:
                    messagebox.showerror("Hata", "Fatura kalemlerini içeren ana tablo bulunamadı.")
                    return
                for i, satir_data in enumerate(range(1, len(hedef_tablo))):
                    satir = hedef_tablo[satir_data]
                    hizmet_kodu, hizmet_aciklamasi, miktar_ham = satir[1], satir[2], satir[4]
                    if hizmet_kodu and hizmet_aciklamasi and miktar_ham:
                        miktar_ham = miktar_ham.replace('\n', ' ').strip()
                        eslesme = re.match(r"([\d.,]+)\s*([a-zA-Z]+)", miktar_ham)
                        sayisal_miktar, birim = ("", "")
                        if eslesme:
                            sayisal_miktar, birim_kisaltma = eslesme.group(1), eslesme.group(2)
                            if birim_kisaltma.upper() == 'M': birim = "Metre"
                            elif birim_kisaltma.upper() == 'ADET': birim = "Adet"
                            else: birim = birim_kisaltma
                        else:
                            sayisal_miktar = miktar_ham
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                        tree.insert("", "end", values=(hizmet_kodu.replace('\n', ' '), hizmet_aciklamasi.replace('\n', ' '), sayisal_miktar, birim), tags=(tag,))
        except Exception as e:
            messagebox.showerror("Hata", f"PDF işlenirken bir hata oluştu: {e}")

    def verileri_kaydet_tek_tek():
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi:
            messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin.")
            return
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)
        if not tree.get_children():
            messagebox.showwarning("Eksik Bilgi", "Kaydedilecek veri bulunmuyor.")
            return
        basarili_kayit, hatali_kayit = 0, 0
        toplam_kayit = len(tree.get_children())
        for row_id in tree.get_children():
            item = tree.item(row_id)['values']
            try:
                miktar_sayi = int(str(item[2]).replace('.', '').replace(',', ''))
            except ValueError:
                hatali_kayit += 1
                continue
            birim_str = str(item[3]).upper()
            birim_enum = "ADET"
            if birim_str == "METRE": birim_enum = "METRE"
            elif birim_str == "ADET": birim_enum = "ADET"
            gonderilecek_tek_urun = {"code": item[0], "description": item[1], "amount": miktar_sayi, "unit": birim_enum}
            try:
                url = f"{PRODUCTS_API_URL}/yards/{secili_santiye_id}/products"
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, data=json.dumps(gonderilecek_tek_urun), headers=headers)
                if response.status_code in [200, 201]:
                    basarili_kayit += 1
                else:
                    hatali_kayit += 1
            except requests.exceptions.RequestException:
                messagebox.showerror("Bağlantı Hatası", "Sunucuya bağlanılamadı. İşlem durduruldu.")
                return
        messagebox.showinfo("İşlem Tamamlandı", f"Toplam {toplam_kayit} üründen:\nBaşarıyla kaydedilen: {basarili_kayit}\nHatalı: {hatali_kayit}")
        if basarili_kayit > 0:
            for i in tree.get_children():
                tree.delete(i)

    button_group = ttk.Frame(top_frame)
    button_group.grid(row=0, column=2, padx=(20, 0), pady=5, sticky="e")
    ttk.Button(button_group, text="PDF Yükle", command=pdf_oku_ve_doldur, style='Accent.TButton').pack(side="left", padx=(0, 5))
    ttk.Button(button_group, text="Verileri Kaydet", command=verileri_kaydet_tek_tek).pack(side="left")
    
    santiyeleri_getir(santiye_secimi_cb)

# --- SEKME 2: ŞANTİYE SORGULAMA ARAYÜZÜNÜ OLUŞTURAN FONKSİYON ---

def create_sorgu_tab(tab):
    main_frame = ttk.Frame(tab, padding="10")
    main_frame.pack(fill="both", expand=True)

    top_frame = ttk.LabelFrame(main_frame, text="Kontroller", padding="10")
    top_frame.pack(fill="x", pady=(0, 10))
    top_frame.columnconfigure(1, weight=1)

    ttk.Label(top_frame, text="Görüntülenecek Şantiye:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
    santiye_secimi_cb = ttk.Combobox(top_frame, state="readonly")
    santiye_secimi_cb.grid(row=0, column=1, pady=5, sticky="ew")
    
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill="both", expand=True)
    tree = ttk.Treeview(tree_frame, columns=("code", "description", "amount", "unit"), show="headings")
    tree.heading("code", text="Ürün Kodu"); tree.column("code", width=150, anchor="w")
    tree.heading("description", text="Açıklama"); tree.column("description", width=450, anchor="w")
    tree.heading("amount", text="Miktar"); tree.column("amount", width=100, anchor="center")
    tree.heading("unit", text="Birim"); tree.column("unit", width=100, anchor="center")
    tree.tag_configure('oddrow', background='#F0F0F0')
    tree.tag_configure('evenrow', background='white')
    
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)

    def verileri_getir():
        for i in tree.get_children(): tree.delete(i)
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi:
            messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin.")
            return
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)
        try:
            url = f"{YARDS_API_URL}/{secili_santiye_id}"
            response = requests.get(url)
            if response.status_code == 200:
                veri = response.json()
                urunler = veri.get("products", []) 
                if not urunler:
                    messagebox.showinfo("Bilgi", "Bu şantiyeye ait ürün bulunamadı.")
                for i, urun in enumerate(urunler):
                    tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                    tree.insert("", "end", values=(urun['code'], urun['description'], urun['amount'], urun['unit']), tags=(tag,))
            else:
                messagebox.showerror("Hata", f"Veriler alınamadı. Sunucu: {response.status_code}\n{response.text}")
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}")

    def verileri_excele_aktar():
        if not tree.get_children():
            messagebox.showwarning("Veri Yok", "Aktarılacak veri bulunmuyor.")
            return
        try:
            dosya_yolu = filedialog.asksaveasfilename(
                initialfile=f'{santiye_secimi_cb.get()}_urun_listesi.xlsx',
                defaultextension=".xlsx",
                filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")])
            if not dosya_yolu: return
            veri_listesi = []
            for row_id in tree.get_children():
                item = tree.item(row_id)['values']
                veri_listesi.append({'Ürün Kodu': item[0], 'Açıklama': item[1], 'Miktar': item[2], 'Birim': item[3]})
            df = pd.DataFrame(veri_listesi)
            df.to_excel(dosya_yolu, index=False)
            messagebox.showinfo("Başarılı", f"Veriler başarıyla şu dosyaya aktarıldı:\n{dosya_yolu}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e aktarılırken bir hata oluştu:\n{e}")

    button_group = ttk.Frame(top_frame)
    button_group.grid(row=0, column=2, padx=(20, 0))
    ttk.Button(button_group, text="Verileri Getir", command=verileri_getir, style='Accent.TButton').pack(side="left", padx=(0, 5))
    ttk.Button(button_group, text="Excel'e Aktar", command=verileri_excele_aktar).pack(side="left")
    
    santiyeleri_getir(santiye_secimi_cb)

# --- ANA UYGULAMA PENCERESİ VE SEKMELER ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Veri Yönetim Sistemi")
    root.geometry("950x700")

    sv_ttk.set_theme("light")

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Segoe UI", 10, 'bold'))

    notebook = ttk.Notebook(root)
    notebook.pack(expand=True, fill='both', padx=10, pady=10)

    fatura_tab = ttk.Frame(notebook)
    sorgu_tab = ttk.Frame(notebook)

    notebook.add(fatura_tab, text='Faturadan Veri Girişi')
    notebook.add(sorgu_tab, text='Şantiye Verilerini Görüntüle')

    create_fatura_tab(fatura_tab)
    create_sorgu_tab(sorgu_tab)

    root.mainloop()