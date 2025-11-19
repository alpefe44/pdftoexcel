import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber
import os
import re
import requests
import json
import pandas as pd
import sv_ttk

# --- Backend Ayarları ---
BASE_URL = "https://isiyer-app-a98cc9d8425a.herokuapp.com"
YARDS_API_URL = f"{BASE_URL}/api/yards"
PRODUCTS_API_URL = f"{BASE_URL}/api/products"
santiyeler_map = {}
combobox_references = []

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
            messagebox.showerror("Bağlantı Hatası", f"Şantiyeler alınamadı. Sunucu: {response.status_code}\n{response.text}")
            return False
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}")
        return False

# --- YENİ ŞANTİYE EKLEME ---

def open_add_yard_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("Yeni Şantiye Ekle")
    dialog.geometry("350x150")
    dialog.resizable(False, False)
    dialog.grab_set()
    dialog.transient(root)

    main_frame = ttk.Frame(dialog, padding="15")
    main_frame.pack(fill="both", expand=True)
    main_frame.columnconfigure(1, weight=1)

    ttk.Label(main_frame, text="Şantiye Adı:").grid(row=0, column=0, padx=(0, 10), sticky="w")
    yard_name_entry = ttk.Entry(main_frame)
    yard_name_entry.grid(row=0, column=1, sticky="ew")
    yard_name_entry.focus()

    def save_new_yard():
        yard_name = yard_name_entry.get().strip()
        if not yard_name:
            messagebox.showwarning("Eksik Bilgi", "Şantiye adı boş bırakılamaz.", parent=dialog)
            return

        payload = {"yardName": yard_name}
        try:
            url = f"{YARDS_API_URL}/add"
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, data=json.dumps(payload), headers=headers)

            if response.status_code == 201:
                messagebox.showinfo("Başarılı", f"'{yard_name}' şantiyesi başarıyla eklendi.", parent=dialog)
                dialog.destroy()
                for cb in combobox_references:
                    santiyeleri_getir(cb)
            else:
                messagebox.showerror("Hata", f"Şantiye eklenemedi. Sunucu: {response.status_code}\n{response.text}", parent=dialog)
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}", parent=dialog)

    button_frame = ttk.Frame(main_frame)
    button_frame.grid(row=1, column=0, columnspan=2, pady=(20, 0), sticky="e")
    ttk.Button(button_frame, text="Kaydet", command=save_new_yard, style='Accent.TButton').pack()

# --- SEKME 1: FATURA GİRİŞİ ---

def create_fatura_tab(tab):
    main_frame = ttk.Frame(tab, padding="10")
    main_frame.pack(fill="both", expand=True)
    top_frame = ttk.LabelFrame(main_frame, text="Kontroller", padding="10")
    top_frame.pack(fill="x", pady=(0, 10))
    top_frame.columnconfigure(1, weight=1)
    ttk.Label(top_frame, text="Kaydedilecek Şantiye:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
    
    cb_frame = ttk.Frame(top_frame)
    cb_frame.grid(row=0, column=1, pady=5, sticky="ew")
    santiye_secimi_cb = ttk.Combobox(cb_frame, state="readonly")
    santiye_secimi_cb.pack(side="left", fill="x", expand=True)
    combobox_references.append(santiye_secimi_cb)
    ttk.Button(cb_frame, text="+", width=3, command=open_add_yard_dialog).pack(side="left", padx=(5,0))
    
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill="both", expand=True)
    
    # Sütunlar: code, mal_hizmet, miktar, birim
    tree = ttk.Treeview(tree_frame, columns=("code", "mal_hizmet", "miktar", "birim"), show="headings")
    tree.heading("code", text="Mal Kodu"); tree.column("code", width=150, anchor="w")
    tree.heading("mal_hizmet", text="Mal"); tree.column("mal_hizmet", width=450, anchor="w")
    tree.heading("miktar", text="Miktar"); tree.column("miktar", width=100, anchor="center")
    tree.heading("birim", text="Birim"); tree.column("birim", width=100, anchor="center")
    tree.tag_configure('oddrow', background='#F0F0F0')
    tree.tag_configure('evenrow', background='white')
    
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)

    # --- PDF OKUMA (ZRI FORMATI - Mal Kodu ve Miktar Arası Okuma) ---
    def pdf_oku_ve_doldur():
        dosya_yolu = filedialog.askopenfilename(title="PDF Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not dosya_yolu: return
        for i in tree.get_children(): tree.delete(i)
        
        try:
            with pdfplumber.open(dosya_yolu) as pdf:
                sayfa = pdf.pages[0]
                tablolar = sayfa.extract_tables()
                hedef_tablo = None
                
                kod_idx = -1
                miktar_idx = -1

                for tablo in tablolar:
                    if tablo:
                        basliklar = [str(b).replace('\n', ' ').strip() for b in tablo[0]]
                        if "Miktar" in basliklar:
                            miktar_idx = basliklar.index("Miktar")
                            if "Mal Kodu" in basliklar:
                                kod_idx = basliklar.index("Mal Kodu")
                                hedef_tablo = tablo
                                break
                            elif "Ürün Kodu" in basliklar:
                                kod_idx = basliklar.index("Ürün Kodu")
                                hedef_tablo = tablo
                                break
                
                if not hedef_tablo:
                    messagebox.showerror("Hata", "Uygun tablo bulunamadı. 'Mal Kodu' ve 'Miktar' sütunları gerekli.")
                    return

                for i, satir_data in enumerate(range(1, len(hedef_tablo))):
                    satir = hedef_tablo[satir_data]
                    if len(satir) <= max(kod_idx, miktar_idx): continue

                    mal_kodu = satir[kod_idx]
                    miktar_ham = satir[miktar_idx]

                    # Mal Kodu ile Miktar arasındaki her şeyi "Mal" olarak al
                    aradaki_sutunlar = satir[kod_idx+1 : miktar_idx]
                    mal_adi_listesi = [str(x).replace('\n', ' ').strip() for x in aradaki_sutunlar if x is not None and str(x).strip() != ""]
                    mal_adi = " ".join(mal_adi_listesi)

                    if mal_kodu and mal_adi and miktar_ham:
                        mal_kodu = str(mal_kodu).replace('\n', ' ').strip()
                        miktar_ham = str(miktar_ham).replace('\n', ' ').strip()
                        
                        eslesme = re.match(r"([\d.,]+)\s*([a-zA-Z]+)", miktar_ham)
                        sayisal_miktar, birim = ("", "")
                        
                        if eslesme:
                            sayisal_miktar = eslesme.group(1)
                            birim_kisaltma = eslesme.group(2)
                            if birim_kisaltma.upper() in ['M', 'MT', 'METRE']: birim = "METRE"
                            elif birim_kisaltma.upper() in ['AD', 'ADET']: birim = "ADET"
                            else: birim = birim_kisaltma.upper()
                        else: 
                            sayisal_miktar = miktar_ham
                        
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                        tree.insert("", "end", values=(mal_kodu, mal_adi, sayisal_miktar, birim), tags=(tag,))
        
        except Exception as e:
            messagebox.showerror("Hata", f"PDF işlenirken hata: {e}")

    # --- VERİ KAYDETME (GÜNCELLENDİ: HEM malHizmet HEM description DOLACAK) ---
    def verileri_kaydet_tek_tek():
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi: messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin."); return
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)
        if not tree.get_children(): messagebox.showwarning("Eksik Bilgi", "Kaydedilecek veri bulunmuyor."); return
        
        basarili, hatali = 0, 0
        toplam = len(tree.get_children())
        
        for row_id in tree.get_children():
            item = tree.item(row_id)['values']
            # item[0] -> code
            # item[1] -> mal (PDF'ten okunan)
            # item[2] -> amount
            # item[3] -> unit
            
            try:
                miktar_sayi = int(str(item[2]).replace('.', '').replace(',', ''))
            except ValueError: hatali += 1; continue
            
            # --- İŞTE BURASI DEĞİŞTİ ---
            # Java tarafında CreateProductRequest { code, malHizmet, description, ... } bekliyor.
            # Biz PDF'teki "Mal" verisini hem 'malHizmet'e hem de 'description'a gönderiyoruz.
            payload = {
                "code": str(item[0]),        
                "malHizmet": str(item[1]),   # Mal adı -> malHizmet'e
                "description": str(item[1]), # Mal adı -> description'a (Artık "-" değil)
                "amount": miktar_sayi,
                "unit": str(item[3]).upper()
            }
            
            try:
                url = f"{PRODUCTS_API_URL}/yards/{secili_santiye_id}/products"
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, data=json.dumps(payload), headers=headers)
                if response.status_code in [200, 201]: basarili += 1
                else: hatali += 1
            except requests.exceptions.RequestException: 
                messagebox.showerror("Bağlantı Hatası", "Sunucuya ulaşılamadı."); return
        
        messagebox.showinfo("Sonuç", f"Toplam: {toplam}\nBaşarılı: {basarili}\nHatalı: {hatali}")
        if basarili > 0: 
             for i in tree.get_children(): tree.delete(i)

    button_group = ttk.Frame(top_frame)
    button_group.grid(row=0, column=2, padx=(20, 0), pady=5, sticky="e")
    ttk.Button(button_group, text="PDF Yükle", command=pdf_oku_ve_doldur, style='Accent.TButton').pack(side="left", padx=(0, 5))
    ttk.Button(button_group, text="Verileri Kaydet", command=verileri_kaydet_tek_tek).pack(side="left")
    santiyeleri_getir(santiye_secimi_cb)

# --- SEKME 2: ŞANTİYE SORGULAMA ---

def create_sorgu_tab(tab):
    main_frame = ttk.Frame(tab, padding="10")
    main_frame.pack(fill="both", expand=True)
    top_frame = ttk.LabelFrame(main_frame, text="Kontroller", padding="10")
    top_frame.pack(fill="x", pady=(0, 10))
    top_frame.columnconfigure(1, weight=1)
    ttk.Label(top_frame, text="Görüntülenecek Şantiye:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")

    cb_frame = ttk.Frame(top_frame)
    cb_frame.grid(row=0, column=1, pady=5, sticky="ew")
    santiye_secimi_cb = ttk.Combobox(cb_frame, state="readonly")
    santiye_secimi_cb.pack(side="left", fill="x", expand=True)
    combobox_references.append(santiye_secimi_cb)
    ttk.Button(cb_frame, text="+", width=3, command=open_add_yard_dialog).pack(side="left", padx=(5,0))
    
    tree_frame = ttk.Frame(main_frame)
    tree_frame.pack(fill="both", expand=True)
    tree = ttk.Treeview(tree_frame, columns=("code", "mal_hizmet", "miktar", "birim"), show="headings")
    
    tree.heading("code", text="Mal Kodu")
    tree.column("code", width=150, anchor="w")
    
    tree.heading("mal_hizmet", text="Mal")
    tree.column("mal_hizmet", width=450, anchor="w")
    
    tree.heading("miktar", text="Miktar")
    tree.column("miktar", width=100, anchor="center")
    
    tree.heading("birim", text="Birim")
    tree.column("birim", width=100, anchor="center")
    
    tree.tag_configure('oddrow', background='#F0F0F0')
    tree.tag_configure('evenrow', background='white')
    
    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)

    def verileri_getir():
        for i in tree.get_children(): tree.delete(i)
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi: messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin."); return
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)
        try:
            url = f"{YARDS_API_URL}/{secili_santiye_id}"
            response = requests.get(url)
            if response.status_code == 200:
                veri = response.json()
                urunler = veri.get("products", []) 
                if not urunler: messagebox.showinfo("Bilgi", "Bu şantiyeye ait ürün bulunamadı.");
                for i, urun in enumerate(urunler):
                    tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                    p_code = urun.get('code', '')
                    p_mal = urun.get('malHizmet', '') 
                    p_amount = urun.get('amount', 0)
                    p_unit = urun.get('unit', '')
                    tree.insert("", "end", values=(p_code, p_mal, p_amount, p_unit), tags=(tag,))
            else:
                messagebox.showerror("Hata", f"Veriler alınamadı. Sunucu: {response.status_code}\n{response.text}")
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}")

    def verileri_excele_aktar():
        if not tree.get_children(): messagebox.showwarning("Veri Yok", "Aktarılacak veri bulunmuyor."); return
        try:
            dosya_yolu = filedialog.asksaveasfilename(
                initialfile=f'{santiye_secimi_cb.get()}_urun_listesi.xlsx',
                defaultextension=".xlsx", filetypes=[("Excel Dosyası", "*.xlsx")])
            if not dosya_yolu: return
            veri_listesi = []
            for row_id in tree.get_children():
                item = tree.item(row_id)['values']
                veri_listesi.append({'Mal Kodu': item[0], 'Mal': item[1], 'Miktar': item[2], 'Birim': item[3]})
            df = pd.DataFrame(veri_listesi)
            df.to_excel(dosya_yolu, index=False)
            messagebox.showinfo("Başarılı", f"Excel kaydedildi:\n{dosya_yolu}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel hatası:\n{e}")

    button_group = ttk.Frame(top_frame)
    button_group.grid(row=0, column=2, padx=(20, 0))
    ttk.Button(button_group, text="Verileri Getir", command=verileri_getir, style='Accent.TButton').pack(side="left", padx=(0, 5))
    ttk.Button(button_group, text="Excel'e Aktar", command=verileri_excele_aktar).pack(side="left")
    santiyeleri_getir(santiye_secimi_cb)

# --- ANA UYGULAMA ---
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