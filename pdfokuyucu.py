import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pdfplumber
import os
import re
import requests
import json
import pandas as pd

# --- Backend Ayarları ---
BASE_URL = "http://localhost:8080"
YARDS_API_URL = f"{BASE_URL}/api/yards"
PRODUCTS_API_URL = f"{BASE_URL}/api/products"
santiyeler_map = {}

# --- GENEL FONKSİYONLAR ---

def santiyeleri_getir(combobox_widget):
    """
    Backend'den şantiye listesini çeker ve verilen combobox'ı doldurur.
    """
    try:
        response = requests.get(YARDS_API_URL)
        if response.status_code == 200:
            santiyeler = response.json()
            santiye_isimleri = []
            santiyeler_map.clear()
            for santiye in santiyeler:
                isim = santiye.get("yardName")
                santiye_id = santiye.get("id")
                if isim and santiye_id is not None:
                    santiye_isimleri.append(isim)
                    santiyeler_map[isim] = santiye_id
            
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

# --- FATURA GİRİŞ PENCERESİ ---

def fatura_penceresini_ac():
    fatura_win = tk.Toplevel(root)
    fatura_win.title("Faturadan Veri Aktar")
    fatura_win.geometry("900x700")

    selection_frame = tk.Frame(fatura_win)
    selection_frame.pack(pady=10, padx=20, fill="x")
    santiye_label = tk.Label(selection_frame, text="Kaydedilecek Şantiye:", font=("Helvetica", 10))
    santiye_label.pack(side="left", padx=(0, 10))
    santiye_secimi_cb = ttk.Combobox(selection_frame, state="readonly", font=("Helvetica", 10))
    santiye_secimi_cb.pack(side="left", fill="x", expand=True)

    tree = ttk.Treeview(fatura_win, columns=("h_kod", "h_aciklama", "miktar", "birim"), show="headings")
    tree.heading("h_kod", text="Hizmet Kodu"); tree.column("h_kod", width=150)
    tree.heading("h_aciklama", text="Hizmet Açıklaması"); tree.column("h_aciklama", width=420)
    tree.heading("miktar", text="Miktar"); tree.column("miktar", width=100)
    tree.heading("birim", text="Birim"); tree.column("birim", width=80)
    tree.pack(pady=10, padx=20, fill="both", expand=True)
    
    def pdf_oku_ve_doldur():
        dosya_yolu = filedialog.askopenfilename(title="Fatura PDF'ini Seçin", filetypes=[("PDF Dosyaları", "*.pdf")])
        if not dosya_yolu: return
        
        fatura_win.title(f"Veri Aktar - {os.path.basename(dosya_yolu)}")
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
                    messagebox.showerror("Hata", "Fatura kalemlerini içeren ana tablo bulunamadı.", parent=fatura_win)
                    return
                for satir in range(1, len(hedef_tablo)):
                    hizmet_kodu = hedef_tablo[satir][1]
                    hizmet_aciklamasi = hedef_tablo[satir][2]
                    miktar_ham = hedef_tablo[satir][4]
                    if hizmet_kodu and hizmet_aciklamasi and miktar_ham:
                        hizmet_aciklamasi = hizmet_aciklamasi.replace('\n', ' ')
                        miktar_ham = miktar_ham.replace('\n', ' ').strip()
                        eslesme = re.match(r"([\d.,]+)\s*([a-zA-Z]+)", miktar_ham)
                        sayisal_miktar, birim = ("", "")
                        if eslesme:
                            sayisal_miktar = eslesme.group(1)
                            birim_kisaltma = eslesme.group(2)
                            if birim_kisaltma.upper() == 'M': birim = "Metre"
                            elif birim_kisaltma.upper() == 'ADET': birim = "Adet"
                            else: birim = birim_kisaltma
                        else:
                            sayisal_miktar = miktar_ham
                        tree.insert("", "end", values=(hizmet_kodu, hizmet_aciklamasi, sayisal_miktar, birim))
        except Exception as e:
            messagebox.showerror("Hata", f"PDF işlenirken bir hata oluştu: {e}", parent=fatura_win)

    def verileri_kaydet_tek_tek():
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi:
            messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin.", parent=fatura_win)
            return
        
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)

        if not tree.get_children():
            messagebox.showwarning("Eksik Bilgi", "Kaydedilecek veri bulunmuyor.", parent=fatura_win)
            return
            
        basarili_kayit, hatali_kayit = 0, 0
        toplam_kayit = len(tree.get_children())

        for row_id in tree.get_children():
            item = tree.item(row_id)['values']
            try:
                miktar_sayi = int(str(item[2]).replace('.', '').replace(',', ''))
            except ValueError:
                messagebox.showerror("Veri Hatası", f"'{item[2]}' miktarı geçerli bir sayı değil. Bu satır atlanacak.", parent=fatura_win)
                hatali_kayit += 1
                continue

            birim_str = str(item[3]).upper()
            birim_enum = "ADET"
            if birim_str == "METRE": birim_enum = "METRE"
            elif birim_str == "ADET": birim_enum = "ADET"
            
            gonderilecek_tek_urun = {
                "code": item[0], "description": item[1],
                "amount": miktar_sayi, "unit": birim_enum
            }
            try:
                url = f"{PRODUCTS_API_URL}/yards/{secili_santiye_id}/products"
                headers = {'Content-Type': 'application/json'}
                response = requests.post(url, data=json.dumps(gonderilecek_tek_urun), headers=headers)
                if response.status_code == 200 or response.status_code == 201:
                    basarili_kayit += 1
                else:
                    hatali_kayit += 1
                    if hatali_kayit == 1:
                         messagebox.showerror("Kayıt Hatası", f"Ürün '{item[0]}' kaydedilemedi. Sunucu: {response.status_code}\n{response.text}", parent=fatura_win)
            except requests.exceptions.RequestException as e:
                messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}", parent=fatura_win)
                return
        
        messagebox.showinfo("İşlem Tamamlandı", f"Toplam {toplam_kayit} üründen:\nBaşarıyla kaydedilen: {basarili_kayit}\nHatalı: {hatali_kayit}", parent=fatura_win)
        
        if basarili_kayit > 0:
            for i in tree.get_children():
                tree.delete(i)

    button_frame = tk.Frame(fatura_win)
    button_frame.pack(pady=10, padx=20, fill="x")
    load_button = tk.Button(button_frame, text="1. PDF Yükle", command=pdf_oku_ve_doldur, font=("Helvetica", 12))
    load_button.pack(side="left", fill="x", expand=True, padx=(0, 5))
    save_button = tk.Button(button_frame, text="2. Verileri Kaydet", command=verileri_kaydet_tek_tek, font=("Helvetica", 12))
    save_button.pack(side="left", fill="x", expand=True, padx=(5, 0))

    if not santiyeleri_getir(santiye_secimi_cb):
        fatura_win.destroy()

# --- ŞANTİYE SORGULAMA PENCERESİ ---
def sorgu_penceresini_ac():
    sorgu_win = tk.Toplevel(root)
    sorgu_win.title("Şantiye Verilerini Görüntüle")
    sorgu_win.geometry("900x700")

    selection_frame = tk.Frame(sorgu_win)
    selection_frame.pack(pady=10, padx=20, fill="x")
    santiye_label = tk.Label(selection_frame, text="Görüntülenecek Şantiye:", font=("Helvetica", 10))
    santiye_label.pack(side="left", padx=(0, 10))
    santiye_secimi_cb = ttk.Combobox(selection_frame, state="readonly", font=("Helvetica", 10))
    santiye_secimi_cb.pack(side="left", fill="x", expand=True)

    tree = ttk.Treeview(sorgu_win, columns=("code", "description", "amount", "unit"), show="headings")
    tree.heading("code", text="Ürün Kodu"); tree.column("code", width=150)
    tree.heading("description", text="Açıklama"); tree.column("description", width=420)
    tree.heading("amount", text="Miktar"); tree.column("amount", width=100)
    tree.heading("unit", text="Birim"); tree.column("unit", width=80)
    tree.pack(pady=10, padx=20, fill="both", expand=True)

    def verileri_getir():
        for i in tree.get_children(): tree.delete(i)
        secili_santiye_ismi = santiye_secimi_cb.get()
        if not secili_santiye_ismi:
            messagebox.showwarning("Eksik Bilgi", "Lütfen bir şantiye seçin.", parent=sorgu_win)
            return
        secili_santiye_id = santiyeler_map.get(secili_santiye_ismi)
        try:
            url = f"{YARDS_API_URL}/{secili_santiye_id}"
            response = requests.get(url)
            if response.status_code == 200:
                veri = response.json()
                urunler = veri.get("products", []) 
                if not urunler:
                    messagebox.showinfo("Bilgi", "Bu şantiyeye ait kayıtlı ürün bulunamadı.", parent=sorgu_win)
                for urun in urunler:
                    tree.insert("", "end", values=(urun['code'], urun['description'], urun['amount'], urun['unit']))
            else:
                messagebox.showerror("Hata", f"Veriler alınamadı. Sunucu: {response.status_code}\n{response.text}", parent=sorgu_win)
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Bağlantı Hatası", f"Sunucuya bağlanılamadı: {e}", parent=sorgu_win)

    # --- YENİ VE İÇİ DOLDURULMUŞ FONKSİYON: EXCEL'E AKTARMA ---
    def verileri_excele_aktar():
        if not tree.get_children():
            messagebox.showwarning("Veri Yok", "Aktarılacak veri bulunmuyor. Lütfen önce verileri getirin.", parent=sorgu_win)
            return
        try:
            dosya_yolu = filedialog.asksaveasfilename(
                initialfile=f'{santiye_secimi_cb.get()}_urun_listesi.xlsx',
                defaultextension=".xlsx",
                filetypes=[("Excel Dosyası", "*.xlsx"), ("Tüm Dosyalar", "*.*")])
            if not dosya_yolu:
                return
            
            veri_listesi = []
            for row_id in tree.get_children():
                item = tree.item(row_id)['values']
                veri_listesi.append({
                    'Ürün Kodu': item[0],
                    'Açıklama': item[1],
                    'Miktar': item[2],
                    'Birim': item[3]
                })

            df = pd.DataFrame(veri_listesi)
            df.to_excel(dosya_yolu, index=False)
            messagebox.showinfo("Başarılı", f"Veriler başarıyla şu dosyaya aktarıldı:\n{dosya_yolu}", parent=sorgu_win)
        except Exception as e:
            messagebox.showerror("Hata", f"Excel'e aktarılırken bir hata oluştu:\n{e}", parent=sorgu_win)

    # Butonları içeren çerçeve
    button_frame = tk.Frame(selection_frame)
    button_frame.pack(side="left", padx=(10, 0))

    get_button = tk.Button(button_frame, text="Verileri Getir", command=verileri_getir, font=("Helvetica", 10))
    get_button.pack(side="left")

    # --- GÜNCELLENMİŞ BUTON: EXCEL'E AKTAR ---
    # Butonun komutu artık doğrudan 'verileri_excele_aktar' fonksiyonunu çağırıyor.
    excel_button = tk.Button(button_frame, text="Excel'e Aktar", command=verileri_excele_aktar, font=("Helvetica", 10))
    excel_button.pack(side="left", padx=(5, 0))

    if not santiyeleri_getir(santiye_secimi_cb):
        sorgu_win.destroy()

# --- ANA PENCERE (MENÜ) ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Ana Menü")
    root.geometry("400x200")

    main_label = tk.Label(root, text="Lütfen yapmak istediğiniz işlemi seçin:", font=("Helvetica", 12))
    main_label.pack(pady=20)

    fatura_btn = tk.Button(root, text="Faturadan Veri Aktar", command=fatura_penceresini_ac, font=("Helvetica", 12), height=2)
    fatura_btn.pack(pady=5, padx=20, fill="x")

    sorgu_btn = tk.Button(root, text="Şantiye Verilerini Görüntüle", command=sorgu_penceresini_ac, font=("Helvetica", 12), height=2)
    sorgu_btn.pack(pady=5, padx=20, fill="x")

    root.mainloop()