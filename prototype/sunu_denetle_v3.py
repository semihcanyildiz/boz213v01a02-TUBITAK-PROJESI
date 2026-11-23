"""Erişilebilir Sunu Denetleyicisi v3

Özellikler:
- Tkinter masaüstü arayüzü
- .pptx dosyası seçme ve dört kural ile denetleme
- Bulguları arayüzde görüntüleme
- Resim alt metni önerisini düzenleme ve yeni .pptx dosyasına uygulama
- Orijinal dosyayı koruma ve değişikliği yeniden açarak doğrulama
- CSV raporu oluşturma
- İsteğe bağlı Gemini alt metin önerisi (API anahtarı yoksa uygulama yine çalışır)

Kurulum:
    py -m pip install python-pptx pillow python-dotenv google-genai

İsteğe bağlı yapay zekâ önerisi için proje klasöründe .env dosyası:
    GEMINI_API_KEY=anahtarınız

Not: Bu prototip yalnızca kısmi otomatik denetim yapar; WCAG uyumluluk sertifikası vermez.
"""

import csv
import io
import os
import re
import sys
import traceback
from abc import ABC, abstractmethod
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    from google import genai
except ImportError:
    genai = None

API_ANAHTARI = os.getenv("GEMINI_API_KEY")
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


# ---------- Yardımcılar ----------
def luminans(rgb):
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def kontrast(on, arka):
    l1, l2 = sorted([luminans(on), luminans(arka)], reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def hex_rgb(v):
    return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))


def slayt_arka_plani(slayt):
    bg = slayt._element.find(f"{P}cSld/{P}bg")
    if bg is not None:
        renk = bg.find(f".//{A}srgbClr")
        if renk is not None:
            return hex_rgb(renk.get("val")), True
    return (255, 255, 255), False


def resim_alt_metin(shape):
    """Resim nesnesinin alternatif metnini döndürür."""
    cnv_list = shape._element.xpath(".//p:cNvPr")
    if not cnv_list:
        return ""
    return (cnv_list[0].get("descr") or "").strip()


# ---------- Denetim kuralları ----------
class Kural(ABC):
    kod, ad, seviye = "", "", ""
    neden, nasil = "", ""

    def __init__(self):
        self.notlar = []

    @abstractmethod
    def denetle(self, prs):
        raise NotImplementedError


class AltMetin(Kural):
    kod, ad, seviye = "1.1.1", "Metin dışı içerik (alt metin)", "A"
    neden = "Ekran okuyucu kullanan öğrenci resmi göremez; alt metin resmin anlamını aktarır."
    nasil = "Resmin alternatif metnini yazın. Yalnızca süs amaçlıysa dekoratif olarak işaretlenmesi değerlendirilebilir."

    def denetle(self, prs):
        bulgu, dekoratif = [], 0
        for no, slayt in enumerate(prs.slides, 1):
            for shape in slayt.shapes:
                if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    continue
                cnv_list = shape._element.xpath(".//p:cNvPr")
                if not cnv_list:
                    continue
                cnv = cnv_list[0]
                xml = etree.tostring(cnv).decode("utf-8")
                if "decorative" in xml and 'val="1"' in xml:
                    dekoratif += 1
                    continue
                descr = (cnv.get("descr") or "").strip()
                ad_nesne = cnv.get("name") or shape.name or "resim"
                base = {
                    "slayt": no,
                    "shape_id": shape.shape_id,
                    "nesne": ad_nesne,
                    "alt_metin_mevcut": descr,
                    "ayrinti": "",
                    "resim_blob": shape.image.blob,
                }
                if not descr:
                    base["ayrinti"] = "alt metin boş"
                    bulgu.append(base)
                elif re.search(r"\.(png|jpe?g|gif|bmp|tiff?)$", descr, re.I):
                    base["ayrinti"] = f"alt metin bir dosya adı: {descr}"
                    bulgu.append(base)
        self.notlar.append(f"Dekoratif işaretli resim: {dekoratif} (denetim dışı bırakıldı)")
        self.notlar.append("Yalnızca alt metnin varlığı/olası dosya adı denetlenir; anlam kalitesi değerlendirilmez.")
        return bulgu


class SlaytBasligi(Kural):
    kod, ad, seviye = "2.4.2", "Slayt başlığı", "A"
    neden = "Başlıklar, ekran okuyucu kullananların slaytlar arasında gezinmesini kolaylaştırır."
    nasil = "Slayta uygun bir başlık alanı ekleyip doldurun."

    def denetle(self, prs):
        bulgu = []
        for no, slayt in enumerate(prs.slides, 1):
            t = slayt.shapes.title
            if t is None:
                bulgu.append({"slayt": no, "shape_id": "", "nesne": "-", "ayrinti": "slaytta başlık alanı yok", "resim_blob": None})
            elif not t.text_frame.text.strip():
                bulgu.append({"slayt": no, "shape_id": t.shape_id, "nesne": t.name, "ayrinti": "başlık alanı boş", "resim_blob": None})
        self.notlar.append("Yalnızca PowerPoint başlık alanının varlığı/doluluğu kontrol edilir; normal metin kutuları başlık olarak tanınmayabilir.")
        return bulgu


class DilEtiketi(Kural):
    kod, ad, seviye = "3.1.1", "Dil etiketi", "A"
    neden = "Dil etiketi, ekran okuyucunun metni uygun telaffuzla okumasına yardımcı olur."
    nasil = "Metnin dilini uygun biçimde Türkçe olarak ayarlayın."

    def denetle(self, prs):
        bulgu, diller = [], set()
        for no, slayt in enumerate(prs.slides, 1):
            eksik = False
            for r in slayt._element.iter(A + "r"):
                t = r.find(A + "t")
                if t is None or not (t.text or "").strip():
                    continue
                rpr = r.find(A + "rPr")
                lang = rpr.get("lang") if rpr is not None else None
                diller.add(lang or "(yok)")
                if not lang or not lang.lower().startswith("tr"):
                    eksik = True
            if eksik:
                bulgu.append({"slayt": no, "shape_id": "", "nesne": "-", "ayrinti": "Türkçe dil etiketi olmayan veya açıkça tanımlanmamış metin var", "resim_blob": None})
        self.notlar.append("Metin parçalarındaki açık dil etiketi okunur; varsayılan stilden devralınan dil görülmeyebilir.")
        self.notlar.append("Bulunan dil etiketleri: " + (", ".join(sorted(diller)) or "metin yok"))
        return bulgu


class Kontrast(Kural):
    kod, ad, seviye = "1.4.3", "Kontrast (asgari)", "AA"
    neden = "Düşük kontrast metnin okunabilirliğini azaltabilir."
    nasil = "Normal metinde en az 4,5:1; büyük metinde en az 3:1 kontrast hedefleyin."

    def denetle(self, prs):
        bulgu, olculemedi, varsayim = [], 0, False
        for no, slayt in enumerate(prs.slides, 1):
            arka, acik = slayt_arka_plani(slayt)
            varsayim = varsayim or not acik
            for r in slayt._element.iter(A + "r"):
                t = r.find(A + "t")
                if t is None or not (t.text or "").strip():
                    continue
                rpr = r.find(A + "rPr")
                renk = rpr.find(f"{A}solidFill/{A}srgbClr") if rpr is not None else None
                if renk is None:
                    olculemedi += 1
                    continue
                try:
                    rgb = hex_rgb(renk.get("val"))
                    boyut = int(rpr.get("sz", "1800")) / 100
                except (ValueError, TypeError):
                    olculemedi += 1
                    continue
                buyuk = boyut >= 18 or (boyut >= 14 and rpr.get("b") == "1")
                esik = 3.0 if buyuk else 4.5
                oran = kontrast(rgb, arka)
                if oran < esik:
                    parca = (t.text or "").strip()[:30]
                    bulgu.append({"slayt": no, "shape_id": "", "nesne": f'"{parca}"', "ayrinti": f"kontrast {oran:.2f}:1, gereken {esik}:1 ({boyut:g} punto)", "resim_blob": None})
        self.notlar.append(f"Rengi tema/devralma olduğu için ölçülemeyen metin parçaları: {olculemedi}")
        if varsayim:
            self.notlar.append("Açık arka plan rengi olmayan slaytlarda arka plan beyaz varsayıldı.")
        self.notlar.append("Bu prototip tema renklerini, gradyanları, resim üstündeki metinleri ve tüm devralınan biçimlendirmeleri çözümlemez.")
        return bulgu


# ---------- Öneri katmanı ----------
class OneriUretici(ABC):
    @abstractmethod
    def oner(self, kural, bulgu_maddesi):
        raise NotImplementedError


class YapayZekaOneriUretici(OneriUretici):
    def __init__(self):
        if genai is None:
            raise RuntimeError("google-genai kurulu değil. İsteğe bağlı öneri için: py -m pip install google-genai")
        if not API_ANAHTARI:
            raise RuntimeError("GEMINI_API_KEY bulunamadı. .env dosyasına ekleyin veya öneriyi elle yazın.")
        self.client = genai.Client(api_key=API_ANAHTARI)
        # Kullanılabilir model adını kendi Google AI Studio/API hesabınızda doğrulayın.
        self.model_adi = os.getenv("GEMINI_MODEL", "models/gemini-3.6-flash")

    def oner(self, kural, bulgu_maddesi):
        if not isinstance(kural, AltMetin) or not bulgu_maddesi.get("resim_blob"):
            return ""
        try:
            resim = Image.open(io.BytesIO(bulgu_maddesi["resim_blob"]))
            prompt = (
                "Bu görseli Türkçe olarak, ekran okuyucu kullanan bir öğrenci için "
                "kısa ve nesnel bir alternatif metin cümlesiyle betimle. Görseldeki "
                "bilgiyi uydurma; emin olmadığın ayrıntıları belirtme. Yalnızca alt metni yaz."
            )
            yanit = self.client.models.generate_content(
                model=self.model_adi,
                contents=[prompt, resim],
            )
            return (yanit.text or "").strip()
        except Exception as exc:
            raise RuntimeError(f"Yapay zekâ önerisi alınamadı: {exc}") from exc


# ---------- Denetim ve CSV ----------
def denetle_dosya(yol):
    prs = Presentation(yol)
    kurallar = [AltMetin(), SlaytBasligi(), DilEtiketi(), Kontrast()]
    sonuclar = []
    for kural in kurallar:
        bulgular = kural.denetle(prs)
        for bulgu in bulgular:
            sonuclar.append({
                "olcut": kural.kod,
                "olcut_adi": kural.ad,
                "seviye": kural.seviye,
                "neden": kural.neden,
                "nasil": kural.nasil,
                **bulgu,
            })
    notlar = [(k.kod, k.ad, n) for k in kurallar for n in k.notlar]
    return prs, sonuclar, notlar


def csv_yaz(yol, sonuclar):
    hedef = Path(yol).with_name(Path(yol).stem + "_rapor.csv")
    alanlar = ["olcut", "olcut_adi", "seviye", "slayt", "shape_id", "nesne", "ayrinti", "alt_metin_mevcut", "neden", "nasil"]
    with open(hedef, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=alanlar, extrasaction="ignore")
        writer.writeheader()
        for satir in sonuclar:
            writer.writerow(satir)
    return str(hedef)


def alt_metin_uygula(kaynak_dosya, hedef_dosya, slayt_no, shape_id, alt_metin):
    """Alt metni resme yazar; kaynak dosyaya dokunmadan yeni .pptx oluşturur."""
    kaynak = Path(kaynak_dosya).expanduser().resolve()
    hedef = Path(hedef_dosya).expanduser().resolve()
    metin = (alt_metin or "").strip()

    if not kaynak.is_file():
        raise FileNotFoundError(f"Kaynak dosya bulunamadı: {kaynak}")
    if kaynak.suffix.lower() != ".pptx" or hedef.suffix.lower() != ".pptx":
        raise ValueError("Kaynak ve hedef dosyalar .pptx olmalıdır.")
    if kaynak == hedef:
        raise ValueError("Orijinal dosyanın üzerine yazılamaz.")
    if hedef.exists():
        raise FileExistsError(f"Hedef dosya zaten var: {hedef}\nBaşka bir dosya adı seçin.")
    if not metin:
        raise ValueError("Alt metin boş bırakılamaz.")
    if len(metin) > 2000:
        raise ValueError("Alt metin çok uzun; lütfen kısaltın.")

    prs = Presentation(str(kaynak))
    if not isinstance(slayt_no, int) or not 1 <= slayt_no <= len(prs.slides):
        raise IndexError("Geçersiz slayt numarası.")
    slayt = prs.slides[slayt_no - 1]
    shape = next((s for s in slayt.shapes if s.shape_id == int(shape_id)), None)
    if shape is None:
        raise LookupError("Seçili slaytta bu nesne kimliği bulunamadı; dosyayı yeniden denetleyin.")
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        raise TypeError("Bu sürüm yalnızca doğrudan resim nesnelerine alt metin yazar.")

    cnv_list = shape._element.xpath(".//p:cNvPr")
    if not cnv_list:
        raise RuntimeError("Resmin alternatif metin alanı bulunamadı.")
    cnv_list[0].set("descr", metin)

    hedef.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(hedef))

    # Kaydedilen dosyayı yeniden açarak yazımı doğrula.
    check_prs = Presentation(str(hedef))
    check_shape = next((s for s in check_prs.slides[slayt_no - 1].shapes if s.shape_id == int(shape_id)), None)
    if check_shape is None:
        raise RuntimeError("Kaydedilen dosyada ilgili resim bulunamadı.")
    check_cnv = check_shape._element.xpath(".//p:cNvPr")
    if not check_cnv or check_cnv[0].get("descr") != metin:
        raise RuntimeError("Alt metin kaydedilen dosyada doğrulanamadı.")
    return str(hedef)


# ---------- Tkinter arayüzü ----------
class SunuUygulamasi:
    def __init__(self, root):
        self.root = root
        self.root.title("Erişilebilir Sunu Denetleyicisi")
        self.root.geometry("1120x760")
        self.root.minsize(900, 620)
        self.kaynak_dosya = ""
        self.sonuclar = []
        self.notlar = []
        self.secilen_indeks = None
        self._arayuz_kur()

    def _arayuz_kur(self):
        ana = ttk.Frame(self.root, padding=14)
        ana.pack(fill="both", expand=True)

        ust = ttk.Frame(ana)
        ust.pack(fill="x")
        ttk.Label(ust, text="Erişilebilir Sunu Denetleyicisi", font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Button(ust, text="PowerPoint seç", command=self.dosya_sec).pack(side="right")

        self.dosya_etiketi = ttk.Label(ana, text="Henüz dosya seçilmedi", foreground="#555555")
        self.dosya_etiketi.pack(fill="x", pady=(8, 12))

        eylemler = ttk.Frame(ana)
        eylemler.pack(fill="x", pady=(0, 10))
        ttk.Button(eylemler, text="Denetle", command=self.denetle).pack(side="left")
        ttk.Button(eylemler, text="CSV raporunu aç/kaydet", command=self.csv_kaydet).pack(side="left", padx=8)
        self.ozet_etiketi = ttk.Label(eylemler, text="Denetim bekleniyor")
        self.ozet_etiketi.pack(side="right")

        govde = ttk.Panedwindow(ana, orient="horizontal")
        govde.pack(fill="both", expand=True)
        sol = ttk.Frame(govde, padding=(0, 0, 8, 0))
        sag = ttk.Frame(govde, padding=(8, 0, 0, 0))
        govde.add(sol, weight=3)
        govde.add(sag, weight=2)

        ttk.Label(sol, text="Denetim bulguları", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        kolonlar = ("olcut", "slayt", "nesne", "ayrinti")
        self.tablo = ttk.Treeview(sol, columns=kolonlar, show="headings", selectmode="browse")
        for col, baslik, genislik in [
            ("olcut", "Ölçüt", 75), ("slayt", "Slayt", 55),
            ("nesne", "Nesne", 135), ("ayrinti", "Bulgu", 300),
        ]:
            self.tablo.heading(col, text=baslik)
            self.tablo.column(col, width=genislik, anchor="w")
        ybar = ttk.Scrollbar(sol, orient="vertical", command=self.tablo.yview)
        self.tablo.configure(yscrollcommand=ybar.set)
        self.tablo.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        self.tablo.bind("<<TreeviewSelect>>", self.bulgu_secildi)

        ttk.Label(sag, text="Bulgu ayrıntısı", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 6))
        self.ayrinti = tk.Text(sag, height=9, wrap="word", state="disabled", font=("Segoe UI", 10))
        self.ayrinti.pack(fill="x", pady=(0, 10))

        ttk.Label(sag, text="Alt metin / düzenleme alanı", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.alt_metin = tk.Text(sag, height=5, wrap="word", font=("Segoe UI", 10))
        self.alt_metin.pack(fill="x", pady=(4, 8))

        btnlar = ttk.Frame(sag)
        btnlar.pack(fill="x")
        ttk.Button(btnlar, text="ÜYZ ile öner", command=self.yapay_zeka_oner).pack(side="left")
        ttk.Button(btnlar, text="Alt metni yeni sunuya uygula", command=self.alt_metin_kaydet).pack(side="left", padx=(6, 0))

        ttk.Label(sag, text="İşlem günlüğü", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(14, 4))
        self.gunluk = tk.Text(sag, height=8, wrap="word", state="disabled", font=("Consolas", 9))
        self.gunluk.pack(fill="both", expand=True)
        self._log("Uygulama hazır. Bir .pptx dosyası seçip Denetle'ye basın.")

    def _log(self, mesaj):
        self.gunluk.configure(state="normal")
        self.gunluk.insert("end", mesaj + "\n")
        self.gunluk.see("end")
        self.gunluk.configure(state="disabled")

    def dosya_sec(self):
        yol = filedialog.askopenfilename(title="PowerPoint dosyası seçin", filetypes=[("PowerPoint sunusu", "*.pptx")])
        if yol:
            self.kaynak_dosya = yol
            self.dosya_etiketi.configure(text=yol)
            self._log(f"Dosya seçildi: {yol}")

    def denetle(self):
        if not self.kaynak_dosya:
            messagebox.showwarning("Dosya gerekli", "Önce bir .pptx dosyası seçin.")
            return
        try:
            _, self.sonuclar, self.notlar = denetle_dosya(self.kaynak_dosya)
            for item in self.tablo.get_children():
                self.tablo.delete(item)
            for i, b in enumerate(self.sonuclar):
                self.tablo.insert("", "end", iid=str(i), values=(b["olcut"], b["slayt"], b["nesne"], b["ayrinti"]))
            self.ozet_etiketi.configure(text=f"{len(self.sonuclar)} bulgu · {len(Presentation(self.kaynak_dosya).slides)} slayt")
            self._log(f"Denetim tamamlandı: {len(self.sonuclar)} bulgu.")
            for kod, ad, not_metni in self.notlar:
                self._log(f"{kod} / {ad}: {not_metni}")
            if not self.sonuclar:
                messagebox.showinfo("Denetim tamamlandı", "Bu prototipin denetlediği kurallarda bulgu bulunmadı. Bu sonuç WCAG uyumluluğu anlamına gelmez.")
        except Exception as exc:
            self._log(traceback.format_exc())
            messagebox.showerror("Denetim hatası", str(exc))

    def csv_kaydet(self):
        if not self.kaynak_dosya or not self.sonuclar:
            messagebox.showinfo("CSV raporu", "Önce dosyayı denetleyin ve en az bir bulgu oluşturun.")
            return
        try:
            yol = csv_yaz(self.kaynak_dosya, self.sonuclar)
            self._log(f"CSV raporu kaydedildi: {yol}")
            messagebox.showinfo("CSV raporu kaydedildi", yol)
        except Exception as exc:
            messagebox.showerror("CSV hatası", str(exc))

    def bulgu_secildi(self, _event=None):
        secim = self.tablo.selection()
        if not secim:
            return
        self.secilen_indeks = int(secim[0])
        b = self.sonuclar[self.secilen_indeks]
        metin = (
            f"Ölçüt: {b['olcut']} — {b['olcut_adi']} (Seviye {b['seviye']})\n"
            f"Slayt: {b['slayt']}\nNesne: {b['nesne']}\n"
            f"Bulgu: {b['ayrinti']}\n\nNeden önemli?\n{b.get('neden', '')}\n\nNasıl düzeltilir?\n{b.get('nasil', '')}"
        )
        self.ayrinti.configure(state="normal")
        self.ayrinti.delete("1.0", "end")
        self.ayrinti.insert("1.0", metin)
        self.ayrinti.configure(state="disabled")
        self.alt_metin.delete("1.0", "end")
        self.alt_metin.insert("1.0", b.get("alt_metin_mevcut", ""))
        if b["olcut"] != "1.1.1" or not b.get("shape_id"):
            self._log("Seçilen bulgu alt metin düzenleme işlemi için uygun bir doğrudan resim bulgusu değil.")

    def yapay_zeka_oner(self):
        if self.secilen_indeks is None:
            messagebox.showinfo("Bulgu seçin", "Önce soldaki listeden bir bulgu seçin.")
            return
        b = self.sonuclar[self.secilen_indeks]
        if b["olcut"] != "1.1.1" or not b.get("resim_blob"):
            messagebox.showinfo("Öneri kullanılamıyor", "Yapay zekâ önerisi bu prototipte yalnızca resim alt metni için kullanılabilir.")
            return
        try:
            uretici = YapayZekaOneriUretici()
            onerilen = uretici.oner(AltMetin(), b)
            if onerilen:
                self.alt_metin.delete("1.0", "end")
                self.alt_metin.insert("1.0", onerilen)
                self._log("Yapay zekâ önerisi metin alanına getirildi; kaydetmeden önce kontrol edin.")
            else:
                messagebox.showinfo("Öneri yok", "Bu bulgu için öneri oluşturulamadı.")
        except Exception as exc:
            messagebox.showerror("Yapay zekâ önerisi", str(exc))
            self._log(str(exc))

    def alt_metin_kaydet(self):
        if self.secilen_indeks is None:
            messagebox.showinfo("Bulgu seçin", "Önce soldaki listeden alt metin bulgusu seçin.")
            return
        b = self.sonuclar[self.secilen_indeks]
        if b["olcut"] != "1.1.1" or not b.get("shape_id"):
            messagebox.showwarning("Desteklenmeyen bulgu", "Bu işlem yalnızca kimliği belirlenmiş doğrudan resim bulgularına uygulanabilir.")
            return
        metin = self.alt_metin.get("1.0", "end").strip()
        hedef = filedialog.asksaveasfilename(
            title="Düzeltilmiş sunuyu yeni dosya olarak kaydet",
            defaultextension=".pptx",
            filetypes=[("PowerPoint sunusu", "*.pptx")],
            initialfile=Path(self.kaynak_dosya).stem + "_duzeltilmis.pptx",
        )
        if not hedef:
            return
        try:
            kaydedilen = alt_metin_uygula(
                self.kaynak_dosya, hedef, int(b["slayt"]), int(b["shape_id"]), metin
            )
            self._log(f"Alt metin kaydedildi ve doğrulandı: {kaydedilen}")
            messagebox.showinfo("Başarılı", f"Alt metin yeni dosyaya kaydedildi ve doğrulandı.\n\n{kaydedilen}\n\nOrijinal dosya değiştirilmedi.")
        except Exception as exc:
            messagebox.showerror("Düzeltme yapılamadı", str(exc))
            self._log(str(exc))


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except tk.TclError:
        pass
    SunuUygulamasi(root)
    root.mainloop()


if __name__ == "__main__":
    # GUI sürümü varsayılan olarak açılır. CLI ile otomatik rapor gerekirse:
    # py sunu_denetle_v3.py --rapor dosya.pptx
    if len(sys.argv) >= 3 and sys.argv[1] == "--rapor":
        dosya = sys.argv[2]
        if not os.path.isfile(dosya):
            print("Dosya bulunamadı:", dosya)
            raise SystemExit(1)
        _, bulgular, notlar = denetle_dosya(dosya)
        for b in bulgular:
            print(f"{b['olcut']} | Slayt {b['slayt']} | {b['nesne']} | {b['ayrinti']}")
        for kod, ad, not_metni in notlar:
            print(f"- {kod} {ad}: {not_metni}")
        print("CSV:", csv_yaz(dosya, bulgular))
    else:
        main()
