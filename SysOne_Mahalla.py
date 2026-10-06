# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════╗
║   SysOne — Avto-yordamchi  (v10)                  ║
║   SysOne Digital Solutions                                       ║
╚══════════════════════════════════════════════════════════════════╝

v9 -> v10 — O'ZGARISHLAR:

  1) AVTOR / BREND: terminal boshida SysOne ASCII-banner + muallif bloki
     (rangli, colorama yoki ANSI orqali).

  2) TASNIF ANIQLIGI (7 ta real ekran asosida qattiqlashtirildi):
     • «Бу фуқарони кўчириб бўлмайди»             → KO'CHIRIB BO'LMAYDI (skip)
     • «...рўйхатида мавжуд» + «қайта қўшиш...»   → JORIY MFY MAVJUD (skip)
     • «...ҳудудда топилмади»                      → HUDUDDA TOPILMADI (skip)
     • «Давом этиш» tugmasi                        → BOSHQA MFY (kutish)
     • «РЕЕСТРДАН» / «сиз тўлдирасиз» / forma       → FORMA TAYYOR (kutish)
     • «Манзил текширилмоқда...»                   → TEKSHIRILMOQDA (kutadi)
     • Toast «Фуқаро маълумоти топилмади»          → QAYTA QIDIRISH (max 6)

  3) Rasm 3 (tekshirilmoqda) endi rasm 7 (joriy MFY) bilan
     chalkashmaydi — «қайта қўшиш мумкин эмас» eng aniq belgi.

  4) Toast aniqlansa — darhol qayta-qidirish (vaqt isrof bo'lmaydi).

  BOSHQARUV (faol):
     F7 → Pauza/Davom | Shift+1 → sekin | Shift+2 → tez
     Har natija .xlsx + .csv ga yoziladi. Holat brauzer panelida ko'rinadi.

  ESLATMA: bot ma'lumotni O'ZI to'ldirmaydi/saqlamaydi. Topilgan holatda
  formani SIZ to'ldirib «Сақлаш» bosasiz — bot oyna yopilguncha kutadi.
  Bu real davlat reestriga soxta yozuv tushmasligi uchun.
"""

import os
import csv
import time
import threading
from dataclasses import dataclass
from enum import Enum

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

try:
    from pynput import keyboard as _kb
    PYNPUT_BOR = True
except Exception:
    PYNPUT_BOR = False

# Rang — colorama bo'lsa yaxshiroq, bo'lmasa ANSI (Win10+ qo'llab-quvvatlaydi)
try:
    import colorama
    colorama.init()
except Exception:
    pass

AVTOR = {
    "brend":    "SysOne Digital Solutions",
    "mahsulot": "Mahalla.ijro.uz Avto-yordamchi",
    "versiya":  "v10.0",
    "muallif":  "Qobilbek Odilov",
    "yil":      "2026",
    "aloqa":    "SysOne  •  Uchko'prik, Farg'ona",
}

class C:
    R   = "\033[0m"
    B   = "\033[1m"
    DIM = "\033[2m"
    CYAN   = "\033[96m"
    BLUE   = "\033[94m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    MAG    = "\033[95m"
    GREY   = "\033[90m"
    WHITE  = "\033[97m"


def banner():
    art = r"""
   ███████╗██╗   ██╗███████╗ ██████╗ ███╗   ██╗███████╗
   ██╔════╝╚██╗ ██╔╝██╔════╝██╔═══██╗████╗  ██║██╔════╝
   ███████╗ ╚████╔╝ ███████╗██║   ██║██╔██╗ ██║█████╗
   ╚════██║  ╚██╔╝  ╚════██║██║   ██║██║╚██╗██║██╔══╝
   ███████║   ██║   ███████║╚██████╔╝██║ ╚████║███████╗
   ╚══════╝   ╚═╝   ╚══════╝ ╚═════╝ ╚═╝  ╚═══╝╚══════╝
"""
    line = "═" * 60
    print(f"{C.CYAN}{C.B}{art}{C.R}")
    print(f"{C.BLUE}   {line}{C.R}")
    print(f"   {C.WHITE}{C.B}{AVTOR['mahsulot']}{C.R}  "
          f"{C.GREEN}{AVTOR['versiya']}{C.R}")
    print(f"   {C.GREY}Brend   :{C.R} {C.WHITE}{AVTOR['brend']}{C.R}")
    print(f"   {C.GREY}Muallif :{C.R} {C.WHITE}{AVTOR['muallif']}{C.R}  "
          f"{C.GREY}({AVTOR['yil']}){C.R}")
    print(f"   {C.GREY}Aloqa   :{C.R} {C.DIM}{AVTOR['aloqa']}{C.R}")
    print(f"{C.BLUE}   {line}{C.R}")
    print(f"   {C.YELLOW}F7{C.R} Pauza/Davom   "
          f"{C.YELLOW}Shift+1{C.R} sekin   "
          f"{C.YELLOW}Shift+2{C.R} tez")
    print(f"{C.BLUE}   {line}{C.R}\n")


# SOZLAMALAR
EXCEL_PATH    = "ro'yhat.xlsx"
NATIJA_LOG    = "natijalar_hisoboti.csv"
NATIJA_XLSX   = "natijalar_hisoboti.xlsx"
PLATFORMA_URL = (
    "https://mahalla.ijro.uz/dashboard/list/population"
    "?region_id=00s0eed0000region000012"
    "&district_id=00s0eed0000region000125"
    "&mahalla_id=66016a6a00016f2f91b2b904"
    "&returnPath=%2Fdashboard%2Fassistance"
)

WAIT_SEC     = 12
POLL         = 0.1
ANIQLASH_KUT = 6
ANIQLASH_CAP = 30

QIDIRISH_QAYTA_MAX     = 6
MANUAL_KUTISH_MAX_SOAT = 6
ESLATMA_ORALIGI_SEC    = 60

TEZLIK_BAZA        = 100
TEZLIK_BOSHLANGICH = 97
TEZLIK_QADAM       = 10
TEZLIK_MIN         = 30
TEZLIK_MAX         = 200


# BOSHQARUV (PAUZA + TEZLIK)
class Boshqaruv:
    def __init__(self):
        self.pauza = False
        self.tezlik = TEZLIK_BOSHLANGICH
        self.joriy_holat = "Boshlanmoqda..."
        self._shift = False
        self._lock = threading.Lock()

    def koef(self) -> float:
        return TEZLIK_BAZA / max(self.tezlik, 1)

    def toggle_pauza(self):
        with self._lock:
            self.pauza = not self.pauza
        print(f"\n   {C.YELLOW}⏸  PAUZA — davom uchun F7.{C.R}" if self.pauza
              else f"\n   {C.GREEN}▶️  DAVOM ETILDI.{C.R}")

    def sekinlat(self):
        with self._lock:
            self.tezlik = max(TEZLIK_MIN, self.tezlik - TEZLIK_QADAM)
        print(f"\n   {C.CYAN}🐢 Tezlik: {self.tezlik}%{C.R}")

    def tezlat(self):
        with self._lock:
            self.tezlik = min(TEZLIK_MAX, self.tezlik + TEZLIK_QADAM)
        print(f"\n   {C.CYAN}🐇 Tezlik: {self.tezlik}%{C.R}")


BOSH = Boshqaruv()


def _tugma_listener():
    if not PYNPUT_BOR:
        print(f"   {C.GREY}ℹ️  pynput o'rnatilmagan — F7/Shift o'chiq "
              f"(pip install pynput){C.R}")
        return None

    def on_press(key):
        try:
            if key in (_kb.Key.shift, _kb.Key.shift_l, _kb.Key.shift_r):
                BOSH._shift = True
                return
            if key == _kb.Key.f7:
                BOSH.toggle_pauza()
                return
            ch = getattr(key, "char", None)
            if BOSH._shift or ch in ("!", "@"):
                if ch in ("1", "!"):
                    BOSH.sekinlat()
                elif ch in ("2", "@"):
                    BOSH.tezlat()
        except Exception:
            pass

    def on_release(key):
        if key in (_kb.Key.shift, _kb.Key.shift_l, _kb.Key.shift_r):
            BOSH._shift = False

    lis = _kb.Listener(on_press=on_press, on_release=on_release)
    lis.daemon = True
    lis.start()
    return lis



# PAUZA
def pauza_kut():
    while BOSH.pauza:
        time.sleep(0.15)

def nafas(sek: float):
    pauza_kut()
    time.sleep(max(sek, 0) * BOSH.koef())

def kut(shart_fn, max_sec: float = WAIT_SEC, tavsif: str = "") -> bool:
    sarflandi = 0.0
    while sarflandi < max_sec:
        pauza_kut()
        try:
            if shart_fn():
                return True
        except Exception:
            pass
        time.sleep(POLL * BOSH.koef())
        sarflandi += POLL
    if tavsif:
        print(f"   {C.GREY}⏳ {tavsif} — vaqt tugadi ({max_sec}s){C.R}")
    return False



# NATIJA
class Natija(Enum):
    KOCHIRIB_BOLMAYDI = "kochirib_bolmaydi"        # rasm 6
    HUDUDDA_TOPILMADI = "hududda_topilmadi"        # rasm 1
    JORIYDA_MAVJUD    = "joriy_mahallada_mavjud"   # rasm 7
    BOSHQA_MFY        = "boshqa_mfy_davom"         # rasm 5
    YANGI_FUQARO      = "forma_tayyor"             # rasm 4
    NOMALUM           = "nomalum_holat"

class Holat(Enum):
    """Tasnif jarayonidagi oraliq signallar (yakuniy natija emas)."""
    TEKSHIRILMOQDA = "tekshirilmoqda"   # rasm 3 — kutishni uzaytir
    QAYTA_URIN     = "qayta_urin"       # rasm 2 toast — qayta qidirish

KUT_HOLATLAR  = {Natija.YANGI_FUQARO, Natija.BOSHQA_MFY}
SKIP_HOLATLAR = {Natija.KOCHIRIB_BOLMAYDI, Natija.HUDUDDA_TOPILMADI,
                 Natija.JORIYDA_MAVJUD, Natija.NOMALUM}

@dataclass
class Fuqaro:
    jshshir: str
    tugilgan_sana: str
    ism_familiya: str = "Noma'lum"


def bor(driver, xpath: str) -> bool:
    try:
        return len(driver.find_elements(By.XPATH, xpath)) > 0
    except Exception:
        return False

def korinadigan_elementlar(driver, xpath: str):
    try:
        return [e for e in driver.find_elements(By.XPATH, xpath) if e.is_displayed()]
    except Exception:
        return []

def korinadi(driver, xpath: str) -> bool:
    return len(korinadigan_elementlar(driver, xpath)) > 0

def body_matni(driver) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""

def elementni_bos(driver, el) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        driver.execute_script("arguments[0].click();", el)
        return True
    except Exception:
        try:
            el.click()
            return True
        except Exception:
            return False

def xpath_bos(driver, xpath: str, timeout: float = WAIT_SEC) -> bool:
    try:
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))
        return elementni_bos(driver, btn)
    except Exception:
        pass
    els = korinadigan_elementlar(driver, xpath)
    if els:
        return elementni_bos(driver, els[0])
    return False

def xpath_bos_takror(driver, xpath: str, urinish: int = 4,
                     har_urinish_sec: float = 1.0) -> bool:
    for _ in range(urinish):
        if xpath_bos(driver, xpath, har_urinish_sec):
            return True
        nafas(0.15)
    return False



# BRAUZER
def overlay_yangila(driver):
    try:
        matn = (
            f"SysOne ⚙ {BOSH.tezlik}%  |  "
            f"{'⏸ PAUZA (F7)' if BOSH.pauza else '▶ ISHLAMOQDA'}  |  "
            f"{BOSH.joriy_holat}"
        )
        js = """
        (function(t){
          var id='__sysone_status__';
          var d=document.getElementById(id);
          if(!d){
            d=document.createElement('div');
            d.id=id;
            d.style.cssText='position:fixed;z-index:2147483647;bottom:12px;left:12px;'+
              'background:rgba(10,14,22,.94);color:#7CF6D6;'+
              'font:600 13px/1.4 system-ui,Segoe UI,Arial;padding:8px 13px;'+
              'border:1px solid rgba(124,246,214,.35);'+
              'border-radius:10px;box-shadow:0 6px 22px rgba(0,0,0,.5);'+
              'pointer-events:none;max-width:72vw;letter-spacing:.2px;';
            document.body.appendChild(d);
          }
          d.textContent=t;
        })(arguments[0]);
        """
        driver.execute_script(js, matn)
    except Exception:
        pass


# KALENDAR
CALENDAR_PANEL = (
    "//div[contains(@class,'ant-picker-dropdown') and not(contains(@class,'ant-picker-dropdown-hidden'))] | "
    "//div[contains(@class,'ant-picker-panel-container')]"
)

def kalendar_bormi(driver) -> bool:
    return korinadi(driver, CALENDAR_PANEL)

def kalendarni_yop(driver):
    if not kalendar_bormi(driver):
        return
    try:
        driver.execute_script("document.body.click();")
    except Exception:
        pass
    if kalendar_bormi(driver):
        try:
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
    kut(lambda: not kalendar_bormi(driver), max_sec=2)

# MA'LUMOT YUKLASH
def jshshir_sana(jshshir: str) -> str:
    try:
        if len(jshshir) != 14 or not jshshir.isdigit():
            return ""
        a = int(jshshir[0])
        k, o, y = jshshir[1:3], jshshir[3:5], jshshir[5:7]
        if a in [1, 2]:   y = "18" + y
        elif a in [3, 4]: y = "19" + y
        elif a in [5, 6]: y = "20" + y
        else: return ""
        return f"{k}.{o}.{y}"
    except Exception:
        return ""

def fuqarolarni_yukla(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Fayl topilmadi: {path}")
    df = (pd.read_csv(path, dtype=str) if path.lower().endswith('.csv')
          else pd.read_excel(path, dtype=str))
    ism_col = next((c for c in df.columns if any(
        k in str(c).upper() for k in
        ['F.I.O', 'FIO', 'ФИО', 'Ф.И.О', 'ISM', 'Ф.И.Ш'])), None)
    result = []
    for _, row in df.iterrows():
        fio = (str(row[ism_col]).strip()
               if ism_col and pd.notna(row.get(ism_col)) else "Noma'lum")
        jshshir = next(
            (str(v).strip().replace('.0', '').replace(' ', '')
             for v in row.values
             if len(str(v).strip().replace('.0', '').replace(' ', '')) == 14
             and str(v).strip().replace('.0', '').replace(' ', '').isdigit()),
            None)
        if jshshir:
            sana = jshshir_sana(jshshir)
            if sana:
                result.append(Fuqaro(jshshir=jshshir, tugilgan_sana=sana,
                                     ism_familiya=fio))
    return result


# NATIJALARNI SAQLASH (CSV + EXCEL)
_natija_satrlar = []

def log_yoz(fuqaro: Fuqaro, natija: Natija, izoh: str = "", urinish=None):
    vaqt = time.strftime("%Y-%m-%d %H:%M:%S")
    qator = {
        "№": len(_natija_satrlar) + 1,
        "F.I.O": fuqaro.ism_familiya,
        "JSHSHIR (PNFL)": fuqaro.jshshir,
        "Tug'ilgan sana": fuqaro.tugilgan_sana,
        "Natija": natija.value,
        "Qidirish urinishlari": urinish if urinish else "",
        "Izoh": izoh,
        "Vaqt": vaqt,
    }
    _natija_satrlar.append(qator)
    try:
        yangi = not os.path.exists(NATIJA_LOG)
        with open(NATIJA_LOG, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if yangi:
                w.writerow(list(qator.keys()))
            w.writerow(list(qator.values()))
    except Exception as e:
        print(f"   {C.RED}⚠️  CSV yozishda xato: {e}{C.R}")
    try:
        pd.DataFrame(_natija_satrlar).to_excel(NATIJA_XLSX, index=False)
    except Exception as e:
        print(f"   {C.RED}⚠️  XLSX xato (pip install openpyxl): {e}{C.R}")


# BRAUZER
def brauzerni_ochish() -> webdriver.Chrome:
    opt = webdriver.ChromeOptions()
    opt.add_argument("--start-maximized")
    opt.add_experimental_option("excludeSwitches", ["enable-automation"])
    opt.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=opt)
    driver.get(PLATFORMA_URL)
    input(f"\n{C.GREEN}✅ Tizimga kiring, sahifa ochilsin, "
          f"keyin ENTER bosing...{C.R}\n")
    return driver



# OYNALAR — XPATH'LAR (sarlavhaga emas, mazmun/tugmaga tayanadi)
DIALOG_GENERIK = ("//div[@role='dialog'] | //div[contains(@class,'swal2-container')] | "
                  "//div[contains(@class,'modal-content')]")

MAIN_DIALOG = "//div[@role='dialog'][.//*[contains(text(),'Оила аъзолари')]]"

DAVOM_BTN_ANY = "//button[contains(.,'Давом этиш') or contains(.,'Davom etish')]"
BEKOR_BTN_ANY = "//button[contains(.,'Бекор қилиш') or contains(.,'Bekor qilish')]"

FORM_REESTR   = ("//*[contains(text(),'РЕЕСТРДАН') or "
                 "contains(text(),'сиз тўлдирас') or "
                 "contains(text(),'siz to')]")
FORM_TELEFON  = ("//input[contains(@placeholder,'+998')] | "
                 "//*[contains(text(),'телефон рақами')]")
FORM_OILA_SEC = "//*[contains(text(),'Оила маълумотлари')]"

CLOSE_BTN_GENERIK = (
    "//div[@role='dialog']//button[@aria-label='Close'] | "
    "//div[@role='dialog']//button[contains(@class,'close')] | "
    "//div[@role='dialog']//*[contains(@class,'anticon-close')]/.. | "
    "//button[contains(@class,'swal2-cancel')]"
)

SPINNER = ("//*[contains(@class,'loading') or contains(@class,'spinner') or "
           "contains(@class,'skeleton') or contains(@class,'ant-spin')]")
QIDIRISH_BTN = MAIN_DIALOG + "//button[contains(.,'Қидириш') or contains(.,'Qidirish')]"



# TASNIF
SIG_KOCHIRIB  = ["кўчириб бўлмайди", "чириб бўлмайди"]            # rasm 6
SIG_JORIY     = ["қайта қўшиш мумкин эмас", "рўйхатида мавжуд"]   # rasm 7
SIG_HUDUD     = ["ҳудудда топилмади", "удудда топилмади",
                 "танланган туман"]                              # rasm 1
SIG_BOSHQA    = ["аллақачон рўйхатга олинган", "эски рўйхат"]     # rasm 3/5
SIG_TEKSHIR   = ["текширил"]                                     # rasm 3
SIG_TOAST_YOQ = ["маълумоти топилмади"]                          # rasm 2
SIG_FORMA     = ["реестрдан", "сиз тўлдирас",
                 "телефон рақами", "оила маълумотлари"]          # rasm 4


def _matnda(txt: str, signs) -> bool:
    return any(s in txt for s in signs)

def dialog_bormi(driver) -> bool:
    try:
        return any(e.is_displayed()
                   for e in driver.find_elements(By.XPATH, DIALOG_GENERIK))
    except Exception:
        return False

def forma_topildi(driver) -> bool:
    return (korinadi(driver, FORM_REESTR) or korinadi(driver, FORM_TELEFON)
            or korinadi(driver, FORM_OILA_SEC))



# OYNANI YOPISH (skip holatlar — «Bekor qilish»)
def darhol_oynani_yop(driver, max_urinish: int = 6) -> bool:
    kalendarni_yop(driver)
    for _ in range(max_urinish):
        if not dialog_bormi(driver):
            return True
        if kalendar_bormi(driver):
            kalendarni_yop(driver); continue
        if korinadi(driver, BEKOR_BTN_ANY):
            if xpath_bos(driver, BEKOR_BTN_ANY, 1.5):
                nafas(0.25); continue
        if korinadi(driver, CLOSE_BTN_GENERIK):
            if xpath_bos(driver, CLOSE_BTN_GENERIK, 1.5):
                nafas(0.25); continue
        try:
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            nafas(0.25)
            if not dialog_bormi(driver):
                return True
        except Exception:
            pass
        try:
            driver.execute_script("document.body.click();")
            nafas(0.2)
        except Exception:
            pass
    return not dialog_bormi(driver)

def mutlaq_yop(driver, max_urinish: int = 8) -> bool:
    if darhol_oynani_yop(driver, max_urinish=max_urinish):
        return True
    print(f"   {C.YELLOW}🔄 Oyna yopilmadi — sahifa qayta yuklanmoqda...{C.R}")
    try:
        driver.get(PLATFORMA_URL)
        kut(lambda: not dialog_bormi(driver), max_sec=WAIT_SEC,
            tavsif="Sahifa qayta yuklanishi")
    except Exception as e:
        print(f"   {C.RED}❌ Qayta yuklashda xato: {e}{C.R}")
    return not dialog_bormi(driver)


def qoshish_oynasini_ochish(driver):
    if dialog_bormi(driver) or kalendar_bormi(driver):
        mutlaq_yop(driver)
    qoshish = "//button[contains(.,'Қўшиш') or contains(.,'Qoʻshish')]"
    if not xpath_bos_takror(driver, qoshish, urinish=3, har_urinish_sec=WAIT_SEC):
        raise Exception("'Qo'shish' tugmasi topilmadi!")
    if not kut(lambda: korinadi(driver, MAIN_DIALOG), max_sec=WAIT_SEC,
               tavsif="Dialog ochilishi"):
        raise Exception("Dialog ochilmadi!")


# MA'LUMOT KIRITISH
def kiritish(element, matn: str):
    element.click()
    element.send_keys(Keys.CONTROL + "a")
    element.send_keys(Keys.DELETE)
    element.send_keys(matn)
    kut(lambda: element.get_attribute("value") == matn, max_sec=3)

def malumot_kiritish(driver, fuqaro: Fuqaro):
    jshshir_xp = (
        MAIN_DIALOG + "//input[contains(@placeholder,'ЖШШИР') or contains(@placeholder,'JSHSHIR')] |"
        + MAIN_DIALOG + "//*[contains(text(),'ЖШШИР')]/following::input[1]"
    )
    inp = WebDriverWait(driver, WAIT_SEC).until(
        EC.element_to_be_clickable((By.XPATH, jshshir_xp)))
    kiritish(inp, fuqaro.jshshir)

    sana_xp = (MAIN_DIALOG +
               "//input[contains(@placeholder,'kk.oo.yyyy') or "
               "contains(@placeholder,'кк.оо.йййй')]")
    inp = WebDriverWait(driver, WAIT_SEC).until(
        EC.element_to_be_clickable((By.XPATH, sana_xp)))
    kiritish(inp, fuqaro.tugilgan_sana)

    kalendarni_yop(driver)



# NATIJANI TASNIFLASH
def _klassifikatsiya(driver):
    """
    Qaytaradi:
      • Natija.*  — yakuniy holat
      • Holat.TEKSHIRILMOQDA — hali tugamadi, kutishni davom ettir
      • Holat.QAYTA_URIN — toast «топилмади», qayta qidirish kerak
      • None — hali aniq emas
    """
    if kalendar_bormi(driver):
        kalendarni_yop(driver)
    txt = body_matni(driver).lower()

    # 0) Hali tekshirilmoqda (rasm 3) — Davom tugmasi yo'q
    if _matnda(txt, SIG_TEKSHIR) and not korinadi(driver, DAVOM_BTN_ANY):
        return Holat.TEKSHIRILMOQDA

    # 1) Ko'chirib bo'lmaydi (rasm 6)
    if _matnda(txt, SIG_KOCHIRIB):
        return Natija.KOCHIRIB_BOLMAYDI

    # 2) Joriy MFY'da allaqachon mavjud (rasm 7) — Davom YO'Q
    if _matnda(txt, SIG_JORIY) and not korinadi(driver, DAVOM_BTN_ANY):
        return Natija.JORIYDA_MAVJUD

    # 3) Ushbu hududda topilmadi (rasm 1)
    if _matnda(txt, SIG_HUDUD):
        return Natija.HUDUDDA_TOPILMADI

    # 4) Boshqa MFY — «Давом этиш» tugmasi bor (rasm 5)
    if korinadi(driver, DAVOM_BTN_ANY):
        return Natija.BOSHQA_MFY
    # matn bor, lekin tugma hali chiqmagan -> tekshirilmoqda
    if _matnda(txt, SIG_BOSHQA):
        return Holat.TEKSHIRILMOQDA

    # 5) Forma tayyor (rasm 4) — fuqaro topildi, to'ldirishga tayyor
    if forma_topildi(driver) or _matnda(txt, SIG_FORMA):
        return Natija.YANGI_FUQARO

    # 6) Toast «Фуқаро маълумоти топилмади» (rasm 2) -> qayta qidirish
    if _matnda(txt, SIG_TOAST_YOQ):
        return Holat.QAYTA_URIN

    return None


def _aniqla_kut(driver, max_sec: float = ANIQLASH_KUT):
    """Natijani aniqlaydi. Tekshirilmoqda/spinner bo'lsa kutishni uzaytiradi.
    Toast (QAYTA_URIN) chiqsa — darhol qaytaradi (vaqt isrof bo'lmaydi)."""
    boshlanish = time.time()
    deadline = boshlanish + max_sec
    while time.time() < deadline and time.time() < boshlanish + ANIQLASH_CAP:
        pauza_kut()
        n = _klassifikatsiya(driver)

        if n is Holat.QAYTA_URIN:
            return Holat.QAYTA_URIN
        if n is Holat.TEKSHIRILMOQDA or bor(driver, SPINNER):
            deadline = min(time.time() + max_sec, boshlanish + ANIQLASH_CAP)
        elif isinstance(n, Natija):
            return n

        time.sleep(POLL * BOSH.koef())
    return None


# QIDIRISH (+ «topilmadi»da 6 martagacha qayta urinish)
def _qidirish_bos(driver):
    kalendarni_yop(driver)
    if not xpath_bos(driver, QIDIRISH_BTN, WAIT_SEC):
        raise Exception("Qidirish tugmasi topilmadi!")
    nafas(0.3)
    kut(lambda: not bor(driver, SPINNER), max_sec=WAIT_SEC, tavsif="Server javobi")

def qidir_va_aniqla(driver):
    _qidirish_bos(driver)
    print(f"   {C.CYAN}🔍 Qidirish bosildi...{C.R}")
    natija = _aniqla_kut(driver)
    urinish = 1

    while (natija is None or natija is Holat.QAYTA_URIN) and urinish < QIDIRISH_QAYTA_MAX:
        urinish += 1
        print(f"   {C.YELLOW}🔁 Ma'lumot chiqmadi — «Қидириш» qayta "
              f"({urinish}/{QIDIRISH_QAYTA_MAX})...{C.R}")
        BOSH.joriy_holat = f"Qayta qidirish {urinish}/{QIDIRISH_QAYTA_MAX}"
        overlay_yangila(driver)
        if korinadi(driver, QIDIRISH_BTN):
            xpath_bos(driver, QIDIRISH_BTN, WAIT_SEC)
            nafas(0.4)
            kut(lambda: not bor(driver, SPINNER), max_sec=WAIT_SEC,
                tavsif="Server javobi")
        natija = _aniqla_kut(driver)

    if natija is None or natija is Holat.QAYTA_URIN:
        natija = Natija.HUDUDDA_TOPILMADI
    return natija, urinish


# HOLAT 1: TOPILDI / DAVOM — FOYDALANUVCHI SAQLAGUNCHA KUTISH (6 SOAT)
def holat_kut_va_davom(driver, fuqaro: Fuqaro, natija: Natija, sarlavha: str):
    print(f"   {C.GREEN}👤 {sarlavha}: {fuqaro.ism_familiya}{C.R}")
    if natija == Natija.BOSHQA_MFY:
        print(f"   {C.WHITE}✋ «Давом этиш» → formani to'ldirib «Сақлаш».{C.R}")
    else:
        print(f"   {C.WHITE}✋ Formani to'ldirib «Сақлаш» (yoki «Бекор қилиш»).{C.R}")
    print(f"   {C.GREY}⏳ Oyna yopilguncha kutadi "
          f"(eng ko'p {MANUAL_KUTISH_MAX_SOAT:.0f} soat)...{C.R}")
    BOSH.joriy_holat = f"{sarlavha} — QO'LDA saqlang: {fuqaro.ism_familiya}"
    overlay_yangila(driver)

    boshlanish = time.time()
    deadline = boshlanish + MANUAL_KUTISH_MAX_SOAT * 3600
    keyingi = boshlanish + ESLATMA_ORALIGI_SEC
    yopildi = False

    while time.time() < deadline:
        pauza_kut()
        if kalendar_bormi(driver):
            kalendarni_yop(driver)
        if not korinadi(driver, MAIN_DIALOG):
            yopildi = True
            break
        if time.time() >= keyingi:
            qoldi = int((deadline - time.time()) / 60)
            print(f"   {C.GREY}⏳ ... kutilmoqda (~{qoldi} daqiqa qoldi){C.R}")
            overlay_yangila(driver)
            keyingi = time.time() + ESLATMA_ORALIGI_SEC
        time.sleep(POLL)

    mutlaq_yop(driver)
    if yopildi:
        log_yoz(fuqaro, natija, "Foydalanuvchi qo'lda saqladi/bekor qildi")
        print(f"   {C.GREEN}✅ Oyna yopildi — keyingisiga...{C.R}")
    else:
        log_yoz(fuqaro, natija, "Xavfsizlik vaqti (6 soat) tugadi")
        print(f"   {C.YELLOW}⏭  Vaqt tugadi — keyingisiga...{C.R}")


# HOLAT 2: SKIP — «bekor qilish» bosib keyingisiga
def holat_skip(driver, fuqaro: Fuqaro, natija: Natija, sabab: str, urinish=None):
    print(f"   {C.MAG}⏭  {sabab} → «Бекор қилиш», keyingisiga...{C.R}")
    BOSH.joriy_holat = f"{sabab} — o'tkazildi: {fuqaro.ism_familiya}"
    overlay_yangila(driver)
    if korinadi(driver, BEKOR_BTN_ANY):
        xpath_bos(driver, BEKOR_BTN_ANY, 2.0)
        nafas(0.4)
    ok = mutlaq_yop(driver)
    log_yoz(fuqaro, natija, sabab + ("" if ok else " (sahifa qayta yuklandi)"),
            urinish)


# ASOSIY TSIKL
def asosiy():
    banner()
    _tugma_listener()

    try:
        fuqarolar = fuqarolarni_yukla(EXCEL_PATH)
        print(f"\n{C.GREEN}📋 {len(fuqarolar)} ta fuqaro yuklandi.{C.R}\n")
    except Exception as e:
        print(f"{C.RED}❌ {e}{C.R}")
        return

    driver = brauzerni_ochish()
    overlay_yangila(driver)
    ok = err = 0

    for i, fuqaro in enumerate(fuqarolar, 1):
        pauza_kut()
        print(f"\n{C.BLUE}{'─'*58}{C.R}")
        print(f"{C.B}[{i}/{len(fuqarolar)}]{C.R} {fuqaro.ism_familiya} "
              f"{C.GREY}| {fuqaro.jshshir} | {fuqaro.tugilgan_sana}{C.R}")
        BOSH.joriy_holat = f"[{i}/{len(fuqarolar)}] {fuqaro.ism_familiya}"
        overlay_yangila(driver)

        urinish = None
        try:
            qoshish_oynasini_ochish(driver)
            malumot_kiritish(driver, fuqaro)
            natija, urinish = qidir_va_aniqla(driver)
            print(f"  {C.CYAN}📌 {natija.value}  "
                  f"(urinish: {urinish}){C.R}")

            if natija in KUT_HOLATLAR:
                sarlavha = ("Ma'lumot topildi" if natija == Natija.YANGI_FUQARO
                            else "Boshqa MFY (Davom etish)")
                holat_kut_va_davom(driver, fuqaro, natija, sarlavha)
            elif natija == Natija.KOCHIRIB_BOLMAYDI:
                holat_skip(driver, fuqaro, natija, "Ko'chirib bo'lmaydi", urinish)
            elif natija == Natija.HUDUDDA_TOPILMADI:
                holat_skip(driver, fuqaro, natija,
                           f"Hududda topilmadi ({urinish} urinish)", urinish)
            elif natija == Natija.JORIYDA_MAVJUD:
                holat_skip(driver, fuqaro, natija, "Joriy MFY'da mavjud", urinish)
            else:
                holat_skip(driver, fuqaro, Natija.NOMALUM, "Noma'lum holat", urinish)
                err += 1
                continue

            ok += 1

        except Exception as e:
            print(f"  {C.RED}❌ Xatolik: {e}{C.R}")
            err += 1
            try:
                log_yoz(fuqaro, Natija.NOMALUM, f"Dastur xatosi: {e}", urinish)
                mutlaq_yop(driver)
            except Exception:
                pass

        if dialog_bormi(driver) or kalendar_bormi(driver):
            mutlaq_yop(driver)

    BOSH.joriy_holat = "Tugadi ✅"
    overlay_yangila(driver)
    print(f"\n{C.BLUE}{'='*58}{C.R}")
    print(f"{C.GREEN}✅ Muvaffaqiyatli: {ok}{C.R}  |  {C.RED}❌ Xatolik: {err}{C.R}")
    print(f"{C.GREY}📄 Natijalar: '{NATIJA_LOG}'  va  '{NATIJA_XLSX}'{C.R}")
    print(f"{C.DIM}   {AVTOR['brend']} — {AVTOR['versiya']}{C.R}")
    input(f"\n{C.WHITE}Brauzer yopish uchun ENTER...{C.R}")
    driver.quit()


if __name__ == "__main__":
    asosiy()
