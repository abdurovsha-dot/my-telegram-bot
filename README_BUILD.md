# SysOne Mahalla — Build qo'llanmasi

`mahalla.ijro.uz` avto-yordamchi dasturini Windows `.exe` ga aylantirish.

---

## Papkada nima bo'lishi kerak

Bitta papkaga quyidagilarni joylang:

```
📁 SysOne_Build/
 ├── SysOne_Mahalla.py        ← asosiy dastur
 ├── build.bat                ← bir tugma bilan build
 ├── requirements.txt         ← kerakli kutubxonalar
 ├── version_info.txt         ← exe xossalari (brend)
 └── sysone.ico               ← ICON (siz keyin joylaysiz)
```

> `sysone.ico` hozir yo'q — keyin qo'shasiz. Iconsiz ham build bo'ladi,
> faqat standart Python icon bilan chiqadi. Icon qo'shgach `build.bat` ni
> qayta ishga tushiring.

---

## Build qilish (3 qadam)

1. **Python** o'rnatilgan bo'lsin (3.10+). O'rnatishda **"Add to PATH"** ni belgilang.
2. Yuqoridagi papkani oching, **`build.bat`** ni ikki marta bosing.
3. Tugagach, **`dist\SysOne Mahalla.exe`** hosil bo'ladi.

build.bat o'zi: kutubxonalarni o'rnatadi → iconni tekshiradi → exe yasaydi.

---

## Icon (sysone.ico) haqida

- Kerak format: **`.ico`** (PNG yoki JPG to'g'ridan-to'g'ri ishlamaydi).
- Eng yaxshisi — ko'p o'lchamli `.ico` (16, 32, 48, 256 px).
- PNG dan ICO ga aylantirish:
  - Onlayn: convertio.co, icoconvert.com
  - Yoki menga PNG/JPG rasmni yuboring — men `.ico` qilib beraman.
- Faylni aynan **`sysone.ico`** deb nomlab, build papkasiga qo'ying.

---

## Exe ni ishlatish

`SysOne Mahalla.exe` yoniga **`ro'yhat.xlsx`** faylini qo'ying (PNFL/JSHSHIR
ustuni bilan). Exe ishga tushganda:

1. SysOne banneri + avtor ma'lumoti chiqadi.
2. Chrome ochiladi → tizimga kirasiz → ENTER bosasiz.
3. Bot navbatma-navbat qidiradi, topilganda formani SIZ to'ldirib saqlaysiz.
4. Natijalar `natijalar_hisoboti.xlsx` va `.csv` ga yoziladi.

**Shartlar:**
- Google Chrome o'rnatilgan bo'lsin.
- Birinchi ishga tushishda internet kerak (Selenium ChromeDriver'ni yuklaydi).

**Boshqaruv:** `F7` pauza/davom · `Shift+1` sekin · `Shift+2` tez.

---

## Tez-tez uchraydigan muammolar

| Muammo | Yechim |
|--------|--------|
| `python topilmadi` | Python'ni PATH bilan qayta o'rnating |
| Antivirus exe'ni o'chiradi | PyInstaller exe'lari ba'zan false-positive beradi — istisno (exclusion) qo'shing yoki imzolang |
| Chrome ochilmadi | Chrome o'rnatilganini va internetni tekshiring |
| exe juda katta (~80–150 MB) | Normal holat — pandas/selenium ichida. `--onefile` shunaqa |
| Icon o'zgarmadi | Eski iconni Windows kesh qiladi — exe nomini o'zgartiring yoki keshni tozalang |

---

© 2026 SysOne Digital Solutions — Qobilbek Odilov
