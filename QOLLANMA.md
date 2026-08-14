# Botni Railway'da 24/7 ishlatish — to'liq yo'riqnoma

Bu bot endi sizning kompyuteringiz o'chsa yoki PowerShell yopilsa ham
ishlashi uchun Railway.app'ga joylashtiriladi (bepul reja yetarli).

⚠️ **Muhim eslatma:** Railway'da skrinshot o'qish (OCR) va ovozli xabar
funksiyalari ishlamasligi mumkin (chunki ular Tesseract va FFmpeg kabi
alohida dasturlarni talab qiladi, ular Railway serverida standart
o'rnatilmagan). Asosiy funksiyalar — balans, xarajat/daromad yozish,
qarzlar, statistika, diagramma, Excel, jamoaviy balans — hammasi to'liq
ishlaydi.

## 1-qadam — GitHub'da hisob ochish
1. https://github.com saytiga o'ting
2. "Sign up" orqali bepul hisob oching (email va parol bilan)

## 2-qadam — Yangi repository yaratish
1. GitHub'ga kirgach, yuqori o'ng burchakdagi "+" belgisini bosing
2. "New repository" ni tanlang
3. Nom bering, masalan: expense-bot
4. "Public" yoki "Private" — farqi yo'q, ikkalasi ham bepul
5. "Create repository" tugmasini bosing

## 3-qadam — Fayllarni yuklash
1. Yaratilgan repository sahifasida "uploading an existing file" havolasini bosing
   (yoki "Add file" → "Upload files")
2. Quyidagi fayllarni sudrab tashlang (yoki tanlab yuklang):
   - expense_bot.py
   - requirements.txt
   - Procfile
   - .gitignore
3. Pastda "Commit changes" tugmasini bosing

## 4-qadam — Railway'da hisob ochish
1. https://railway.app saytiga o'ting
2. "Login with GitHub" tugmasi orqali kiring — GitHub hisobingiz bilan
   avtomatik bog'lanadi, alohida parol kerak emas

## 5-qadam — Botni joylashtirish (Deploy)
1. Railway'da "New Project" tugmasini bosing
2. "Deploy from GitHub repo" ni tanlang
3. 2-qadamda yaratgan "expense-bot" repository'ni tanlang
4. Railway avtomatik requirements.txt va Procfile'ni topib, kerakli
   kutubxonalarni o'zi o'rnatadi (bir necha daqiqa vaqt oladi)

## 6-qadam — BOT_TOKEN ni kiritish
1. Railway loyihangiz ichida "Variables" bo'limiga o'ting
2. "New Variable" tugmasini bosing
3. Nomi: BOT_TOKEN
4. Qiymati: sizning bot tokeningiz (masalan 8879962610:AAE7xPzH...)
5. Saqlang — Railway avtomatik botni qayta ishga tushiradi

## 7-qadam — Tekshirish
1. Railway'da "Deployments" bo'limida jarayon tugashini kuting (1-3 daqiqa)
2. Loglarda "Bot ishga tushdi..." degan yozuv chiqishi kerak
3. Telegram'da botingizga /start yozib ko'ring — javob bersa, tayyor!

Shundan keyin bot doim ishlab turadi — kompyuteringizni o'chirsangiz ham,
PowerShell'ni yopsangiz ham, hech narsa o'zgarmaydi, bot ishlashda davom
etadi.

## Eslatma: ma'lumotlar bazasi
Railway'ning bepul rejasida fayl tizimi har safar qayta deploy qilinganda
tozalanishi mumkin (balansingiz, tarixingiz o'chib ketishi mumkin). Agar
uzoq muddat ishlatmoqchi bo'lsangiz, keyinroq Railway'ning doimiy
saqlaydigan bazasiga (PostgreSQL yoki Volume) o'tkazib berishim mumkin —
ayting, yordam beraman.
