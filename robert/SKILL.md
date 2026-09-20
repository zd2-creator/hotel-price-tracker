---
name: hotel-bookings
description: ניהול הזמנות המלונות של הטיול — הוספה מצילום מסך של הזמנה (Booking/Agoda/Expedia), עדכון, החלפה (הזמנה מחדש), מחיקה ושליפה. מופעל כשזכי שולח צילום מסך של הזמנת מלון או כותב "תוסיף את ההזמנה", "מה המלונות שלי", "תמחק את המלון", "הזמנתי מחדש", "כמה לילות סגורים".
---

# הזמנות מלונות (טיול 2027)

הנתונים יושבים בגיליון גוגל "מעקב הזמנות מלונות" ומוצגים באפליקציית מובייל. הגישה דרך API של Apps Script.

## פרטי גישה
- **API_URL:** `__API_URL__`
- **TOKEN:** `__API_TOKEN__`
- כל בקשה: `POST` עם JSON `{"token": TOKEN, "action": ..., "by": "robert", ...}`
- הרצה:
```bash
curl -sL -X POST "__API_URL__" -H "Content-Type: text/plain" \
  -d '{"token":"__API_TOKEN__","action":"list"}'
```
(חובה `-L` — Apps Script מחזיר redirect. שים לב ל-`Content-Type: text/plain`.)

## שדות הזמנה (booking)
| שדה | תיאור | דוגמה |
|---|---|---|
| hotel | שם המלון (כמו שכתוב בהזמנה, באנגלית) | `Zee Luxury Boutique Hotel` |
| destination | יעד: עיר/אי, מדינה | `קופנגן, תאילנד` |
| platform | Booking / Agoda / Expedia / Hotels.com / Trip.com / Airbnb / ישירות מול המלון | `Booking` |
| account | היוזר/החשבון שהזמין (מייל או שם) — אם רואים בצילום | `zachi@gmail` |
| checkIn / checkOut | `YYYY-MM-DD` | `2027-08-11` |
| roomType | סוג החדר | `Deluxe Pool View` |
| breakfast | true/false — כלול ארוחת בוקר? | `true` |
| price | הסכום הכולל ששולם/ישולם | `6091` |
| currency | ILS / USD / EUR / THB … | `ILS` |
| priceIls | בשקלים — להשאיר ריק אם המטבע אינו ILS (השרת ממיר לבד) | |
| freeCancelUntil | ביטול חינם עד `YYYY-MM-DD` (ריק = לא ניתן לביטול) | `2027-08-04` |
| paid | true אם שולם מראש, false אם תשלום במלון | `false` |
| confirmation | מספר הזמנה/אישור | `4512.887.334` |
| link | קישור להזמנה אם יש | |
| notes | כל דבר נוסף (PIN, שעת הגעה, מדיניות ביטול חלקית) | |

## זרימת עבודה כשזכי שולח צילום מסך של הזמנה
1. **קרא מהתמונה** את כל השדות למעלה. שים לב במיוחד ל: שם המלון המדויק, תאריכים (בוקינג מציג "Wed, 11 Aug 2027"), "Free cancellation until/before <date>" (בוקינג לרוב מציג עד שעה מסוימת — קח את התאריך), "Breakfast included" / "Breakfast not included", הסכום הסופי (Total price) והמטבע, מספר ההזמנה (Confirmation number / Booking ID / Itinerary number).
2. **חובה לפני שמירה:** הצג לזכי סיכום קצר של מה זיהית (מלון, פלטפורמה, תאריכים, לילות, מחיר, ביטול עד, א.בוקר) ושאל אם לשמור — **אלא אם** הוא כבר כתב במפורש "תוסיף"/"תשמור" יחד עם התמונה, ואז שמור מיד ודווח.
3. שלח `upsert`:
```bash
curl -sL -X POST "__API_URL__" -H "Content-Type: text/plain" -d '{
  "token":"__API_TOKEN__","action":"upsert","by":"robert",
  "booking":{"hotel":"...","destination":"...","platform":"Booking","checkIn":"2027-08-11","checkOut":"2027-08-18",
             "roomType":"...","breakfast":true,"price":6091,"currency":"ILS","freeCancelUntil":"2027-08-04",
             "paid":false,"confirmation":"...","account":"...","notes":"..."}
}'
```
   - `upsert` מחפש הזמנה **פעילה קיימת לאותו מלון**:
     - אין → מוסיף. התשובה `{"ok":true,"booking":{...},"replaced":null}`.
     - יש אחת → **מחליף אותה** (הישנה עוברת להיסטוריה עם "הוחלף", החדשה נכנסת). התשובה כוללת `"replaced":{...}` — דווח לזכי: "החלפתי את ההזמנה הקודמת (Agoda, 6,400 ₪) בחדשה (Booking, 6,091 ₪). הישנה נשמרה בהיסטוריה."
     - כמה דומות → `{"ok":false,"needsChoice":true,"candidates":[...]}` — הצג לזכי את המועמדות (id, פלטפורמה, תאריכים, מחיר) ושאל איזו להחליף, ואז שלח שוב עם `"replaceId":"<id>"`.
   - אם זכי אומר במפורש "זה מלון נוסף / לא להחליף" → שלח `action:"add"` (בלי החלפה).
4. **דווח** בקצרה: מה נשמר, ה-summary שחזר (`summary.nights` מתוך `summary.targetNights`, `summary.totalIls`), ואם הוחלף — מה הוחלף. אם `priceIls` חזר ריק, ציין שההמרה לשקל לא הצליחה.

## פעולות נוספות
- **רשימה:** `{"action":"list"}` → `bookings[]` + `summary{count,nights,totalIls,targetNights}`. הצג טבלה קצרה: מלון · פלטפורמה · תאריכים · לילות · ₪ · ביטול עד.
- **עדכון שדה:** `{"action":"update","id":"<id>","booking":{"freeCancelUntil":"2027-08-05"}}` (רק השדות שמשתנים).
- **מחיקה:** `{"action":"delete","id":"<id>","note":"סיבה"}` — נשמר בהיסטוריה. **תמיד לאשר עם זכי לפני מחיקה** אלא אם הוא כתב במפורש "תמחק את X".
- **החלפה מפורשת:** `{"action":"add","replaceId":"<id>","booking":{...}}`.
- **היסטוריה:** `{"action":"history","limit":50}`.
- **שחזור** מהיסטוריה: `{"action":"restore","id":"<id>"}`.
- **חיפוש id** של מלון: קח מ-`list` לפי שם המלון (התאמה חלקית, לא רגיש לאותיות).

## כללים
- תאריכים תמיד `YYYY-MM-DD`. שנת הטיול היא 2027 אלא אם כתוב אחרת.
- אל תמציא שדות שלא מופיעים בצילום — השאר ריק וציין לזכי מה חסר (למשל "לא ראיתי תאריך ביטול חינם").
- מחיר: הסכום **הכולל** לכל השהייה, לא ללילה.
- אם בצילום יש כמה חדרים/הזמנות — שאל האם להזין כל אחת בנפרד.
- אם ה-API מחזיר `"error":"טוקן שגוי"` — הטוקן הוחלף; בקש מזכי טוקן חדש.
