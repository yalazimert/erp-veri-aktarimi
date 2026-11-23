# Kayıt Görme Yetkisi – Kullanım Kılavuzu

**Tablo:** `rp_kayit_gorme_tanim`
**Kural:** `giren_user_id` kolonundaki kullanıcının girdiği kayıtları `gorecek_user_id` kolonundaki kullanıcı da görür.
Herkes kendi kayıtlarını zaten görür, bunun için satır eklemeye gerek yok.

Kurulum dosyası: `sql/kayit_gorme_yetkisi.sql` (tablo, fonksiyon, ilk veriler).

## 1. Mevcut tanımları görmek (kullanıcı adlarıyla)
```sql
SELECT t.giren_user_id,   g.us_username kaydi_giren,
       t.gorecek_user_id, v.us_username kaydi_gorecek
  FROM rp_kayit_gorme_tanim t
  LEFT JOIN users g ON g.us_id = t.giren_user_id
  LEFT JOIN users v ON v.us_id = t.gorecek_user_id
 ORDER BY v.us_username, g.us_username;
```

## 2. Kullanıcının ID'sini bulmak
```sql
SELECT us_id, us_username FROM users WHERE us_username ILIKE '%ahmet%';
```

## 3. Yeni yetki eklemek
```sql
INSERT INTO rp_kayit_gorme_tanim (giren_user_id, gorecek_user_id)
VALUES (11961, 11927)
ON CONFLICT DO NOTHING;
```
Birden fazla eşleşme:
```sql
INSERT INTO rp_kayit_gorme_tanim (giren_user_id, gorecek_user_id) VALUES
 (11961, 11927),
 (11962, 11927)
ON CONFLICT DO NOTHING;
```
`ON CONFLICT DO NOTHING` sayesinde aynı satırı iki kez ekleseniz de hata almazsınız.
Bir kaydı birden fazla kişinin görmesi gerekiyorsa her görecek kişi için ayrı satır ekleyin.

## 4. Yetki kaldırmak
```sql
DELETE FROM rp_kayit_gorme_tanim
 WHERE giren_user_id = 11961 AND gorecek_user_id = 11927;
```

## 5. Bir kullanıcının kimleri gördüğünü kontrol etmek
```sql
BEGIN;
SELECT rpa_kayit_gorme(11927);
SELECT k.user_id, u.us_username
  FROM rp_kayit_gorme k LEFT JOIN users u ON u.us_id = k.user_id;
COMMIT;
```

## 6. Excel'den toplu yenilemek
1. Excel'i iki kolonlu CSV olarak kaydedin: `giren_user_id`, `gorecek_user_id`. Görecek alanı boş olan satırları silin.
2. Tabloyu boşaltın:
   ```sql
   TRUNCATE rp_kayit_gorme_tanim;
   ```
3. DBeaver ya da pgAdmin'de tabloya sağ tıklayın, Import Data ile CSV'yi seçin.

`TRUNCATE` bütün yetkileri siler. Yeniden yükleme yapılana kadar herkes yalnızca kendi kayıtlarını görür.

## 7. Yeni bir rapora yetki eklemek
```sql
DO $$
DECLARE
  result CONSTANT refcursor := 'result';
BEGIN
  PERFORM rpa_kayit_gorme(@UsrId@);

  OPEN result FOR
  SELECT m.doc_no,
         COALESCE(u1.us_username, '-') kaydi_giren
    FROM psmt_order_m m
   INNER JOIN rp_kayit_gorme k ON k.user_id = m.create_user_id   -- yetki filtresi
    LEFT JOIN users u1          ON u1.us_id  = m.create_user_id
   ORDER BY m.doc_no;
END; $$;
FETCH ALL FROM result;
```
Başka tabloda: `INNER JOIN rp_kayit_gorme k ON k.user_id = <tablo>.create_user_id`

## Önemli notlar
- Değişiklikler anında geçerli olur. Fonksiyona ya da rapora dokunmanız gerekmez.
- Yetki zincirleme çalışmaz. A'nın kayıtlarını B, B'ninkileri C görüyorsa, C, A'nın kayıtlarını görmez.
  C'nin de görmesi gerekiyorsa (A, C) satırını ayrıca ekleyin.
- `rp_kayit_gorme_tanim` tablosunda test satırı var: (11961, 11927). Gerekmiyorsa silin (bkz. madde 4).
