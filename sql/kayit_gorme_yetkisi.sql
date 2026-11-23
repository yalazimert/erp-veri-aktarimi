-- =====================================================================
-- KAYIT GÖRME YETKİSİ
-- Kural: Bir kullanıcı kendi girdiği kayıtları + "Kaydı Görecek" olarak
--        atandığı kullanıcıların girdiği kayıtları görür.
-- =====================================================================

-- 1) Tanım tablosu (Excel'deki eşleşmeler). Excel değişince sadece bu
--    tablo güncellenir, raporlara dokunulmaz.
CREATE TABLE IF NOT EXISTS rp_kayit_gorme_tanim (
    giren_user_id   INTEGER NOT NULL,   -- Kaydı Giren UsrId
    gorecek_user_id INTEGER NOT NULL,   -- Kaydı Görecek UsrId
    PRIMARY KEY (giren_user_id, gorecek_user_id)
);

-- "Kaydı Görecek" boş olan satırlar eklenmez; o kullanıcılar zaten
-- fonksiyon gereği kendi kayıtlarını görür.
INSERT INTO rp_kayit_gorme_tanim (giren_user_id, gorecek_user_id) VALUES
 (11969,11949),(11988,11946),(11996,11985),(11973,11946),(11979,12013),
 (11978,11946),(11972,11946),(11987,11949),(11981,11945),(11945,12013),
 (12089,11946),(11999,11948),(11974,11946),(11962,11946),(12000,11949),
 (12001,11949),(12002,11947),(12003,11948),(12004,11949),(11989,11946),
 (12005,11993),(12007,11993),(11982,11948),(12009,11949),(11984,11947),
 (12012,11948),(11986,11949),(11992,11950),(11991,11950),(11951,11950),
 (12061,11950),(11998,11950),(12045,11950)
ON CONFLICT DO NOTHING;

-- 2) Yetki fonksiyonu (rpa_invd_whouse ile aynı kalıp).
--    pUserID'nin görebileceği "kaydı giren" kullanıcıları oturuma özel
--    RP_KAYIT_GORME geçici tablosuna doldurur.
--    pZincir = TRUE  -> hiyerarşi zincirlenir
--                       (11981 -> 11945 -> 12013 ise 12013, 11981'i de görür)
--    pZincir = FALSE -> sadece doğrudan atanmışlar
CREATE OR REPLACE FUNCTION rpa_kayit_gorme(pUserID INTEGER, pZincir BOOLEAN DEFAULT FALSE)
RETURNS VOID
LANGUAGE plpgsql
AS $f$
BEGIN
    CREATE TEMPORARY TABLE IF NOT EXISTS rp_kayit_gorme (
        user_id INTEGER PRIMARY KEY
    ) ON COMMIT DELETE ROWS;

    DELETE FROM rp_kayit_gorme;

    INSERT INTO rp_kayit_gorme (user_id)
    WITH RECURSIVE g(user_id) AS (
        SELECT pUserID                     -- kendi kayıtları
        UNION                              -- UNION: döngü olursa sonsuza gitmez
        SELECT t.giren_user_id
          FROM rp_kayit_gorme_tanim t
          JOIN g ON t.gorecek_user_id = g.user_id
         WHERE pZincir OR g.user_id = pUserID
    )
    SELECT user_id FROM g;
END;
$f$;

-- 3) Rapor içinde kullanım (örnek):
--
-- DO $$
-- DECLARE result CONSTANT refcursor := 'result';
-- BEGIN
--   PERFORM rpa_kayit_gorme(@UsrId@);
--   OPEN result FOR
--   SELECT m.*
--     FROM <kayit_tablosu> m
--    INNER JOIN rp_kayit_gorme k ON k.user_id = m.create_user_id;
-- END $$;
-- FETCH ALL FROM result;
