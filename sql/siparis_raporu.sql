DO $$
DECLARE
  result CONSTANT refcursor := 'result';
BEGIN
  PERFORM rpa_kayit_gorme(@UsrId@);   -- ekranı açan kullanıcının görebileceği kişiler

  OPEN result FOR
SELECT
  -------------------BELGE MASTER BILGILERI----------------------------------------------------
  PID.ORDER_D_ID,
  PIM.DOC_NO                               AS SIPARIS_NO,
  PIM.ORDER_M_ID                           AS ORDER_M_ID,
  PIM.DOC_DATE                             AS SIPARIS_TARIH,
  EXTRACT(MONTH FROM PIM.DOC_DATE)         AS "AY",
  PIM.DUE_DAY                              AS SIPARISIN_VADESI,
  GDT.DOC_TRA_CODE                         AS HAREKET_KOD,
  GDT.DESCRIPTION                          AS HAREKET_AD,
  CASE PIM.PURCHASE_SALES
    WHEN 1 THEN 'Alis'
    WHEN 2 THEN 'Satis'
    WHEN 3 THEN 'Satis_Iade'
    WHEN 4 THEN 'Alis_Iade'
  END                                      AS ALIS_SATIS,
  FFT.ENTITY_CODE                          AS MUSTERI_KODU,
  FFT.ENTITY_NAME                          AS MUSTERI_ADI,
  COALESCE(FFT.TAX_NO, '')                 AS CARI_VERGI_NO,
  COALESCE(FFT.IDENTIFY_NO, '')            AS CARI_TC_NO,
  FSP.SALES_PERSON_CODE                    AS SATICI_KOD,
  PIM.ENTITY_ORDER_NO                      AS MUSTERI_SIPARIS_NO,
  --------------------BELGE DETAY BILGILERI----------------------------------------------------
  PID.LINE_NO                              AS SIRA_NO,
  CASE PID.LINE_TYPE
    WHEN 1 THEN 'S'
    WHEN 2 THEN 'M'
    WHEN 3 THEN 'H'
  END                                      AS SATIR_TIP,
  IIT.DCARD_CODE                           AS URUN_KODU,
  IIT.DCARD_NAME                           AS URUN_TANIM,
  PID.ITEM_NAME_MANUAL                     AS STOK_HIZMET_AD_GIRILEN,
  HBR.UNIT_CODE                            AS BIRIM_CINSI,
  WHO.WHOUSE_CODE                          AS DEPO_KOD,
  PID.ENTITY_ORDER_NO                      AS D_MUSTERI_SIPARIS_NO,
  PID.NOTE2                                AS SATIR_ACIKLAMASI,
  PID.VAT_RATE                             AS ORAN_KDV,
  CASE PID.ORDER_STATUS
    WHEN 1 THEN 'ACIK'
    WHEN 2 THEN 'KAPALI'
  END                                      AS D_SIPARIS_DURUM,
  --------------------BELGE MIKTAR BILGILERI---------------------------------------------------
  PID.QTY                                  AS MIKTAR,
  PID.QTY_SHIPPING                         AS SEVK_EDILEN_MIKTAR,
  COALESCE(PID.QTY, 0) - COALESCE(PID.QTY_SHIPPING, 0) AS SEVKIYAT_SONRASI_KALAN,
  --------------------SEVK EMRI BILGILERI-------------------------------------------------------
  SEM.SEVK_EMRI_TARIHI                     AS SEVK_EMRI_TARIHI,
  SEM.SEVK_EMRI_MIKTARI                    AS SEVK_EMRI_MIKTARI,
  --------------------BELGE TUTAR BILGILERI----------------------------------------------------
  HPR.CUR_CODE                             AS PARA_BIRIM,
  PID.AMT_TRA                              AS TOPLAM_FIYAT,
  PID.AMT                                  AS TUTAR_TL,
  ROUND(PID.AMT_VAT / NULLIF(PID.CUR_RATE_TRA, 0), 2) AS TUTAR_KDV,
  PID.UNIT_PRICE_TRA                       AS BIRIM_FIYAT,
  PID.AMT_WITH_DISC_TRA                    AS "İSKONTOLU TUTAR",
  CASE
    WHEN HPR.CUR_CODE = 'USD'
      THEN ROUND(COALESCE(PID.UNIT_PRICE_TRA, 0)::numeric, 4)::numeric(18,4)
    ELSE ROUND(COALESCE(PID.UNIT_PRICE / NULLIF(USDCRQNB.CUR_RATE_TRA, 0), 0)::numeric, 4)::numeric(18,4)
  END                                      AS BIRIM_USD_QNB_SATIS,
  CASE
    WHEN HPR.CUR_CODE = 'USD'
      THEN ROUND(COALESCE(PID.AMT_TRA, 0)::numeric, 2)::numeric(18,2)
    ELSE ROUND(COALESCE((PID.UNIT_PRICE * PID.QTY) / NULLIF(USDCRQNB.CUR_RATE_TRA, 0), 0)::numeric, 2)::numeric(18,2)
  END                                      AS TOPLAM_USD_QNB_SATIS,
  --------------------REZERV BILGILERI---------------------------------------------------------
  STSREZ.REZ_TARIHI,
  STSREZ.REZ_NO,
  STSREZ.REZ_DEPO,
  STSREZ.REZ_MIKTARI,
  --------------------IRSALIYE / FATURA--------------------------------------------------------
  SEVK.IDOC_NO                             AS IRSALIYE_NO,
  SEVK.IDOC_DATE                           AS IRS_TRH,
  SEVK.FDOC_NO                             AS FATURA_NO,
  SEVK.FDOC_DATE                           AS FATURA_TARIH,
  --------------------SATINALMA SIPARISI-------------------------------------------------------
  OOO.BELGE_NO                             AS SATINALMA_SIP_NO,
  OOO.MIKTAR                               AS SATINALMA_MIKTAR,
  OOO.BELGE_TARIHI                         AS SATINALMA_SIP_TRH,
  --------------------STOK KATEGORILERI--------------------------------------------------------
  C01.DESCRIPTION                          AS URUN_TEDARIKCISI,
  C02.DESCRIPTION                          AS URUN_TIPI,
  C03.DESCRIPTION                          AS MARKA,
  C04.DESCRIPTION                          AS HEDEF_KODU,
  C05.DESCRIPTION                          AS BUSINESS_UNIT,
  --------------------DEPO MIKTARLARI----------------------------------------------------------
  DEP.ADANA_QTY1                           AS ADANA_TAKİP_MİKTAR,
  DEP.ADANA_QTY2                           AS ADANA_SERBEST_MİKTAR,
  DEP.ANKARA_QTY1                          AS ANKARA_TAKİP_MİKTAR,
  DEP.ANKARA_QTY2                          AS ANKARA_SERBEST_MİKTAR,
  DEP.IZMIR_QTY1                           AS İZMİR_TAKİP_MİKTAR,
  DEP.IZMIR_QTY2                           AS İZMİR_SERBEST_MİKTAR,
  DEP.MERKEZ_QTY1                          AS MERKEZ_TAKİP_MİKTAR,
  DEP.MERKEZ_QTY2                          AS MERKEZ_KULLANILABILIR,
  --------------------CARI KATEGORILERI / NOTLAR-----------------------------------------------
  G01.DESCRIPTION                          AS BÖLGE,
  G02.DESCRIPTION                          AS SEKTÖR,
  FFT.ADD_STRING01                         AS UYARI,
  PIM.ZZ_NOTE                              AS SİPARİS_GENEL_AÇIKLAMASI,
  OOO.KONSDATE                             AS KONSIMENTO_TARIHI,
  OOO.YUKLEME_NO                           AS YUKLEME_NO
FROM PSMT_ORDER_D PID
  -------------------------MASTER BAGLANTI-----------------------------------------------------------------------------
  INNER JOIN PSMT_ORDER_M PIM ON PID.ORDER_M_ID = PIM.ORDER_M_ID
  -------------------------KAYIT GORME YETKISI-------------------------------------------------------------------------
  INNER JOIN RP_KAYIT_GORME KGY ON KGY.USER_ID = PIM.CREATE_USER_ID
  -------------------------FIRMA ISYERI--------------------------------------------------------------------------------
  INNER JOIN GNLD_COMPANY CMP ON PID.CO_ID = CMP.CO_ID
  INNER JOIN GNLD_BRANCH  BRA ON PID.BRANCH_ID = BRA.BRANCH_ID
  -------------------------HAREKET KOD BAGLANTISI----------------------------------------------------------------------
  INNER JOIN GNLD_DOC_TRA GDT ON PIM.DOC_TRA_ID = GDT.DOC_TRA_ID
  -------------------------CARI KARTLARI-------------------------------------------------------------------------------
  INNER JOIN FIND_ENTITY FFT ON PIM.ENTITY_ID = FFT.ENTITY_ID
  -------------------------STOK HIZMET MUHASEBE KARTLARI---------------------------------------------------------------
  LEFT JOIN INVW_ITEM_TABLES IIT ON PID.DCARD_ID = IIT.DCARD_ID AND PID.LINE_TYPE = IIT.LINETYPE
  -------------------------DEPO BAGLANTISI-----------------------------------------------------------------------------
  LEFT JOIN INVD_WHOUSE WHO ON PID.WHOUSE_ID = WHO.WHOUSE_ID
  -------------------------STOK KART BAGLANTISI------------------------------------------------------------------------
  LEFT JOIN INVD_ITEM ITE ON PID.DCARD_ID = ITE.ITEM_ID AND PID.LINE_TYPE = 1
  -------------------------HAREKET BIRIM / PARA BIRIMI-----------------------------------------------------------------
  LEFT JOIN INVD_UNIT     HBR ON PID.UNIT_ID    = HBR.UNIT_ID
  LEFT JOIN GNLD_CURRENCY HPR ON PID.CUR_TRA_ID = HPR.CUR_ID
  -------------------------SATICI KOD BAGLANTISI-----------------------------------------------------------------------
  LEFT JOIN FIND_SALES_PERSON FSP ON PID.SALES_PERSON_ID = FSP.SALES_PERSON_ID
  -------------------------SEVK EMRI (satir basina tek kayit)----------------------------------------------------------
  LEFT JOIN
  (
    SELECT ROM.SOURCE_M_ID,
           ROD.SOURCE_D_ID,
           STRING_AGG(DISTINCT TO_CHAR(ROM.DOC_DATE, 'DD.MM.YYYY'), ', ') AS SEVK_EMRI_TARIHI,
           SUM(COALESCE(ROD.QTY_SHIPPING, 0))                             AS SEVK_EMRI_MIKTARI
      FROM PSMT_REFERRAL_ORDERS_D ROD
     INNER JOIN PSMT_REFERRAL_ORDERS_M ROM ON ROM.REFERRAL_ORDERS_M_ID = ROD.REFERRAL_ORDERS_M_ID
     GROUP BY ROM.SOURCE_M_ID, ROD.SOURCE_D_ID
  ) SEM ON SEM.SOURCE_M_ID = PIM.ORDER_M_ID AND SEM.SOURCE_D_ID = PID.ORDER_D_ID
  -------------------------ILISKILI SATINALMA SIPARISI (satir basina tek kayit)----------------------------------------
  LEFT JOIN
  (
    SELECT D.ZZ_RELADET_D_ID,
           STRING_AGG(DISTINCT M.DOC_NO, '; ')                               AS BELGE_NO,
           SUM(D.QTY)                                                        AS MIKTAR,
           MIN(M.DOC_DATE)                                                   AS BELGE_TARIHI,
           STRING_AGG(DISTINCT TO_CHAR(SS.KONSDATE, 'DD.MM.YYYY'), '; ')     AS KONSDATE,
           STRING_AGG(DISTINCT SS.DOC_NO, '; ')                              AS YUKLEME_NO
      FROM PSMT_ORDER_D D
     INNER JOIN PSMT_ORDER_M M ON M.ORDER_M_ID = D.ORDER_M_ID
      LEFT JOIN
      (
        SELECT KAYNAK_SASDID,
               MIN(KONSDATE) AS KONSDATE,
               STRING_AGG(DISTINCT DOC_NO, '; ') AS DOC_NO
          FROM ZZMW_SHIPMENT_SOURCE
         GROUP BY KAYNAK_SASDID
      ) SS ON SS.KAYNAK_SASDID = D.ORDER_D_ID
     WHERE D.ZZ_RELADET_D_ID IS NOT NULL
       AND M.PURCHASE_SALES = 1
     GROUP BY D.ZZ_RELADET_D_ID
  ) OOO ON OOO.ZZ_RELADET_D_ID = PID.ORDER_D_ID
  -------------------------QNB SATIS DOLAR KURU (rapor bitis tarihindeki ya da ondan onceki son kur)------------------
  LEFT JOIN
  (
    SELECT CR.CUR_RATE_TRA
      FROM GNLD_DAILY_CUR_RATE CR
     WHERE CR.CUR_RATE_TYPE_ID = 278
       AND CR.CUR_FROM_ID      = 116
       AND CR.DOC_DATE <= (CASE '@DocDateL@'
                             WHEN CHR(64) || 'DocDateL' || CHR(64) THEN CURRENT_DATE
                             WHEN 'null'                          THEN CURRENT_DATE
                             ELSE TO_DATE('@DocDateL@', 'DD.MM.YYYY')
                           END)
     ORDER BY CR.DOC_DATE DESC
     LIMIT 1
  ) USDCRQNB ON TRUE
  -------------------------CARI KATEGORI KOD BAGLANTISI----------------------------------------------------------------
  LEFT JOIN GNLD_CATEGORIES G01 ON FFT.CATEGORIES1_ID = G01.CATEGORIES_ID
  LEFT JOIN GNLD_CATEGORIES G02 ON FFT.CATEGORIES2_ID = G02.CATEGORIES_ID
  -------------------------STOK KATEGORI KOD BAGLANTISI----------------------------------------------------------------
  LEFT JOIN GNLD_CATEGORIES C01 ON ITE.CATEGORIES1_ID = C01.CATEGORIES_ID
  LEFT JOIN GNLD_CATEGORIES C02 ON ITE.CATEGORIES2_ID = C02.CATEGORIES_ID
  LEFT JOIN GNLD_CATEGORIES C03 ON ITE.CATEGORIES3_ID = C03.CATEGORIES_ID
  LEFT JOIN GNLD_CATEGORIES C04 ON ITE.CATEGORIES4_ID = C04.CATEGORIES_ID
  LEFT JOIN GNLD_CATEGORIES C05 ON ITE.CATEGORIES5_ID = C05.CATEGORIES_ID
  -------------------------DEPO ve REZERV MIKTARLARI (4 depo tek seferde)---------------------------------------------
  -- 3688 ADANA, 3686 ANKARA, 3687 IZMIR, 3102 MERKEZ  /  QTY1 = takip, QTY2 = serbest (takip - rezerv)
  LEFT JOIN
  (
    SELECT IT.ITEM_ID,
           SUM(CASE WHEN IT.WHOUSE_ID = 3688 THEN IT.QTY_PRM END)                                AS ADANA_QTY1,
           SUM(CASE WHEN IT.WHOUSE_ID = 3688 THEN IT.QTY_PRM - COALESCE(RS.SUM_RESERVE, 0) END)  AS ADANA_QTY2,
           SUM(CASE WHEN IT.WHOUSE_ID = 3686 THEN IT.QTY_PRM END)                                AS ANKARA_QTY1,
           SUM(CASE WHEN IT.WHOUSE_ID = 3686 THEN IT.QTY_PRM - COALESCE(RS.SUM_RESERVE, 0) END)  AS ANKARA_QTY2,
           SUM(CASE WHEN IT.WHOUSE_ID = 3687 THEN IT.QTY_PRM END)                                AS IZMIR_QTY1,
           SUM(CASE WHEN IT.WHOUSE_ID = 3687 THEN IT.QTY_PRM - COALESCE(RS.SUM_RESERVE, 0) END)  AS IZMIR_QTY2,
           SUM(CASE WHEN IT.WHOUSE_ID = 3102 THEN IT.QTY_PRM END)                                AS MERKEZ_QTY1,
           SUM(CASE WHEN IT.WHOUSE_ID = 3102 THEN IT.QTY_PRM - COALESCE(RS.SUM_RESERVE, 0) END)  AS MERKEZ_QTY2
      FROM INVD_BWH_ITEM IT
      LEFT JOIN INVW_SUM_RESERVE RS ON IT.WHOUSE_ID = RS.WHOUSE_ID AND IT.ITEM_ID = RS.ITEM_ID
     WHERE IT.BRANCH_ID = 6750
       AND IT.WHOUSE_ID IN (3688, 3686, 3687, 3102)
     GROUP BY IT.ITEM_ID
  ) DEP ON DEP.ITEM_ID = ITE.ITEM_ID
  -------------------------SIPARISE AIT ACIK REZERVLER (satir basina tek kayit)----------------------------------------
  LEFT JOIN
  (
    SELECT REZ.ORDER_D_ID,
           STRING_AGG(DISTINCT TO_CHAR(REZ.CREATE_DATE, 'DD.MM.YYYY'), ', ') AS REZ_TARIHI,
           STRING_AGG('WS3-' || REZ.RESERVE_M_ID, ', ')                      AS REZ_NO,
           STRING_AGG(DISTINCT E.WHOUSE_CODE, ', ')                          AS REZ_DEPO,
           SUM(REZ.QTY_REMAIN_PRM)                                           AS REZ_MIKTARI
      FROM INVT_RESERVE_M REZ
      LEFT JOIN INVD_WHOUSE E ON E.WHOUSE_ID = REZ.WHOUSE_ID
     WHERE REZ.OPEN_CLOSE = 1
     GROUP BY REZ.ORDER_D_ID
  ) STSREZ ON STSREZ.ORDER_D_ID = PID.ORDER_D_ID
  -------------------------IRSALIYE VE FATURA NUMARA / TARIH (satir basina tek kayit)----------------------------------
  LEFT JOIN
  (
    SELECT X.SOURCE_M_ID,
           X.SOURCE_D_ID,
           STRING_AGG(DISTINCT X.IDOC_NO,   ', ') AS IDOC_NO,
           STRING_AGG(DISTINCT X.IDOC_DATE, ', ') AS IDOC_DATE,
           STRING_AGG(DISTINCT X.FDOC_NO,   ', ') AS FDOC_NO,
           STRING_AGG(DISTINCT X.FDOC_DATE, ', ') AS FDOC_DATE
      FROM
      (
        -- yurtici: siparis -> irsaliye (-> fatura)
        SELECT ID.SOURCE_M_ID, ID.SOURCE_D_ID,
               IM.DOC_NO                           AS IDOC_NO,
               TO_CHAR(IM.DOC_DATE, 'DD.MM.YYYY')  AS IDOC_DATE,
               FM.DOC_NO                           AS FDOC_NO,
               TO_CHAR(FM.DOC_DATE, 'DD.MM.YYYY')  AS FDOC_DATE
          FROM INVT_ITEM_D ID
          LEFT JOIN INVT_ITEM_M    IM ON ID.ITEM_M_ID = IM.ITEM_M_ID
          LEFT JOIN PSMT_INVOICE_D FD ON ID.ITEM_D_ID = FD.SOURCE_D_ID AND ID.ITEM_M_ID = FD.SOURCE_M_ID
          LEFT JOIN PSMT_INVOICE_M FM ON FD.INVOICE_M_ID = FM.INVOICE_M_ID
        UNION ALL
        -- yurtici: siparis -> fatura
        SELECT FD.SOURCE_M_ID, FD.SOURCE_D_ID,
               NULL, NULL,
               FM.DOC_NO,
               TO_CHAR(FM.DOC_DATE, 'DD.MM.YYYY')
          FROM PSMT_INVOICE_D FD
          LEFT JOIN PSMT_INVOICE_M FM ON FD.INVOICE_M_ID = FM.INVOICE_M_ID
        UNION ALL
        -- ihracat: siparis -> ihracat dosyasi faturasi
        SELECT ED.SOURCE_M_ID, ED.SOURCE_D_ID,
               NULL, NULL,
               EM.INVOICE_DOC_NO,
               TO_CHAR(EM.INVOICE_DATE, 'DD.MM.YYYY')
          FROM FTMT_FILE_E_D ED
          LEFT JOIN FTMT_FILE_E_M EM ON ED.FILE_E_M_ID = EM.FILE_E_M_ID
      ) X
     GROUP BY X.SOURCE_M_ID, X.SOURCE_D_ID
  ) SEVK ON PID.ORDER_D_ID = SEVK.SOURCE_D_ID AND PID.ORDER_M_ID = SEVK.SOURCE_M_ID
WHERE
  ------------------------FIRMA ISYERI FILTRESI------------------------------------------------------------------
      (CMP.CO_CODE     IN (SELECT CODE FROM RP_SPLIT('@CoCode@'))      OR '@CoCode@'      = 'null')
  AND (BRA.BRANCH_CODE IN (SELECT CODE FROM RP_SPLIT('@BranchCodes@')) OR '@BranchCodes@' = 'null')
  ------------------------HESAP FILTRELERI-----------------------------------------------------------------------
  AND (FFT.ENTITY_CODE >= '@EntityCodeF@' OR '@EntityCodeF@' = 'null')
  AND (FFT.ENTITY_CODE <= '@EntityCodeL@' OR '@EntityCodeL@' = 'null')
  ------------------------HAREKET KOD FILTRESI-------------------------------------------------------------------
  AND (GDT.DOC_TRA_CODE IN (SELECT CODE FROM RP_SPLIT('@DocTraCodes@')) OR '@DocTraCodes@' = 'null')
  ------------------------DEPO FILTRESI--------------------------------------------------------------------------
  AND (WHO.WHOUSE_CODE IN (SELECT CODE FROM RP_SPLIT('@WhouseCodes@')) OR '@WhouseCodes@' = 'null')
  ------------------------TEMSILCI FILTRESI----------------------------------------------------------------------
  AND (FSP.SALES_PERSON_CODE = '@SalesPersonCodeF@' OR '@SalesPersonCodeF@' = 'null')
  ------------------------SIPARIS DURUM FILTRESI (0 hepsi, 1 acik, 2 kapali)-------------------------------------
  AND (   ('@ReportOpenClose@' = '1' AND PID.ORDER_STATUS = 1)
       OR ('@ReportOpenClose@' = '2' AND PID.ORDER_STATUS = 2)
       OR ('@ReportOpenClose@' = '0'))
  ------------------------TARIH FILTRESI-------------------------------------------------------------------------
  AND PIM.DOC_DATE::date >= (CASE '@DocDateF@'
                               WHEN CHR(64) || 'DocDateF' || CHR(64) THEN TO_DATE('01.01.0001', 'DD.MM.YYYY')
                               WHEN 'null'                          THEN TO_DATE('01.01.0001', 'DD.MM.YYYY')
                               ELSE TO_DATE('@DocDateF@', 'DD.MM.YYYY')
                             END)
  AND PIM.DOC_DATE::date <= (CASE '@DocDateL@'
                               WHEN CHR(64) || 'DocDateL' || CHR(64) THEN TO_DATE('31.12.9999', 'DD.MM.YYYY')
                               WHEN 'null'                          THEN TO_DATE('31.12.9999', 'DD.MM.YYYY')
                               ELSE TO_DATE('@DocDateL@', 'DD.MM.YYYY')
                             END)
  ------------------------BELGE NO FILTRESI----------------------------------------------------------------------
  AND (PIM.DOC_NO >= '@DocNoF@' OR '@DocNoF@' = 'null')
  AND (PIM.DOC_NO <= '@DocNoL@' OR '@DocNoL@' = 'null')
  ------------------------SADECE SATIS SIPARISLERI---------------------------------------------------------------
  AND PIM.PURCHASE_SALES = 2
  ------------------------KATEGORI FILTRESI----------------------------------------------------------------------
  AND (C01.CATEGORIES_CODE IN (SELECT CODE FROM RP_SPLIT('@Categories1Codes@')) OR '@Categories1Codes@' = 'null')
  AND (C02.CATEGORIES_CODE IN (SELECT CODE FROM RP_SPLIT('@Categories2Codes@')) OR '@Categories2Codes@' = 'null')
ORDER BY CMP.CO_CODE,
         BRA.BRANCH_CODE,
         PIM.DOC_DATE,
         PIM.ORDER_M_ID,
         PID.LINE_NO;
END; $$;
FETCH ALL FROM result;
