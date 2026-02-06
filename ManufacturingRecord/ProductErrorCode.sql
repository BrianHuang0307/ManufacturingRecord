SELECT        shb10 生產料件,
                sfb06 特性編碼,
                shb06 工序,
                shb01 移轉單號,
                shb03 完工日期,
                shb04 "員工編號(報工)",
                shb05 工單編號,
                -- shb08 "線/班別",
                shb082 作業名稱,
                shb09 機台編號,
                shb111 良品轉出數量,
                shb112 當站報廢數量,
                qce03 異常原因,
                shc05 數量,
                ROUND(NVL(shc05 / NULLIF(sfb09, 0), 0), 8) AS 異常比例,
                shc06 責任歸屬製程,
                -- SHCPLANT 所屬營運中心, -- TURVO-1
                -- SHCLEGAL 所屬法人, -- TURVO-1
                SHCORIU 資料建立者,
                TA_SHC01 異常備註,
                TA_SHC02 "不良品/重工品", -- 不良:1、重工:2
                sfb09 完工入庫數量
FROM t1_shb_file shbf
LEFT JOIN t1_shc_file shcf ON shbf.shb01 = shcf.shc01
LEFT JOIN t1_sfb_file sfbf ON shbf.shb05 = sfbf.sfb01
LEFT JOIN t1_qce_file qcef ON shcf.shc04 = qcef.qce01
WHERE sfb02 = 1 --1.一般工單 15.試產工單
AND sfb04 = 8 --1.開立 2.發放 3.列印 4.發料 5.WIP 6.FQC 7.入庫 8.結案
AND shb09 IS NOT NULL
AND sfb87 = 'Y'
AND shc04 IS NOT NULL
-- AND ta_shc01 IS NOT NULL
-- AND shcdate > TO_DATE('2025/1/1', 'YYYY/MM/DD')

/*
SELECT          shb10 生產料件,
                sfb06 特性編碼,
                shb06 工序,
                shb01 移轉單號,
                shb03 完工日期,
                shb04 員工編號,
                shb05 工單編號,
                shb07 工作中心,
                shb08 "線/班別",
                shb081 作業編號,
                shb082 作業名稱,
                shb09 機台編號,
                shb111 良品轉出數量,
                shb112 當站報廢數量,
                shb113 重工轉出數量,
                shb114 當站下線數量,
                shb115 Bonus數量,
                shc03 行序,
                shc04 缺點碼,
                qce03 異常原因,
                shc05 數量,
                ROUND(NVL(shc05 / NULLIF(sfb09, 0), 0), 8) AS 異常比例,
                shc06 責任歸屬製程,
                SHCACTI 資料有效碼,
                SHCUSER 資料所有者,
                SHCGRUP 資料所有部門,
                SHCMODU 資料修改者,
                SHCDATE 最近修改日,
                SHCPLANT 所屬營運中心,
                SHCLEGAL 所屬法人,
                SHCORIG 資料建立部門,
                SHCORIU 資料建立者,
                TA_SHC01 異常備註,
                TA_SHC02 "不良品/重工品",
                sfb09 完工入庫數量
FROM t1_shb_file shbf
LEFT JOIN t1_shc_file shcf ON shbf.shb01 = shcf.shc01
LEFT JOIN t1_sfb_file sfbf ON shbf.shb05 = sfbf.sfb01
LEFT JOIN t1_qce_file qcef ON shcf.shc04 = qcef.qce01
WHERE shcdate > TO_DATE('2025/1/1', 'YYYY/MM/DD')
AND ta_shc01 IS NOT NULL
AND sfb02 = 1 --1.一般工單 15.試產工單
AND sfb04 = 8 --1.開立 2.發放 3.列印 4.發料 5.WIP 6.FQC 7.入庫 8.結案
AND shb09 IS NOT NULL
AND sfb87 = 'Y'
*/
