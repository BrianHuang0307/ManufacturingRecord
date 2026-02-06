SELECT DISTINCT shb10 生產料件,
                imaud03 品名簡稱,
                sfb95 特性編碼,
                shb06 工序,
                shb09 機台編號,
                shb05 工單編號,
				shb03 完工日期,
                shb031 完工時間,
                shb07 工作中心,
                shb081 作業編號,
                shb082 作業名稱,
                shb032 "投入工時(分鐘)",
                shb033 "投入機時(分鐘)",
                shb111 良品轉出數量,
                shb112 當站報廢數量,
                ecb19 "標準人工工時(秒/pcs)",
                ecb21 "標準機器工時(秒/pcs)",
                sfb09 完工入庫數量
FROM t1_shb_file shbf
LEFT JOIN t1_sfb_file sfbf ON shbf.shb05 = sfbf.sfb01
LEFT JOIN t1_ecb_file ecbf ON shbf.shb10 = ecbf.ecb01
                          AND sfbf.sfb95 = ecbf.ecb02
                          AND shbf.shb06 = ecbf.ecb03
LEFT JOIN ima_tmp imat ON shbf.shb10 = imat.ima01
WHERE (shb03 between :from_date AND :to_date)
AND shbplant = 'TURVO-1'
AND sfb02 = 1 --1.一般工單 15.試產工單
AND sfb04 = 8 --1.開立 2.發放 3.列印 4.發料 5.WIP 6.FQC 7.入庫 8.結案
AND shb09 IS NOT NULL
AND sfb87 = 'Y'
AND shb032 > 0
ORDER BY shb10, sfb95, SHB06, SHB09