SELECT DISTINCT shb10,--生產料件
                sfb95,--特性編碼
                shb05,--工單編號
                shb09,--機台編號
				shb06,--工序
				shb03 --報工日期
FROM shb_file shbf
LEFT JOIN sfb_file sfbf
ON shbf.shb05 = sfbf.sfb01
WHERE (shb03 between :from_date AND :to_date)
AND shbplant = 'TURVO-1'
AND sfb02 = 1
AND sfb04 = 8
AND shb09 IS NOT NULL
AND sfb87 = 'Y'
ORDER BY shb10, sfb95, SHB06, SHB09