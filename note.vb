=IF($F2="Z01",
SUMIFS(Sale_Target[Budget 
Jan/2027],Sale_Target[[Product'#4]:[Product'#4]],$E2,Sale_Target[[D-Chanel]:[D-Chanel]],$F2),
(1/24)*SUMIFS(Sale_Target[Budget 
Jan/2027],Sale_Target[[Product'#4]:[Product'#4]],$E2,Sale_Target[[D-Chanel]:[D-Chanel]],$F2)
+ (2/24)*SUMPRODUCT(('Sale Target (R)'!$AC$2:AG$3000) * ('Sale Target (R)'!$F$2:$F$3000=$F2) * ('Sale Target (R)'!$E$2:$E$3000=$E2))
- (2/24)*SUMIFS(Sale_Target[Budget 
Jan/2027],Sale_Target[[Product'#4]:[Product'#4]],$E2,Sale_Target[[D-Chanel]:[D-Chanel]],$F2)
)
