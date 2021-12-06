SELECT M.单位, M.养老个人, M.养老单位, M.医保个人, M.医保单位, M.失业个人, M.失业单位, M.工伤单位, N.生育单位 FROM 
    (SELECT K.单位, K.养老个人, K.养老单位, K.医保个人, K.医保单位, K.失业个人, K.失业单位, L.工伤单位, K.单位编号 FROM 
        (SELECT I.单位, I.养老个人, I.养老单位, I.医保个人, I.医保单位, I.失业个人, J.失业单位, I.单位编号 FROM 
            (SELECT G.单位, G.养老个人, G.养老单位, G.医保个人, G.医保单位, H.失业个人, G.单位编号 FROM 
                (SELECT E.单位, E.养老个人, E.养老单位, E.医保个人, F.医保单位, E.单位编号 FROM 
                    (SELECT C.单位, C.养老个人, C.养老单位, D.医保个人, C.单位编号 FROM 
                            (SELECT A.单位, A.养老个人, B.养老单位, A.单位编号 FROM 
                                (SELECT * FROM 
                                        (SELECT 五险明细表.单位, Sum(五险明细表.个人缴纳) AS 养老个人, 五险明细表.单位编号 FROM 五险明细表 WHERE (((五险明细表.险种)="养老"))
                                        GROUP BY 五险明细表.单位, 五险明细表.单位编号)  AS A 
                                    LEFT JOIN 
                                        (SELECT * FROM 
                                            (SELECT 五险明细表.单位, Sum(五险明细表.单位缴纳) AS 养老单位 FROM 五险明细表 WHERE (((五险明细表.险种)="养老")) 
                                            GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
                                        AS [%$##@_Alias])AS B 
                                    ON A.单位=B.单位)  
                            AS [%$##@_Alias])  AS C 
                    LEFT JOIN 
                        (SELECT * FROM 
                            (SELECT 五险明细表.单位, Sum(五险明细表.个人缴纳) AS 医保个人 FROM 五险明细表 WHERE (((五险明细表.险种)="医保")) 
                            GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
                        AS [%$##@_Alias])  AS D 
                    ON C.单位=D.单位)  AS E 
                LEFT JOIN 
                (SELECT * FROM 
                    (SELECT 五险明细表.单位, Sum(五险明细表.单位缴纳) AS 医保单位 FROM 五险明细表 WHERE (((五险明细表.险种)="医保")) 
                    GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
                AS [%$##@_Alias])  
            AS F ON E.单位=F.单位)  AS G 
            LEFT JOIN 
            (SELECT * FROM 
                (SELECT 五险明细表.单位, Sum(五险明细表.个人缴纳) AS 失业个人 FROM 五险明细表 WHERE (((五险明细表.险种)="失业")) 
                GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
            AS [%$##@_Alias])  
        AS H ON H.单位=G.单位)  AS I 
        LEFT JOIN 
        (SELECT * FROM 
            (SELECT 五险明细表.单位, Sum(五险明细表.单位缴纳) AS 失业单位 FROM 五险明细表 WHERE (((五险明细表.险种)="失业")) 
            GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
        AS [%$##@_Alias])  
    AS J ON J.单位=I.单位)  AS K 
    LEFT JOIN 
    (SELECT * FROM (SELECT 五险明细表.单位, Sum(五险明细表.单位缴纳) AS 工伤单位 FROM 五险明细表 WHERE (((五险明细表.险种)="工伤")) 
    GROUP BY 五险明细表.单位, 五险明细表.单位编号)  
AS [%$##@_Alias])  
AS L ON L.单位=K.单位)  AS M 
LEFT JOIN (SELECT * FROM (SELECT 五险明细表.单位, Sum(五险明细表.单位缴纳) AS 生育单位 
FROM 五险明细表 WHERE (((五险明细表.险种)="生育")) GROUP BY 五险明细表.单位, 五险明细表.单位编号)  AS [%$##@_Alias])  
AS N ON M.单位 = N.单位
ORDER BY M.单位编号;