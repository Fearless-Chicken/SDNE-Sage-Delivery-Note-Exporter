SELECT 
    CASE [BIJOU].[dbo].[F_DOCENTETE].[DO_Type] 
        WHEN 3 THEN 'BL'
        ELSE 'AUTRE'
    END AS [Identifiant du segment],
	[BIJOU].[dbo].[F_DOCENTETE].[DO_Piece] AS [Ref BL],
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Tiers] AS [Ref client],
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Date] AS [Date de départ],
    [BIJOU].[dbo].[F_COMPTET].[CT_Intitule] AS [Nom destinataire],
    [BIJOU].[dbo].[F_COMPTET].[CT_Adresse] AS [Libellé de voie],
    [BIJOU].[dbo].[F_COMPTET].[CT_CodePostal] AS [Code postal],
    [BIJOU].[dbo].[F_COMPTET].[CT_Ville] AS [Localité],
    contact.[CT_TelPortable] AS [Mobile],
    contact.[CT_EMail] AS [Mail]
FROM 
    [BIJOU].[dbo].[F_DOCENTETE] 
JOIN 
    [BIJOU].[dbo].[F_COMPTET] 
    ON [BIJOU].[dbo].[F_DOCENTETE].[DO_Tiers] = [BIJOU].[dbo].[F_COMPTET].[CT_Num]
OUTER APPLY (
    SELECT TOP 1 *
    FROM [BIJOU].[dbo].[F_CONTACTT]
    WHERE [F_CONTACTT].[CT_Num] = [F_DOCENTETE].[DO_Tiers]
) AS contact
WHERE
	[BIJOU].[dbo].[F_DOCENTETE].[Exporté] IS NULL
AND
	[BIJOU].[dbo].[F_DOCENTETE].[DO_Date] BETWEEN
		CAST('{dateDeb}T00:00:00.000' AS datetime)
		AND
		CAST('{dateFin}T23:59:59.997' AS datetime)
AND
	[BIJOU].[dbo].[F_DOCENTETE].[DO_Type] = 3
    ORDER BY 
        [BIJOU].[dbo].[F_DOCENTETE].[DO_Date] DESC
;
