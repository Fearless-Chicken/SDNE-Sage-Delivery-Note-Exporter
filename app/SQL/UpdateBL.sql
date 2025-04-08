UPDATE 
    [BIJOU].[dbo].[F_DOCENTETE]
SET 
	[BIJOU].[dbo].[F_DOCENTETE].[Exporté] = 'Oui',
	[BIJOU].[dbo].[F_DOCENTETE].[Date d'export] = '{todaydate}'
WHERE 
    [BIJOU].[dbo].[F_DOCENTETE].[DO_Piece] = '{BL}';