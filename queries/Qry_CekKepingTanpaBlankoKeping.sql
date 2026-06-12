SELECT
  T_PcsSetBOM.BaseProdukID,
  T_PcsSetBOM.KepingID,
  T_ProfilKeping.BlankoKepingID,
  T_ProfilKeping.[DeskripsiProfilKeping]
FROM
  T_ProfilKeping
  RIGHT JOIN T_PcsSetBOM ON T_ProfilKeping.KepingID = T_PcsSetBOM.KepingID;
