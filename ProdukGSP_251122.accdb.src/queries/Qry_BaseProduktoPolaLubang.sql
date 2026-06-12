SELECT
  T_PcsSetBOM.BaseProdukID,
  T_PcsSetBOM.KepingID,
  T_ProfilKeping.BlankoKepingAutoID,
  T_ProfilKeping.DeskripsiProfilKeping,
  T_ProfilKeping.RivetID,
  T_ProfilKeping.[Indikator-Lis],
  T_ProfilKeping.Coak,
  T_ProfilKeping.Koak,
  T_ProfilKeping.Catatan
FROM
  T_ProfilKeping
  INNER JOIN T_PcsSetBOM ON T_ProfilKeping.KepingID = T_PcsSetBOM.KepingID;
