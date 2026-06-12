SELECT
  T_BlankoKeping.BlankoKepingID,
  T_BlankoKeping.DeskripsiBlankoKeping,
  T_ProfilKeping.KepingID,
  T_ProfilKeping.[DeskripsiProfilKeping],
  T_ProfilKeping.[DeskripsiProfilKeping]
FROM
  T_BlankoKeping
  RIGHT JOIN T_ProfilKeping ON T_BlankoKeping.BlankoKepingID = T_ProfilKeping.BlankoKepingID;
