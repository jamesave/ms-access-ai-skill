SELECT
  T_ProfilKeping.BlankoKepingAutoID,
  T_ProfilKeping.[DeskripsiProfilKeping],
  T_BlankoKeping.DeskripsiBlankoKeping,
  T_BlankoKeping.Panjang,
  T_BlankoKeping.Lebar,
  T_BlankoKeping.Tebal,
  T_BlankoKeping.DiameterDrum,
  T_BlankoKeping.DiameterBrakeShoe,
  T_BlankoKeping.CatatanBlankoKeping
FROM
  T_BlankoKeping
  RIGHT JOIN T_ProfilKeping ON T_BlankoKeping.BlankoKepingAutoID = T_ProfilKeping.BlankoKepingAutoID;
