SELECT
  T_ProfilKeping.KepingID,
  [T_ProfilKeping].BlankoKepingID AS Expr1,
  T_UkuranPolaLubang.PolaLubangID,
  T_UkuranPolaLubang.PolaLubang,
  T_ProfilKeping.DeskripsiProfilKeping,
  T_UkuranPolaLubang.Catatan
FROM
  T_UkuranPolaLubang
  LEFT JOIN T_ProfilKeping ON T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID;
