SELECT
  T_ProfilKeping.KepingAutoID,
  T_ProfilKeping.KepingID,
  T_ProfilKeping.BlankoKepingAutoID,
  T_ProfilKeping.[DeskripsiProfilKeping],
  T_ProfilKeping.PolaLubangID,
  T_ProfilKeping.RivetID,
  T_ProfilKeping.[Indikator-Lis],
  T_ProfilKeping.Coak,
  T_ProfilKeping.Koak,
  T_ProfilKeping.Catatan,
  T_BlankoKeping.DeskripsiBlankoKeping,
  T_BlankoKeping.Panjang,
  T_BlankoKeping.Lebar,
  T_BlankoKeping.Tebal_TipisFlag,
  T_BlankoKeping.Tebal,
  T_BlankoKeping.DiameterDrum,
  T_BlankoKeping.Tebal2,
  T_BlankoKeping.DiameterBrakeShoe,
  T_UkuranRivet.KodeRivet,
  T_UkuranRivet.[Ukuran],
  T_UkuranPolaLubang.PolaLubang,
  T_UkuranPolaLubang.PolaLubangID,
  T_ListPolaLubang.LokasiImage,
  T_ListPolaLubang.JumlahLubang
FROM
  T_UkuranRivet
  RIGHT JOIN (
    T_ListPolaLubang
    RIGHT JOIN (
      T_BlankoKeping
      RIGHT JOIN (
        T_ProfilKeping
        LEFT JOIN T_UkuranPolaLubang ON T_ProfilKeping.PolaLubangID = T_UkuranPolaLubang.PolaLubangID
      ) ON T_BlankoKeping.BlankoKepingID = T_ProfilKeping.BlankoKepingID
    ) ON T_ListPolaLubang.PolaLubang = T_UkuranPolaLubang.PolaLubang
  ) ON T_UkuranRivet.RivetID = T_ProfilKeping.RivetID;
