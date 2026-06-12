SELECT
  T_ProfilKeping.KepingAutoID,
  T_ProfilKeping.KepingID,
  T_ProfilKeping.BlankoKepingAutoID,
  T_ProfilKeping.DeskripsiProfilKeping,
  T_ProfilKeping.PolaLubangID,
  T_ProfilKeping.RivetID,
  T_ProfilKeping.[Indikator-Lis],
  T_ProfilKeping.Coak,
  T_ProfilKeping.Koak,
  T_ProfilKeping.UkuranLubangID,
  T_ProfilKeping.Catatan,
  T_ProfilKeping.LubangFlag,
  T_ProfilKeping.TebalTipisFlag,
  T_BlankoKeping.BlankoKepingID,
  T_BlankoKeping.DeskripsiBlankoKeping,
  T_BlankoKeping.Panjang,
  T_BlankoKeping.Lebar,
  T_BlankoKeping.Tebal,
  T_BlankoKeping.Tebal2,
  T_BlankoKeping.Panjang2,
  T_BlankoKeping.DiameterDrum,
  T_UkuranRivet.KodeRivet,
  T_UkuranRivet.Ukuran,
  T_UkuranPolaLubang.PolaLubang,
  T_UkuranPolaLubang.DetailPolaLubang,
  T_ListPolaLubang.LokasiImage,
  T_ListPolaLubang.JumlahLubang,
  T_BlankoKeping.DiameterBrakeShoe,
  T_ProfilKeping.ImageCode,
  T_ProfilKeping.NoteGambar,
  GetPKImagePath ([T_ProfilKeping.KepingAutoID]) AS PKImageCode,
  GetHPImagePath (T_UkuranPolaLubang.PolaLubang) AS HPImageCode,
  T_UkuranLubang.UkuranLubang
FROM
  T_UkuranRivet
  RIGHT JOIN (
    (
      T_ListPolaLubang
      RIGHT JOIN T_UkuranPolaLubang ON T_ListPolaLubang.PolaLubang = T_UkuranPolaLubang.PolaLubang
    )
    RIGHT JOIN (
      T_UkuranLubang
      RIGHT JOIN (
        T_BlankoKeping
        RIGHT JOIN T_ProfilKeping ON T_BlankoKeping.BlankoKepingAutoID = T_ProfilKeping.BlankoKepingAutoID
      ) ON T_UkuranLubang.UkuranLubangID = T_ProfilKeping.UkuranLubangID
    ) ON T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID
  ) ON T_UkuranRivet.RivetID = T_ProfilKeping.RivetID;
