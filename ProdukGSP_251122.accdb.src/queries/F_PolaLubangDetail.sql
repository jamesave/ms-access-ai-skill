SELECT
  T_UkuranPolaLubang.*,
  T_ListPolaLubang.LokasiImage,
  GetHPImagePath([T_ListPolaLubang.PolaLubang]) AS ImageCode
FROM
  T_ListPolaLubang
  RIGHT JOIN T_UkuranPolaLubang ON T_ListPolaLubang.PolaLubang = T_UkuranPolaLubang.PolaLubang
ORDER BY
  T_ListPolaLubang.JumlahLubang;
