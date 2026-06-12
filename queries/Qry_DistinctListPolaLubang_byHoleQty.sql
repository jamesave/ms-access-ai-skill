SELECT
  DISTINCT T_ListPolaLubang.JumlahLubang,
  T_ListPolaLubang.PolaLubang,
  T_ListPolaLubang.LokasiImage
FROM
  T_ListPolaLubang
  INNER JOIN T_PolaLubang ON T_ListPolaLubang.PolaLubang = T_PolaLubang.PolaLubang
ORDER BY
  T_ListPolaLubang.JumlahLubang;
