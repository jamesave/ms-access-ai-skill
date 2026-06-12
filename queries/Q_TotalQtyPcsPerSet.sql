SELECT
  T_PcsSetBOM.SetID,
  T_PcsSetBOM.Deskripsi_deprecated,
  Sum(T_PcsSetBOM.Qty) AS SumOfQty,
  T_PcsSetBOM.KategoriProdukID_deprecated
FROM
  T_ProfilKeping
  INNER JOIN T_PcsSetBOM ON [T_ProfilKeping].KepingID = T_PcsSetBOM.KepingID
GROUP BY
  T_PcsSetBOM.SetID,
  T_PcsSetBOM.Deskripsi_deprecated,
  T_PcsSetBOM.KategoriProdukID_deprecated;
