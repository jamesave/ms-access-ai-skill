SELECT
  T_PcsSetBOM.SetAutoID,
  T_PcsSetBOM.SetID,
  T_PcsSetBOM.Qty,
  T_PcsSetBOM.KategoriProdukID,
  T_PcsSetBOM.KepingID,
  T_DataKeping.DeskripsiKeping,
  T_DataKeping.Lebar_deprecated,
  T_DataKeping.Panjang_deprecated,
  T_DataKeping.Tebal_deprecated,
  T_DataKeping.DiameterDrum_deprecated
FROM
  T_DataKeping
  RIGHT JOIN T_PcsSetBOM ON T_DataKeping.KepingID = T_PcsSetBOM.KepingID;
