SELECT
  T_BaseProduk.BaseProdukName,
  T_PcsSetBOM.KepingID,
  T_PcsSetBOM.Qty
FROM
  T_BaseProduk
  LEFT JOIN T_PcsSetBOM ON T_BaseProduk.BaseProdukID = T_PcsSetBOM.BaseProdukID;
