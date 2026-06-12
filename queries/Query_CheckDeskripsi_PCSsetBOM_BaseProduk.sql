SELECT
  T_BaseProduk.SetAutoID,
  T_PcsSetBOM.SetAutoID,
  T_PcsSetBOM.SetID,
  T_PcsSetBOM.Deskripsi_deprecated,
  T_BaseProduk.Deskripsi
FROM
  (
    T_PcsSetBOM
    INNER JOIN T_BaseProduk ON T_PcsSetBOM.SetAutoID = T_BaseProduk.SetAutoID
  )
  INNER JOIN T_ProdukOEM ON T_PcsSetBOM.SetAutoID = T_ProdukOEM.SetID
ORDER BY
  T_BaseProduk.SetAutoID;
