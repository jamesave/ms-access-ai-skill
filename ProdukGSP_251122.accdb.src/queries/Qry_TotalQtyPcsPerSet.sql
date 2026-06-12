SELECT
  T_BaseProduk.SetAutoID,
  T_ListCustomer.NamaCustomer,
  T_BaseProduk.KategoriProdukID,
  T_BaseProduk.Deskripsi,
  T_BaseProduk.AxlePosID,
  Sum(T_PcsSetBOM.Qty) AS SumOfQty
FROM
  T_ListCustomer
  INNER JOIN (
    (
      T_PcsSetBOM
      INNER JOIN T_BaseProduk ON T_PcsSetBOM.SetAutoID = T_BaseProduk.SetAutoID
    )
    INNER JOIN T_Produk ON (
      T_BaseProduk.SetAutoID = T_Produk.SetAutoID
    )
    AND (
      T_PcsSetBOM.SetAutoID = T_Produk.SetAutoID
    )
  ) ON T_ListCustomer.CustomerID = T_Produk.CustomerID
GROUP BY
  T_BaseProduk.SetAutoID,
  T_ListCustomer.NamaCustomer,
  T_BaseProduk.KategoriProdukID,
  T_BaseProduk.Deskripsi,
  T_BaseProduk.AxlePosID
ORDER BY
  T_ListCustomer.NamaCustomer,
  T_BaseProduk.KategoriProdukID;
