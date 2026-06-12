SELECT
  DISTINCT T_BaseProduk.KategoriProdukID,
  T_KategoriProduk.KategoriProduk
FROM
  T_KategoriProduk
  LEFT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID;
