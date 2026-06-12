SELECT
  T_BaseProduk.BaseProdukID,
  T_BaseProduk.BaseProdukName,
  T_BaseProduk.Deskripsi,
  T_BaseProduk.Catatan,
  T_ListStatusBaseProduk.StatusProduk,
  T_KategoriProduk.KategoriProduk,
  T_ListAxlePos.AxlePos,
  Val(
    Mid([BaseProdukName], 2, 4)
  ) AS SortedBaseProdukName,
  T_KategoriProduk.KategoriProdukID,
  ExtractMiddleID([BaseProdukName]) AS SortKey
FROM
  T_ListStatusBaseProduk
  RIGHT JOIN (
    T_ListAxlePos
    RIGHT JOIN (
      T_KategoriProduk
      RIGHT JOIN T_BaseProduk ON T_KategoriProduk.KategoriProdukID = T_BaseProduk.KategoriProdukID
    ) ON T_ListAxlePos.AxlePosID = T_BaseProduk.AxlePosID
  ) ON T_ListStatusBaseProduk.StatusProdukID = T_BaseProduk.StatusProdukID
ORDER BY
  ExtractMiddleID([BaseProdukName]);
