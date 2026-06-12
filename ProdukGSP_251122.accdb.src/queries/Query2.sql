SELECT
  T_BaseProduk.*,
  ExtractMiddleID([BaseProdukName]) AS SortID
FROM
  T_BaseProduk
ORDER BY
  ExtractMiddleID([BaseProdukName]);
