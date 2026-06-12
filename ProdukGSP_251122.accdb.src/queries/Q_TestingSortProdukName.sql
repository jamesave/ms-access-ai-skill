SELECT
  T_BaseProduk.*,
  ExtractMiddleID([BaseProdukName]) AS SortKey
FROM
  T_BaseProduk
ORDER BY
  ExtractMiddleID([BaseProdukName]);
