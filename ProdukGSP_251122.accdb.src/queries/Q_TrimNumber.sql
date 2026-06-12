SELECT
  CInt(
    Mid([SetID], 2, 4)
  ) AS Expr1
FROM
  T_PcsSetBOM
ORDER BY
  CInt(
    Mid([SetID], 2, 4)
  );
