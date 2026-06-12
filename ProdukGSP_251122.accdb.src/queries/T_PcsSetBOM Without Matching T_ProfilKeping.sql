SELECT
  T_PcsSetBOM.*
FROM
  T_PcsSetBOM
  LEFT JOIN T_ProfilKeping ON T_PcsSetBOM.[KepingID] = T_ProfilKeping.[KepingID]
WHERE
  (
    (
      (T_ProfilKeping.KepingID) Is Null
    )
  );
