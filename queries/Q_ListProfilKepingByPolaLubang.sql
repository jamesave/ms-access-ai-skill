SELECT
  T_ProfilKeping.*,
  T_UkuranPolaLubang.PolaLubang
FROM
  T_UkuranPolaLubang
  RIGHT JOIN T_ProfilKeping ON T_UkuranPolaLubang.PolaLubangID = T_ProfilKeping.PolaLubangID
WHERE
  (
    (
      (T_UkuranPolaLubang.PolaLubang)= "12-A"
    )
  );
