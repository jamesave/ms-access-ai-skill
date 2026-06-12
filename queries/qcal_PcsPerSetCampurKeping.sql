SELECT
  DISTINCT T_PcsSetBOM.BaseProdukID,
  DSum(
    "Qty", "T_PcsSetBOM", "BaseProdukID=" & [T_PcsSetBOM]![BaseProdukID]
  ) AS [pcs/set],
  DCount(
    "KepingID", "T_PcsSetBOM", "BaseProdukID=" & [T_PcsSetBOM]![BaseProdukID]
  ) AS QtyJenisKeping,
  IIf([QtyJenisKeping] = 1, 0, 1) AS CampurKepingFlagCalc
FROM
  T_PcsSetBOM;
