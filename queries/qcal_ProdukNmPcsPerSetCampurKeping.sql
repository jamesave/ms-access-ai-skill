SELECT
  qcal_PcsPerSetCampurKeping.BaseProdukID,
  qcal_PcsPerSetCampurKeping.[pcs/set],
  qcal_PcsPerSetCampurKeping.QtyJenisKeping,
  qcal_PcsPerSetCampurKeping.CampurKepingFlagCalc,
  T_BaseProduk.BaseProdukName
FROM
  qcal_PcsPerSetCampurKeping
  INNER JOIN T_BaseProduk ON qcal_PcsPerSetCampurKeping.BaseProdukID = T_BaseProduk.BaseProdukID;
