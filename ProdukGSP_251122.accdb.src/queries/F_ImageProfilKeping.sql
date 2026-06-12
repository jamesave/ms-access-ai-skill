SELECT
  T_ProfilKeping.KepingAutoID,
  T_ProfilKeping.KepingID,
  T_ProfilKeping.DeskripsiProfilKeping,
  T_ProfilKeping.ImageCode,
  T_ProfilKeping.NoteGambar,
  GetPKImagePath ([KepingAutoID]) AS ImagePath
FROM
  T_ProfilKeping;
