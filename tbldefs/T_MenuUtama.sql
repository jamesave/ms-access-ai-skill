CREATE TABLE [T_MenuUtama] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [TargetObject] VARCHAR (255),
  [ItemDescription] LONGTEXT,
  [DetailDescription] LONGTEXT,
  [SortOrder] SINGLE
)
