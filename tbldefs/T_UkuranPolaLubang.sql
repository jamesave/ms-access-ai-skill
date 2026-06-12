CREATE TABLE [T_UkuranPolaLubang] (
  [PolaLubangID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [PolaLubang] VARCHAR (20) CONSTRAINT [T_ListPolaLubangT_UkuranPolaLubang] REFERENCES [T_ListPolaLubang] ([PolaLubang]),
  [DetailPolaLubang] VARCHAR (255),
  [L] VARCHAR (20),
  [X] VARCHAR (20),
  [A] VARCHAR (20),
  [B] VARCHAR (20),
  [C] VARCHAR (20),
  [D] VARCHAR (20),
  [E] VARCHAR (20),
  [F] VARCHAR (20),
  [G] VARCHAR (20),
  [H] VARCHAR (20)
)
