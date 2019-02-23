CREATE TABLE [tbl_wrk_Cover_Pct_Lifeform] (
  [UnitCode] VARCHAR (4),
  [Stream_Name] VARCHAR (50),
  [Visit_Year] SHORT ,
  [PlotID] SHORT ,
  [TreeL] DOUBLE ,
  [TreeA] DOUBLE ,
  [ShrubL] DOUBLE ,
  [ShrubA] DOUBLE ,
  [PGrassL] DOUBLE ,
  [PGrassA] DOUBLE ,
  [AGrassL] DOUBLE ,
  [AGrassA] DOUBLE ,
  [ForbL] DOUBLE ,
  [ForbA] DOUBLE ,
  [FernL] DOUBLE ,
  [FernA] DOUBLE ,
  [VineL] DOUBLE ,
  [VineA] DOUBLE ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([UnitCode], [Stream_Name], [Visit_Year], [PlotID])
)