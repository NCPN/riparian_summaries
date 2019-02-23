CREATE TABLE [tbl_wrk_Richness_Wetland] (
  [Unit_Code] VARCHAR (255),
  [Stream_Name] VARCHAR (50),
  [Visit_Year] SHORT ,
  [Plot_ID] SHORT ,
  [OBL] LONG ,
  [FACW] LONG ,
  [FAC] SHORT ,
  [FACU] SHORT ,
  [UPL] SHORT ,
  [CULT] SHORT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Unit_Code], [Plot_ID])
)
