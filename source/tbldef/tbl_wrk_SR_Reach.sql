CREATE TABLE [tbl_wrk_SR_Reach] (
  [Unit_Code] VARCHAR (4),
  [Stream_Name] VARCHAR (50),
  [Plot_ID] SHORT ,
  [Visit_Year] SHORT ,
  [Species_Code] VARCHAR (15),
  [Species_Name] VARCHAR (50),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Unit_Code], [Plot_ID], [Species_Code])
)
