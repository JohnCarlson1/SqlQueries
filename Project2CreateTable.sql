-- Create Tables

-- Toy Table
Create Table tToy
(
 serialNO Varchar(15) Primary Key,
 modelNO varchar(10) references tModel (ModelNO),
 pricepaid money
);

drop table tToy

Insert Into tToy (SerialNO, ModelNO, PricePaid)
Select Distinct SerialNumber, ToyModelNumber, PricePaid
From tReplicaAll
Where SerialNumber is not null

Select Distinct *
From tToy
Where SerialNO is not null
Order By SerialNo


-- Test Table
CREATE TABLE tTest
(
 TestID		int Primary Key,
 TestDate	datetime,
 TestDescr varchar(255),
 TestResult varchar(255),
 TestComp varchar(1),
 PersonID	int references tPerson (PersonID),
 ReportID	int references tProblemReport (ReportID),
 TesterID int references tTest (TestID)
 );

 drop table tTest

 Insert Into tTest (TestID, TestDate, TestDescr, TestResult, TestComp, PersonID, ReportID, TesterID)
 Select Distinct TestID, TestDate, TestDescription, TestResults, TestComplete, PersonID, problemreportreportID, TesterID
 From tReplicaAll
 Where TestID is not null

 Select Distinct *
 from tTest
 Where TestID is not null
 Order By TestID

 -- Problem Report Table
CREATE TABLE tProblemReport
(
 ReportID        	int Primary Key,
 ReportDate			DateTime,
 CompleteDate		DateTime,
 ProblemDescr		Varchar(255),
 Injury				varchar(3),
 InjuryDescr		varchar(255),
 ProblemReportSerialNO varchar(10),
 ProblemType		char(1)
 REFERENCES			tProblemType (ProblemType),
 PersonID			Int 
 References			tPerson (PersonID)
);

drop table tProblemReport

Insert Into tProblemReport (ReportID, ReportDate, CompleteDate, ProblemDescr, Injury, InjuryDescr, ProblemReportSerialNO, ProblemType, PersonID)
Select Distinct		ProblemReportreportID, ReportDate, CompleteDate, ProblemDescription, InjuryYN, InjuryDescription, problemreportserialnumber, ProblemTypeID, PersonID
From		tReplicaAll
Where ProblemReportreportID is not null

Select Distinct *
from tProblemReport
where ReportID is not null
Order By ReportID

-- ProblemType Table

Create Table tProblemType
(
 Problemtype char(1) Primary Key,
 TypeDescr  varchar(255)
);

Drop Table tProblemType

Insert Into tProblemType (ProblemType, TypeDescr)
Select distinct ProblemTypeID, TypeDescription
From treplicaall
where ProblemTypeID is not null

Select Distinct *
from tProblemType
Where Problemtype is not null
Order By ProblemType
 


-- Model Table
 Create Table tModel
 (
  ModelNO			varchar(10) Primary Key,
  ModelName			varchar(255),
  ModelDescr		varchar(255),
  StandardPrice		money
);

drop table tmodel

Insert into tModel (ModelNO, ModelName, ModelDescr, StandardPrice)
Select Distinct ModelNumber, ModelName, ModelDescription, StandardPrice		
from		treplicaall
Where ModelNumber is not null

Select *
from		tmodel
where		ModelNO is not null
ORder by	ModelNO;


-- Person Table
Create Table tPerson
(
 PersonID			int Primary Key,
 LastName			varchar(20),
 FirstName			varchar(20),
 Phone				char(20),
 PersonType			varchar(1)
);

drop table tperson

 Insert Into tPerson (PersonID, LastName, FirstName, Phone, PersonType)
 Select Distinct PersonID,
		LastName,
		FirstName,
		Phone,
		PersonType
From	tReplicaAll
Where PersonID is not null

Select Distinct * 
From tperson
where PersonID is not null
Order by PersonID 

