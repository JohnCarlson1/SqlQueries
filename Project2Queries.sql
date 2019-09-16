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


-- Query 1

Select		tTest.ReportID,
			ReportDate,
			CompleteDate,
			ProblemDescr 'ProblemDescription',
			tTest.TestID,
			tTest.TestDate
From		tProblemReport
Inner Join	tTest
on			tProblemReport.ReportID = tTest.ReportID
Where		tTest.TestDate < ReportDate



-- Query 2

Select		Convert(Varchar(12), ReportDate, 107) 'ReportDateOutput',
			ReportID,
			tToy.SerialNO,
			ISNULL(Convert(Varchar(12), CompleteDate, 107), 'Not Complete') 'CompleteDate',
			DateDiff(day, CompleteDate, GetDate()) AS DaysInSystem,
			tModel.ModelNO,
			tModel.ModelName,
			tPerson.Lastname + ',' + substring(tperson.FirstName, 1, 1) + '.',
			CASE
				WHEN tPerson.PersonType = 'C' THEN 'Customer'
				WHEN tPerson.PersonType = 'E' THEN 'Employee'
				ELSE 'Distributor'
			END 'PersonType',
			tProblemType.TypeDescr
From		tProblemReport
Inner Join	tPerson
On			tProblemReport.PersonID = tPerson.PersonID
Inner Join	tToy
on			tProblemReport.ProblemReportSerialNO = tToy.SerialNO
Inner Join	tModel
on			tToy.ModelNO = tModel.ModelNO
Inner Join  tProblemType
on			tProblemReport.ProblemType = tproblemType.ProblemType
Where		month(ReportDate) = 10
Order By	ReportDate

-- Query 3

Select		Convert(Varchar(12), ReportDate, 107) 'ReportDateOutput',
			tProblemReport.ReportID,
			tToy.SerialNO,
			Convert(Varchar(12), CompleteDate, 107) 'CompleteDate',
			DateDiff(day, CompleteDate, GetDate()) AS DaysInSystem,
			tModel.ModelNO,
			tModel.ModelName,
			Reporter.Lastname + ',' + substring(Reporter.FirstName, 1, 1) + '.' reportername,
			CASE
				WHEN Reporter.PersonType = 'C' THEN 'Customer'
				WHEN Reporter.PersonType = 'E' THEN 'Employee'
				ELSE 'Distributor'
			END 'PersonType',
			Convert(Varchar(12), tTest.TestDate, 107),
			tTest.TestDescr,
			Tester.LastName + ',' + substring(Tester.FirstName, 1, 1) + '.' testername,
			tTest.TestComp
From		tProblemReport
Inner Join	tToy
on			tProblemReport.ProblemReportSerialNO = tToy.SerialNO
Inner Join	tModel
on			tToy.ModelNO = tModel.ModelNO
left outer Join	tPerson Reporter
On			tProblemReport.PersonID = Reporter.PersonID
Left Outer Join	tTest
on			tProblemReport.ReportID = tTest.ReportID
Left outer Join	tPerson Tester
On			tTest.PersonID = Tester.PersonID
Where		month(ReportDate) = 10
Order By ReportDateOutput

-- Query 4
CREATE VIEW vCountTest 
as SELECT      
			reportID,
			Count(TestID) 'CountofTest'
FROM		tTest	  
Group By	ReportID;

Select		Convert(Varchar(12), ReportDate, 107) 'ReportDateOutput',
			tProblemReport.ReportID,
			tToy.SerialNO,
			Convert(Varchar(12), CompleteDate, 107) 'CompleteDate',
			CASE
			WHEN CompleteDate IS NULL
			THEN cast(GETDATE() as int) - cast(ReportDate as int)
			ELSE cast(GETDATE() as int) - cast(CompleteDate as int)
			END as 'DaysInSystem',
			tModel.ModelNO,
			tModel.ModelName,
			tPerson.Lastname + ',' + substring(tperson.FirstName, 1, 1) + '.' as 'Reporter Name',
			CASE
				WHEN tPerson.PersonType = 'C' THEN 'Customer'
				WHEN tPerson.PersonType = 'E' THEN 'Employee'
				ELSE 'Distributor'
			END 'PersonType',
			isnull(CountofTest, 0)  'CountOfTest'
From		tProblemReport
Inner Join	tPerson
On			tProblemReport.PersonID = tPerson.PersonID
Inner Join	tToy
on			tProblemReport.ProblemReportSerialNO = tToy.SerialNO
Inner Join	tModel
on			tToy.ModelNO = tModel.ModelNO
Inner Join  tProblemType
on			tProblemReport.ProblemType = tproblemType.ProblemType
left outer join vCountTest
on			tProblemReport.ReportID = vCountTest.ReportID
ORDER by    ReportDate




-- Query 5
-- View to generate Max Test Counts
CREATE VIEW vMaxTest2 
as SELECT      
			reportID,
			Count(TestID) 'CountofTest',
			isnull(datediff(day,getdate(),TestDate), 0) 'DaysInSystem',
			isnull(convert(varchar(12), TestDate, 107), 'Not Complete') 'Complete Date'
FROM		tTest	  
Group By	ReportID, TestDate;

select * from vMaxTest2

Select		Convert(Varchar, ReportDate, 107) 'ReportDateOutput',
			tProblemReport.ReportID,
			tToy.SerialNO,
			vMaxTest2.[Complete Date],
			CASE
			WHEN CompleteDate IS NULL
			THEN cast(GETDATE() as int) - cast(ReportDate as int)
			ELSE cast(GETDATE() as int) - cast(CompleteDate as int)
			END as 'DaysInSystem',
			tModel.ModelNO,
			tModel.ModelName,
			tPerson.Lastname + ',' + substring(tperson.FirstName, 1, 1) + '.' as 'Reporter Name',
			CASE
				WHEN tPerson.PersonType = 'C' THEN 'Customer'
				WHEN tPerson.PersonType = 'E' THEN 'Employee'
				ELSE 'Distributor'
			END 'PersonType',
			vMaxTest2.CountofTest
From		tProblemReport
Inner Join	tTest
On			tProblemReport.ReportID = tTest.ReportID
Inner Join	tPerson
On			tProblemReport.PersonID = tPerson.PersonID
Inner Join	tToy
on			tProblemReport.ProblemReportSerialNO = tToy.SerialNO
Inner Join	tModel
on			tToy.ModelNO = tModel.ModelNO
Inner Join  tProblemType
on			tProblemReport.ProblemType = tproblemType.ProblemType
left outer join vMaxTest2
on			tProblemReport.ReportID = vMaxTest2.ReportID
WHERE		CountofTest = '4' AND TestDate = 'Not Complete'
ORDER by    ReportDate

-- Query 6
-- View to Count Reports
Create View	vReportCount As
Select		tmodel.ModelNO,
			modeldescr,
			tproblemreport.ProblemReportSerialNO,
			count(tToy.SerialNO) 'CountofReports'
From		tModel

Inner Join	tToy
On			tmodel.ModelNO = tToy.ModelNO
Left Outer Join tProblemReport 
On			tToy.serialNO = tProblemReport.ProblemReportSerialNO
Group By	tmodel.ModelNO, ModelDescr, tproblemreport.ProblemReportSerialNO, tProblemReport.ReportDate

--View to find the earliest and recent dates
Create View vDate AS
Select		tmodel.ModelNO,
			ModelDescr,
			Max(tProblemReport.ReportDate) 'MostRecentReportDate',
			Min(tProblemReport.ReportDate) 'EarliestReportDate',
			Max(tTest.TestDate) 'MostRecentTestDate',
			Min(tTest.TestDate) 'EarliestTestDate'
From		tModel
Left Outer Join	tToy
On			tmodel.ModelNO = tToy.modelNO
Left Outer Join	tProblemReport
on			tToy.SerialNO = tProblemReport.ProblemReportSerialNO
Left Outer Join	tTest
on			tTest.ReportID = tProblemReport.ReportID
Group By tmodel.ModelNO, ModelDescr

--View to distinguish Injury vs no Injury
Create View	vInjury As
Select		tmodel.ModelNO,
			tmodel.ModelDescr,
			tProblemReport.Injury,
			tProblemReport.InjuryDescr,
			Case
				When tProblemReport.Injury = 'Y'
				then 1
				else 0
			End As CountofInjuryReports,

			Count(tTest.ReportID) CountofTest
From tModel
Left Outer Join tToy
On				tModel.ModelNO = tToy.modelNO
Left Outer Join	tProblemReport
On				tToy.SerialNO = tProblemReport.ProblemReportSerialNO
Left Outer Join	tTest
On				tTest.ReportID = tProblemReport.ReportID
Group By		tModel.ModelNO, ModelDescr, tProblemReport.Injury, tProblemReport.InjuryDescr, tTest.ReportID

--View To Sum tests and Injury Reports
Create View		vSumInjury As
Select			ModelNO,
				ModelDescr,
				Sum(CountofTest) CountofTest,
				Sum(CountofInjuryReports) CountofInjuryReports
From			vInjury
Group By		ModelNO, ModelDescr

--Views to Count Reports
Create View vSumReport As
Select		ModelNO,
			sum(CountofReports) CountofReports
From		vReportCount
Group By	ModelNO


-- Query 6 
Select		tModel.ModelNO,
			tModel.ModelDescr,
			ISNULL(vSumReport.CountofReports, 0) CountofReports,
			vSumInjury .CountofInjuryReports,
			ISNULL(Convert(Varchar, vDate.MostRecentReportDate, 107), 'n/a') MostRecentReportDate,
			ISNULL(Convert(Varchar, vDate.EarliestReportDate, 107), 'n/a') EarliestReportDate,
			vSumInjury .CountofTest,
			ISNULL(Convert(Varchar, vDate.MostRecentTestDate, 107), 'n/a') MostRecentTestDate,
			ISNULL(Convert(Varchar, vDate.EarliestTestDate, 107), 'n/a') EarliestTestDate
From		tModel
Left Outer Join	vSumInjury 
On			tModel.ModelNO = vSumInjury.ModelNO
Left Outer Join vSumReport
On			tModel.ModelNO = vSumReport.ModelNO
Left Outer Join vDate
On			tmodel.ModelNO = vDate.ModelNO

--Query 7
Select		tModel.ModelNO,
			tModel.ModelDescr,
			ISNULL(vSumReport.CountofReports, 0) CountofReports,
			vSumInjury.CountofInjuryReports,
			ISNULL(Convert(Varchar, vDate.MostRecentReportDate, 107), 'n/a') MostRecentReportDate,
			ISNULL(Convert(Varchar, vDate.EarliestReportDate, 107), 'n/a') EarliestReportDate,
			vSumInjury.CountofTest,
			ISNULL(Convert(Varchar, vDate.MostRecentTestDate, 107), 'n/a') MostRecentTestDate,
			ISNULL(Convert(Varchar, vDate.EarliestTestDate, 107), 'n/a') EarliestTestDate
From		tModel
Left Outer Join	vSumInjury 
On			tModel.ModelNO = vSumInjury.ModelNO
Left Outer Join vSumReport
On			tModel.ModelNO = vSumReport.ModelNO
Left Outer Join vDate
On			tmodel.ModelNO = vDate.ModelNO
Where		vSumReport.CountofReports = (select Max(CountofReports) From vSumReport)

-- Query 8
Select		tModel.ModelNO 'ModelNO',
			tmodel.ModelName 'ModelName',
			tProblemReport.ReportID 'ReportID',
			Convert(Varchar, tProblemReport.ReportDate, 110) 'ReportDate',
			tProblemReport.ProblemDescr 'ProblemDescription',
			tProblemType.TypeDescr 'TypeDescription',
			tPerson.PersonID 'ReportingPerson',
			tTest.TestID 'TestID',
			Convert(Varchar, tTest.TestDate, 110) 'TestDate',
			tTest.TestDescr 'TestDescription',
			tPerson.Lastname + ',' + substring(tperson.FirstName, 1, 1) + '.' as 'TestingPerson'
From		tModel
Left Outer Join	tToy
On			tmodel.ModelNO = tToy.ModelNO
Left Outer Join tProblemReport
On			tToy.SerialNO = tProblemReport.ProblemReportSerialNO
Left Outer Join tProblemType
On			tProblemType.ProblemType = tProblemReport.ProblemType
Left Outer Join tTest
On			tTest.ReportID = tProblemReport.ReportID
Left Outer join tPerson
On			tPerson.PersonID = tProblemReport.PersonID
Left Outer Join tPerson p
On			p.PersonID = tTest.TesterID
Where		tProblemReport.ProblemDescr Like '%battery%' 
AND			tProblemReport.Injury = 'yes'

-- Question 9	
-- View to count reports of customers
Create View vReportCount2 As
Select		tPerson.PersonType,
			Count(ReportID) CountofReports
From		tProblemReport
Inner Join tPerson
on			tperson.PersonID = tProblemReport.ReportID
Group By	tperson.PersonType;

Select		Case
				When tPerson.PersonType = 'c'
				Then 'Customer'
				When tPErson.PersonType = 'D'
				Then 'Distributor'
				When tPerson.PersonType = 'E'
				Then 'Employee'
			End Reporter,
			CountofReports
From		vReportCount2
Inner Join  tPerson
On			tPerson.PersonID = vReportCount2.CountofReports
Where		CountofReports =
			(Select Max(CountofReports)
			From vReportCount2)


-- Query 10
 
Select           	tModel.ModelNO 'ModelNumber',
                     tModel.ModelName 'ModelName',
                     tModel.ModelDescr 'ModelDescription',
                     tModel.StandardPrice 'StandardPrice'
From             	tModel
where            	tModel.ModelNO NOT IN (Select ModelNO from tProblemReport);


			
			



	











