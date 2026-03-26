

-- Stored Procedure
--Create with input value
CREATE PROCEDURE usp_GetAllDepartments 
AS
BEGIN
	SELECT DepartmentID,
	Name,
	GroupName
	FROM HumanResources.Department
	ORDER BY DepartmentID ASC
END

EXEC usp_GetAllDepartments

------------------------------------------------------

CREATE PROCEDURE usp_GetDepartmentByID
	@DepartmentID INT
AS
BEGIN
	SELECT DepartmentID,
	[Name],
	GroupName
	FROM HumanResources.Department
	WHERE DepartmentID = @DepartmentID
END

EXEC usp_GetDepartmentByID
@DepartmentID = 9

------------------------------------------------------

CREATE PROCEDURE usp_GetDepartmentByIDAndGroup
	@DepartmentID INT,
	@GroupName NVARCHAR(50)
AS
BEGIN
	SELECT DepartmentID,
	[Name],
	GroupName
	FROM HumanResources.Department
	WHERE DepartmentID = @DepartmentID AND GroupName = @GroupName
END

EXEC usp_GetDepartmentByIDAndGroup 
@DepartmentID = 9, @GroupName = 'Executive General and Administration'

-----------------------------------------------------

CREATE PROCEDURE usp_GetDepartmentsByGroupName

	@GroupName NVARCHAR(50)
AS
BEGIN
	SELECT DepartmentID,
	[Name],
	GroupName
	FROM HumanResources.Department
	WHERE GroupName = @GroupName
	ORDER BY [Name] ASC
END

EXEC usp_GetDepartmentsByGroupName @GroupName = 'Research and Development'

-------------------------------------------------------------

CREATE PROCEDURE usp_SearchDepartmentsByName
	@NamePart NVARCHAR(50)
AS
BEGIN
	SELECT DepartmentID,
	[Name],
	GroupName
	FROM HumanResources.Department
	WHERE [Name] LIKE '%' + @NamePart + '%'
	ORDER BY DepartmentID
END

EXEC usp_SearchDepartmentsByName @NamePart = 'and'

DROP PROCEDURE usp_SearchDepartmentsByName

--------------------------------------------------------------

CREATE PROCEDURE usp_GetDepartmentsOptionalByGroup
	@GroupName  NVARCHAR(50) = NULL
AS
BEGIN 
	IF @GroupName IS NULL
	BEGIN 
		SELECT DepartmentID,
		[Name],
		GroupName
		FROM HumanResources.Department
	END
	ELSE
	BEGIN 
		SELECT DepartmentID,
		[Name],
		GroupName
		FROM HumanResources.Department
		WHERE GroupName = @GroupName
		ORDER BY DepartmentID
	END
END

EXEC usp_GetDepartmentsOptionalByGroup @GroupName = 'Manufacturing'

--------------------------------------------------------------------

CREATE PROCEDURE usp_CountDepartmentsByGroup
	@GroupName NVARCHAR(50)
AS
BEGIN
	SELECT COUNT(DepartmentID) AS DepartmentCount FROM HumanResources.Department
	WHERE GroupName = @GroupName
END

EXEC usp_CountDepartmentsByGroup @GroupName = 'Executive General and Administration'

--------------------------------------------------------------------

CREATE PROCEDURE usp_GetEmployeesByGender
	@Gender NCHAR(1)
AS
BEGIN
	IF @Gender = 'M'
	BEGIN
		SELECT BusinessEntityID,
		NationalIDNumber,
		JobTitle,
		BirthDate,
		Gender
		FROM HumanResources.Employee
		WHERE Gender = @Gender
		ORDER BY BusinessEntityID ASC
	END
	IF @Gender = 'F'
	BEGIN
	SELECT BusinessEntityID,
		NationalIDNumber,
		JobTitle,
		BirthDate,
		Gender
		FROM HumanResources.Employee
		WHERE Gender = @Gender
		ORDER BY BusinessEntityID ASC
	END
END

EXEC usp_GetEmployeesByGender @Gender = 'F'

----------------------------------------------------------------------------

CREATE PROCEDURE usp_GetEmployeesBornAfterDate
	@GivenDate DATE
AS
BEGIN
	SELECT BusinessEntityID,
	JobTitle,
	BirthDate,
	MaritalStatus,
	Gender
	FROM HumanResources.Employee
	WHERE BirthDate > @GivenDate
	ORDER BY BirthDate ASC
END

EXEC usp_GetEmployeesBornAfterDate @GivenDate = '1990-05-28'

----------------------------------------------------------------------------------

CREATE PROCEDURE usp_CountEmployeesByGender
    @Gender NCHAR(1),
    @EmpCount INT OUTPUT
AS
BEGIN
    SELECT @EmpCount = COUNT(*)
    FROM HumanResources.Employee
    WHERE Gender = @Gender
END


DECLARE @Result INT 
EXEC usp_CountEmployeesByGender
	@Gender = 'F',
	@EmpCount = @Result OUTPUT
SELECT @Result + COUNT(*) FROM Sales.SalesOrderDetail

-----------------------------------------------------------------------------------------------

CREATE PROCEDURE usp_AddVacationHours
	@BusinessEntityID INT,
	@HoursToAdd SMALLINT
AS
BEGIN
	SET NOCOUNT ON
	IF NOT EXISTS(SELECT 1 FROM HumanResources.Employee WHERE BusinessEntityID = @BusinessEntityID)
	BEGIN
		SELECT 'Employee not found' AS [Message]
		RETURN
	END
	IF @HoursToAdd <= 0
	BEGIN
		SELECT 'HoursToAdd must be greater than 0.' AS [Message]
		RETURN
	END
	IF EXISTS(SELECT 1 FROM HumanResources.Employee WHERE BusinessEntityID = @BusinessEntityID AND 
		(VacationHours + @HoursToAdd) > 240)
	BEGIN
        SELECT 'VacationHours cannot exceed 240.' AS [Message]
        RETURN
    END
	UPDATE HumanResources.Employee SET VacationHours = VacationHours + @HoursToAdd, 
	ModifiedDate = GETDATE()
	WHERE BusinessEntityID = @BusinessEntityID

	SELECT BusinessEntityID,
	JobTitle,
	VacationHours,
	ModifiedDate
	FROM HumanResources.Employee
	WHERE BusinessEntityID = @BusinessEntityID
END

EXEC usp_AddVacationHours @BusinessEntityID = 3, @HoursToAdd = 40

-----------------------------------------------------------------------------------

CREATE PROCEDURE usp_SubtractVacationHours
	@BusinessEntityID INT,
	@HoursToSubtract SMALLINT
AS
BEGIN 
	SET NOCOUNT ON

	IF NOT EXISTS(SELECT 1 FROM HumanResources.Employee WHERE BusinessEntityID=@BusinessEntityID)
	BEGIN
		SELECT 'There is not Employee' AS Message
		RETURN
	END

	IF @HoursToSubtract <= 0
	BEGIN
		SELECT 'Value is entered mistake' AS Message
		RETURN
	END
	IF(SELECT VacationHours FROM HumanResources.Employee WHERE BusinessEntityID = @BusinessEntityID) < @HoursToSubtract
	BEGIN
		SELECT 'Vac can not be more less than Hours' AS Message
		RETURN
	END
	UPDATE HumanResources.Employee SET VacationHours = VacationHours - @HoursToSubtract,
	ModifiedDate = GETDATE()
	WHERE BusinessEntityID = @BusinessEntityID

	SELECT BusinessEntityID,
	JobTitle,
	VacationHours,
	ModifiedDate
	FROM HumanResources.Employee
	WHERE BusinessEntityID = @BusinessEntityID

END

EXEC usp_SubtractVacationHours @BusinessEntityID = 1, @HoursToSubtract = 19


SELECT * FROM HumanResources.Employee
---------------------------------------------------------------------------------------------------

CREATE PROCEDURE usp_InsertDepartment
	@Name NVARCHAR(50),
	@GroupName NVARCHAR(50)
AS
BEGIN
	SET NOCOUNT ON

	IF @Name IS NULL OR @Name = ''
	BEGIN
		SELECT 'Name can not be NULL or Empty' AS Message
		RETURN
	END
	IF @GroupName IS NULL OR @GroupName = ''
	BEGIN
		SELECT 'GroupName can not be NULL or Empty' AS Message
		RETURN
	END
	IF EXISTS(SELECT 1 FROM HumanResources.Department WHERE [Name] = @Name OR GroupName = @GroupName)
	BEGIN
		SELECT 'Name or GroupName already exists' AS Message
		RETURN
	END

	INSERT INTO HumanResources.Department (Name, GroupName, ModifiedDate)
	VALUES (@Name, @GroupName, GETDATE())

	SELECT DepartmentID,
	Name,
	GroupName,
	ModifiedDate
	FROM HumanResources.Department
	WHERE [Name] = @Name 
END

EXEC usp_InsertDepartment @Name = 'Logistic', @GroupName = 'Freight'

------------------------------------------------------------------------------------------

CREATE FUNCTION DSFDS 
(
	@Num1 INT, 
	@Num2 INT
)
RETURNS INT 
AS
BEGIN
    DECLARE @Result INT
    SET @Result = @Num1 + @Num2
    RETURN @Result
END

-------------------------------------------------------------------------------------------------------
CREATE ROLE role_reportReaders
CREATE ROLE role_SalesEditors

CREATE USER UserA WITHOUT LOGIN
CREATE USER UserB WITHOUT LOGIN
CREATE USER UserC WITHOUT LOGIN

--Question 1
GRANT SELECT ON Person.Person TO UserA

--Question 2
GRANT SELECT, INSERT ON Sales.Currency TO UserB

--Question 3 
DENY DELETE ON Sales.Currency TO UserB

--Question 4
REVOKE INSERT ON Sales.Currency TO UserB

--Question 5





















