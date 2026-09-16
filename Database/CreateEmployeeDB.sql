/* =============================================================
   EmployeeDB - SQL Server setup script
   Creates: EmployeeDB database, Employee & Education tables
            (1-to-many), WordTemplate table, dummy data and
            USP_GetEmployeeWithEducationFlat stored procedure.
   Target : DESKTOP-VFU2Q4K\SQLEXPRESS (Windows Authentication)
   ============================================================= */

IF DB_ID(N'EmployeeDB') IS NULL
    CREATE DATABASE EmployeeDB;
GO

USE EmployeeDB;
GO

/* ---------------- TABLES ---------------- */

IF OBJECT_ID(N'dbo.Employee', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Employee
    (
        Id            INT IDENTITY(1,1) PRIMARY KEY,
        FullName      NVARCHAR(200)  NOT NULL,
        DateOfBirth   DATETIME2      NULL,
        Gender        NVARCHAR(20)   NULL,
        Address       NVARCHAR(300)  NULL,
        City          NVARCHAR(100)  NULL,
        StateProvince NVARCHAR(100)  NULL,
        ZipPostalCode NVARCHAR(20)   NULL,
        Country       NVARCHAR(100)  NULL,
        PhoneNumber   NVARCHAR(50)   NULL,
        EmailAddress  NVARCHAR(200)  NOT NULL
    );
END
GO

IF OBJECT_ID(N'dbo.Education', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Education
    (
        Id            INT IDENTITY(1,1) PRIMARY KEY,
        EmployeeId    INT            NOT NULL,
        Degree        NVARCHAR(200)  NOT NULL,
        University    NVARCHAR(200)  NOT NULL,
        YearGraduated INT            NOT NULL,
        CONSTRAINT FK_Education_Employee
            FOREIGN KEY (EmployeeId) REFERENCES dbo.Employee(Id)
            ON DELETE CASCADE          -- one-to-many: delete employee -> delete education rows
    );
END
GO

IF OBJECT_ID(N'dbo.WordTemplate', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.WordTemplate
    (
        Id           INT IDENTITY(1,1) PRIMARY KEY,
        TemplateName NVARCHAR(255) NOT NULL,
        FilePath     NVARCHAR(500) NOT NULL,
        CONSTRAINT UQ_WordTemplate_TemplateName UNIQUE (TemplateName)
    );
END
GO

/* ---------------- DUMMY DATA ---------------- */

IF NOT EXISTS (SELECT 1 FROM dbo.Employee)
BEGIN
    SET IDENTITY_INSERT dbo.Employee ON;

    INSERT INTO dbo.Employee (Id, FullName, DateOfBirth, Gender, Address, City, StateProvince, ZipPostalCode, Country, PhoneNumber, EmailAddress)
    VALUES
    (1, 'Alice Wonderland',  '1998-05-15', 'Female', '123 Rabbit Hole, Apt 4B',     'Wonderland City', 'Fantasy',    '90210', 'Wonderland', '555-123-4567', 'alice.w@example.com'),
    (2, 'Bob The Builder',   '1990-11-20', 'Male',   '456 Construction Site',       'Toolsville',      'Workshire',  '10001', 'Buildland',  '555-987-6543', 'bob.b@example.com'),
    (3, 'Charlie Chaplin',   '1889-04-16', 'Male',   '789 Silent Film Studio',      'Hollywood',       'California', '90028', 'USA',        '555-111-2222', 'charlie.c@example.com'),
    (4, 'Ali Khan',          '1995-07-14', 'Male',   '12-A Gulberg III',            'Lahore',          'Punjab',     '54660', 'Pakistan',   '0300-1234567', 'ali.khan@example.com');

    SET IDENTITY_INSERT dbo.Employee OFF;
END
GO

IF NOT EXISTS (SELECT 1 FROM dbo.Education)
BEGIN
    SET IDENTITY_INSERT dbo.Education ON;

    INSERT INTO dbo.Education (Id, EmployeeId, Degree, University, YearGraduated)
    VALUES
    (1, 1, 'B.Sc. in Fictional Studies',    'Wonderland University',      2020),
    (2, 1, 'M.A. in Logic & Illogic',       'Looking Glass Institute',    2022),
    (3, 2, 'Associate in Construction Tech','Bob''s Community College',   2010),
    (4, 2, 'Certificate in Heavy Machinery','Toolsville Polytechnic',     2012),
    (5, 2, 'B.Tech in Civil Engineering',   'Buildland Tech University',  2014),
    (6, 3, 'Diploma in Physical Comedy',    'Slapstick Academy',          1908),
    (7, 4, 'BSCS (Bachelor of Computer Science)', 'Punjab University',    2017),
    (8, 4, 'MSCS (Master of Computer Science)',   'FAST-NU Lahore',       2019);
    -- Employee 4 has multiple education records (appended in the template table)
    -- Employee 3 intentionally has 1 record; employees with none also work (LEFT JOIN)

    SET IDENTITY_INSERT dbo.Education OFF;
END
GO

/* ---------------- STORED PROCEDURE ---------------- */
/* Returns ONE flat result set: employee columns repeated for each
   education record (one row per education). Employees without any
   education still return a single row with NULL education columns. */

IF OBJECT_ID(N'dbo.USP_GetEmployeeWithEducationFlat', N'P') IS NOT NULL
    DROP PROCEDURE dbo.USP_GetEmployeeWithEducationFlat;
GO

CREATE PROCEDURE dbo.USP_GetEmployeeWithEducationFlat
    @ApplicationId INT
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        e.Id            AS EmployeeId,
        e.FullName,
        e.DateOfBirth,
        e.Gender,
        e.Address,
        e.City,
        e.StateProvince,
        e.ZipPostalCode,
        e.Country,
        e.PhoneNumber,
        e.EmailAddress,
        ed.Id           AS EducationId,
        ed.Degree,
        ed.University,
        ed.YearGraduated
    FROM dbo.Employee e
    LEFT JOIN dbo.Education ed
        ON ed.EmployeeId = e.Id
    WHERE e.Id = @ApplicationId;
END
GO

/* ---------------- VERIFY ---------------- */
PRINT '--- Employees ---';
SELECT * FROM dbo.Employee;
PRINT '--- Education ---';
SELECT * FROM dbo.Education;
PRINT '--- WordTemplate ---';
SELECT * FROM dbo.WordTemplate;
PRINT '--- Proc test: employee 4 (multiple education rows) ---';
EXEC dbo.USP_GetEmployeeWithEducationFlat @ApplicationId = 4;
GO
