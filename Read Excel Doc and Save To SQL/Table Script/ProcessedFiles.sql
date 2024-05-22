 -- Create a table in the database to log processed files


CREATE TABLE ProcessedFiles (
    FileName NVARCHAR(255) PRIMARY KEY,
    LastModified DATETIME
);