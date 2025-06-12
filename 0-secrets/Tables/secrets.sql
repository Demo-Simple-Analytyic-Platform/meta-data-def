CREATE TABLE [dbo].[secrets] (
    [nm_secret] NVARCHAR (128)                                      NULL,
    [ds_secret] NVARCHAR (999) MASKED WITH (FUNCTION = 'default()') NULL
);
